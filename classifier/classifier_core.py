"""Classifier core: Sonnet first pass + Opus critique with prompt caching.

Public entrypoint:
    classify(hash_key) -> classification_id
        Reads raw error from Postgres, retrieves context, runs both LLM
        passes, writes intel.classifications row, returns classification_id.

Designed to be called from CLI, batch worker, and Windows service.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Optional

import psycopg2
import psycopg2.extras
import structlog


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(format="%(message)s", stream=sys.stdout, level=logging.INFO)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.dev.ConsoleRenderer(),
    ],
)
log = structlog.get_logger("classifier")


# ---------------------------------------------------------------------------
# Config -- bumped defaults so JSON output never gets truncated
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY  = os.getenv("ANTHROPIC_API_KEY", "")
SONNET_MODEL       = os.getenv("SONNET_MODEL", "claude-sonnet-4-6")
OPUS_MODEL         = os.getenv("OPUS_MODEL", "claude-opus-4-7")
SONNET_MAX_TOKENS  = int(os.getenv("SONNET_MAX_TOKENS", "4000"))
OPUS_MAX_TOKENS    = int(os.getenv("OPUS_MAX_TOKENS", "5000"))
TOTAL_K            = int(os.getenv("RETRIEVAL_TOTAL_K", "10"))

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Postgres helpers
# ---------------------------------------------------------------------------
def fetch_raw_error(hash_key: str) -> Optional[Dict[str, Any]]:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT hash_key, source, system_id, occurred_at, severity,
                          short_text, long_text, error_id, transaction, program,
                          user_name, instance, work_process, object, sub_object,
                          job_name
                   FROM raw.raw_errors WHERE hash_key = %s""",
                (hash_key,),
            )
            row = cur.fetchone()
            return dict(row) if row else None


INSERT_CLASSIFICATION_SQL = """
INSERT INTO intel.classifications (
  classification_id, error_hash_key, created_at,
  top_proposal_title, composite_confidence, badge, retried,
  proposals, citations, critique_notes, model_version,
  sonnet_response, opus_critique, adaptive_split,
  sonnet_tokens, opus_tokens, status, error_message, retry_count, summary_md
)
VALUES (
  %(classification_id)s, %(error_hash_key)s, CURRENT_TIMESTAMP,
  %(top_proposal_title)s, %(composite_confidence)s, %(badge)s, %(retried)s,
  %(proposals)s, %(citations)s, %(critique_notes)s, %(model_version)s,
  %(sonnet_response)s, %(opus_critique)s, %(adaptive_split)s,
  %(sonnet_tokens)s, %(opus_tokens)s, %(status)s, %(error_message)s, %(retry_count)s,
  %(summary_md)s
);
"""


def write_classification(row: Dict[str, Any]) -> None:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute(INSERT_CLASSIFICATION_SQL, row)
            conn.commit()


# ---------------------------------------------------------------------------
# Anthropic client
# ---------------------------------------------------------------------------
def _build_anthropic():
    import anthropic
    return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def _call_claude(
    client,
    model: str,
    system_blocks: list,
    user_messages: list,
    max_tokens: int,
) -> tuple[str, Dict[str, Any]]:
    """Returns (text_response, token_usage_dict)."""
    resp = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system_blocks,
        messages=user_messages,
    )

    text_parts = []
    for block in resp.content:
        if hasattr(block, "text"):
            text_parts.append(block.text)
    text = "".join(text_parts)

    usage = resp.usage
    token_info = {
        "input_tokens": getattr(usage, "input_tokens", 0),
        "output_tokens": getattr(usage, "output_tokens", 0),
        "cache_creation_input_tokens": getattr(usage, "cache_creation_input_tokens", 0),
        "cache_read_input_tokens": getattr(usage, "cache_read_input_tokens", 0),
        "stop_reason": getattr(resp, "stop_reason", None),
    }
    return text, token_info


# ---------------------------------------------------------------------------
# Robust JSON parser
# ---------------------------------------------------------------------------
def _parse_json_response(text: str) -> Dict[str, Any]:
    """Robustly extract a JSON object from Claude's response.

    Handles:
      * ```json fences
      * Leading/trailing prose
      * (Best-effort) truncated JSON via bracket/quote repair
    """
    text = text.strip()

    # Strip ```json ... ``` fences if present
    if text.startswith("```"):
        first_nl = text.find("\n")
        if first_nl != -1:
            text = text[first_nl + 1:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()

    # Locate first '{'
    first_brace = text.find("{")
    if first_brace == -1:
        raise json.JSONDecodeError("No JSON object found", text, 0)

    candidate = text[first_brace:]

    # Strategy 1: parse from first brace as-is
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass

    # Strategy 2: trim to last '}' and try again
    last_brace = candidate.rfind("}")
    if last_brace != -1:
        try:
            return json.loads(candidate[:last_brace + 1])
        except json.JSONDecodeError:
            pass

    # Strategy 3: repair truncation -- close unterminated strings + brackets
    repaired = _repair_truncated_json(candidate)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError as exc:
        # Log full failure context so the operator can see what Claude returned
        log.error(
            "json_parse_unrepairable",
            text_len=len(candidate),
            text_head=candidate[:400],
            text_tail=candidate[-400:],
            err=str(exc),
        )
        raise


def _repair_truncated_json(text: str) -> str:
    """Best-effort repair of mid-string truncation.

    Walks the string char-by-char tracking quote state and bracket depth,
    then closes whatever's open. Imperfect (won't fix corrupted middle of
    a value) but recovers most truncation-at-end cases.
    """
    in_string = False
    escape_next = False
    brace_depth = 0
    bracket_depth = 0

    for ch in text:
        if escape_next:
            escape_next = False
            continue
        if ch == "\\" and in_string:
            escape_next = True
            continue
        if ch == '"' and not escape_next:
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "{":
            brace_depth += 1
        elif ch == "}":
            brace_depth -= 1
        elif ch == "[":
            bracket_depth += 1
        elif ch == "]":
            bracket_depth -= 1

    repaired = text
    if in_string:
        # If the truncation happened inside a string, close the string.
        # Strip trailing backslash that would otherwise escape our closing quote.
        if repaired.endswith("\\"):
            repaired = repaired[:-1]
        repaired += '"'

    # Close any open brackets and braces in correct nesting order
    repaired += "]" * max(bracket_depth, 0)
    repaired += "}" * max(brace_depth, 0)
    return repaired


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
@dataclass
class ClassificationResult:
    classification_id: str
    sonnet_diagnosis: Dict[str, Any]
    opus_critique: Dict[str, Any]
    final_confidence: float
    final_summary_md: str
    final_root_cause: str
    final_severity: str
    final_remediation: list
    badge: str
    adaptive_split: Dict[str, Any]
    sonnet_tokens: Dict[str, int]
    opus_tokens: Dict[str, int]


def _make_classification_id(hash_key: str) -> str:
    raw = f"{hash_key}:{datetime.utcnow().isoformat()}"
    return hashlib.sha256(raw.encode()).hexdigest()[:32]


def _badge_for_confidence(confidence: float) -> str:
    if confidence >= 0.85:
        return "HIGH_CONFIDENCE"
    if confidence >= 0.65:
        return "MEDIUM_CONFIDENCE"
    if confidence >= 0.45:
        return "LOW_CONFIDENCE"
    return "UNCERTAIN"


def classify(hash_key: str, retry_count: int = 0) -> ClassificationResult:
    """Run full classification pipeline on a single error."""
    if not ANTHROPIC_API_KEY:
        raise SystemExit("Set ANTHROPIC_API_KEY env var")

    log.info(
        "classify_starting",
        hash_key=hash_key[:16] + "...",
        sonnet_max=SONNET_MAX_TOKENS,
        opus_max=OPUS_MAX_TOKENS,
    )

    # 1. Fetch error
    raw_error = fetch_raw_error(hash_key)
    if not raw_error:
        raise ValueError(f"No raw error found for hash_key={hash_key}")

    # 2. Retrieve context
    from retrieval import adaptive_retrieve
    query_text = f"{raw_error.get('short_text') or ''}\n{(raw_error.get('long_text') or '')[:6000]}"
    retrieval = adaptive_retrieve(query_text, total_k=TOTAL_K)
    log.info(
        "retrieval_done",
        strategy=retrieval.strategy,
        raw_avg=round(retrieval.raw_avg_score, 3),
        corpus_avg=round(retrieval.corpus_avg_score, 3),
        n_raw=retrieval.pulled_from_raw,
        n_corpus=retrieval.pulled_from_corpus,
    )

    # 3. Sonnet first pass
    from prompts import build_sonnet_messages, build_opus_critique_messages
    client = _build_anthropic()

    sys_blocks, user_msgs = build_sonnet_messages(raw_error, retrieval)
    sonnet_text, sonnet_tokens = _call_claude(
        client, SONNET_MODEL, sys_blocks, user_msgs, SONNET_MAX_TOKENS
    )
    log.info("sonnet_done", **sonnet_tokens)

    # Warn if Sonnet hit the cap (truncation risk)
    if sonnet_tokens.get("stop_reason") == "max_tokens":
        log.warning(
            "sonnet_truncated_at_cap",
            output_tokens=sonnet_tokens.get("output_tokens"),
            cap=SONNET_MAX_TOKENS,
            hint="Consider raising SONNET_MAX_TOKENS",
        )

    try:
        sonnet_diagnosis = _parse_json_response(sonnet_text)
    except json.JSONDecodeError as exc:
        log.error("sonnet_json_parse_failed", err=str(exc), text_head=sonnet_text[:500])
        raise

    # 4. Opus critique pass
    sys_blocks2, user_msgs2 = build_opus_critique_messages(
        raw_error, retrieval, sonnet_diagnosis
    )
    opus_text, opus_tokens = _call_claude(
        client, OPUS_MODEL, sys_blocks2, user_msgs2, OPUS_MAX_TOKENS
    )
    log.info("opus_done", **opus_tokens)

    if opus_tokens.get("stop_reason") == "max_tokens":
        log.warning(
            "opus_truncated_at_cap",
            output_tokens=opus_tokens.get("output_tokens"),
            cap=OPUS_MAX_TOKENS,
            hint="Consider raising OPUS_MAX_TOKENS",
        )

    try:
        opus_critique = _parse_json_response(opus_text)
    except json.JSONDecodeError as exc:
        log.error("opus_json_parse_failed", err=str(exc), text_head=opus_text[:500])
        raise

    # 5. Merge: prefer Opus refinements when present
    verdict = opus_critique.get("verdict", "CONFIRM")
    if verdict in ("REJECT", "REFINE"):
        final_root_cause = opus_critique.get("refined_root_cause") or sonnet_diagnosis.get("root_cause", "")
        final_remediation = opus_critique.get("refined_remediation_steps") or sonnet_diagnosis.get("remediation_steps", [])
        final_severity = opus_critique.get("refined_severity") or sonnet_diagnosis.get("severity", "MEDIUM")
        final_confidence = float(opus_critique.get("refined_confidence") or sonnet_diagnosis.get("confidence", 0.5))
        final_summary = opus_critique.get("final_summary_md") or sonnet_diagnosis.get("summary_md", "")
    else:
        final_root_cause = sonnet_diagnosis.get("root_cause", "")
        final_remediation = sonnet_diagnosis.get("remediation_steps", [])
        final_severity = sonnet_diagnosis.get("severity", "MEDIUM")
        final_confidence = float(sonnet_diagnosis.get("confidence", 0.5))
        final_summary = sonnet_diagnosis.get("summary_md", "")

    badge = _badge_for_confidence(final_confidence)

    # 6. Build classification row
    classification_id = _make_classification_id(hash_key)
    row = {
        "classification_id": classification_id,
        "error_hash_key": hash_key,
        "top_proposal_title": (final_root_cause or "")[:1024],
        "composite_confidence": round(final_confidence, 4),
        "badge": badge,
        "retried": retry_count > 0,
        "proposals": json.dumps({
            "root_cause": final_root_cause,
            "underlying_cause": sonnet_diagnosis.get("underlying_cause"),
            "severity": final_severity,
            "category": sonnet_diagnosis.get("category"),
            "remediation_steps": final_remediation,
            "preventive_measures": sonnet_diagnosis.get("preventive_measures") or [],
        }, ensure_ascii=False),
        "citations": json.dumps(sonnet_diagnosis.get("citations") or [], ensure_ascii=False),
        "critique_notes": opus_critique.get("critique_notes", ""),
        "model_version": f"sonnet={SONNET_MODEL}; opus={OPUS_MODEL}",
        "sonnet_response": json.dumps(sonnet_diagnosis, ensure_ascii=False),
        "opus_critique": json.dumps(opus_critique, ensure_ascii=False),
        "adaptive_split": json.dumps(retrieval.adaptive_split(), ensure_ascii=False),
        "sonnet_tokens": json.dumps(sonnet_tokens),
        "opus_tokens": json.dumps(opus_tokens),
        "status": "classified",
        "error_message": None,
        "retry_count": retry_count,
        "summary_md": final_summary,
    }
    write_classification(row)

    log.info(
        "classify_complete",
        classification_id=classification_id,
        confidence=round(final_confidence, 3),
        badge=badge,
        verdict=verdict,
    )

    return ClassificationResult(
        classification_id=classification_id,
        sonnet_diagnosis=sonnet_diagnosis,
        opus_critique=opus_critique,
        final_confidence=final_confidence,
        final_summary_md=final_summary,
        final_root_cause=final_root_cause,
        final_severity=final_severity,
        final_remediation=final_remediation,
        badge=badge,
        adaptive_split=retrieval.adaptive_split(),
        sonnet_tokens=sonnet_tokens,
        opus_tokens=opus_tokens,
    )


def classify_with_failure_logging(hash_key: str) -> Optional[ClassificationResult]:
    """Wrapper that catches errors and writes a 'failed' row so retry logic
    can see them. Used by the batch worker / service."""
    retry_count = 0
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """SELECT COALESCE(MAX(retry_count), 0)
                   FROM intel.classifications
                   WHERE error_hash_key = %s AND status = 'failed'""",
                (hash_key,),
            )
            row = cur.fetchone()
            if row and row[0] is not None:
                retry_count = row[0] + 1

    try:
        return classify(hash_key, retry_count=retry_count)
    except Exception as exc:  # noqa: BLE001
        log.exception("classify_failed", hash_key=hash_key[:16] + "...", err=str(exc))
        cid = _make_classification_id(hash_key)
        write_classification({
            "classification_id": cid,
            "error_hash_key": hash_key,
            "top_proposal_title": None,
            "composite_confidence": None,
            "badge": "FAILED",
            "retried": retry_count > 0,
            "proposals": None,
            "citations": None,
            "critique_notes": None,
            "model_version": f"sonnet={SONNET_MODEL}; opus={OPUS_MODEL}",
            "sonnet_response": None,
            "opus_critique": None,
            "adaptive_split": None,
            "sonnet_tokens": None,
            "opus_tokens": None,
            "status": "failed",
            "error_message": str(exc)[:5000],
            "retry_count": retry_count,
            "summary_md": None,
        })
        return None
