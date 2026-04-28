"""First-pass Claude classifier.

Receives an error + retrieved corpus matches, returns ranked proposals with
self-confidence and the supporting corpus_ids per proposal.
"""
from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional

import structlog
from anthropic import Anthropic
from tenacity import retry, stop_after_attempt, wait_exponential

from config import settings
from embeddings.pinecone_index import Match

log = structlog.get_logger(__name__)


SYSTEM_PROMPT = """You are an expert SAP S/4HANA support engineer. You diagnose SAP runtime errors \
(ABAP dumps, application log errors, system log entries, update task failures, CCMS alerts, and \
job cancellations) and propose ranked, actionable solutions.

You will be given:
1. A single live SAP error with its source, transaction, program, and message text.
2. Up to 20 retrieved corpus entries from public SAP communities (each tagged with corpus_id).

You must respond with a single JSON object (no prose, no markdown) with this schema:

{
  "module": "BASIS|SD|MM|FI|CO|PP|HANA|BODS|BTP|OTHER",
  "root_cause_summary": "1-2 sentences explaining likely root cause",
  "proposals": [
    {
      "rank": 1,
      "title": "Short proposal title (≤120 chars)",
      "steps": ["concrete step 1", "concrete step 2", ...],
      "model_self_confidence": 0.0-1.0,
      "supporting_corpus_ids": ["corp-id-1", "corp-id-2"],
      "risks": ["any risks or caveats"]
    }
  ]
}

Rules:
- Provide 1-5 proposals, ranked by your assessment of likelihood × safety.
- model_self_confidence reflects YOUR honest belief (well-calibrated). Do not inflate.
- supporting_corpus_ids must be a subset of the corpus_ids you were given. Empty list is allowed.
- If the corpus contains contradictory advice, mention it in the relevant proposal's risks.
- Steps should be concrete (transaction codes, parameters, table fields) — not generic advice.
"""


def _format_matches(matches: List[Match]) -> str:
    blocks: List[str] = []
    for i, m in enumerate(matches):
        md = m.metadata
        blocks.append(
            f"[corpus_id={m.corpus_id} similarity={m.score:.3f} accepted={md.get('accepted_flag')}]\n"
            f"Signature: {md.get('error_signature','')[:300]}\n"
            f"Module: {md.get('system_module','')}\n"
            f"Source: {md.get('source_url','')}\n"
            f"Solution excerpt: {md.get('solution_excerpt','')[:600]}\n"
        )
    return "\n---\n".join(blocks)


def _format_error(error: Dict[str, Any]) -> str:
    parts = [
        f"Source: {error.get('SOURCE')}",
        f"Severity: {error.get('SEVERITY')}",
        f"Occurred at: {error.get('OCCURRED_AT')}",
        f"System: {error.get('SYSTEM_ID')}",
    ]
    if error.get("TRANSACTION"):
        parts.append(f"Transaction: {error['TRANSACTION']}")
    if error.get("PROGRAM"):
        parts.append(f"Program: {error['PROGRAM']}")
    if error.get("OBJECT") or error.get("SUB_OBJECT"):
        parts.append(f"Object: {error.get('OBJECT')}/{error.get('SUB_OBJECT')}")
    parts.append(f"Short text: {error.get('SHORT_TEXT','')}")
    if error.get("LONG_TEXT"):
        parts.append(f"Long text: {str(error['LONG_TEXT'])[:4000]}")
    return "\n".join(parts)


def _extract_json(text: str) -> Dict[str, Any]:
    """Strip ```json fences if present, then parse."""
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.MULTILINE)
    # Find the outermost { ... }
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON object found in model output: {text[:200]}")
    return json.loads(cleaned[start : end + 1])


class SinglePassClassifier:
    def __init__(self) -> None:
        self.client = Anthropic(api_key=settings.anthropic_api_key)
        self.model = settings.claude_model

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=2, min=2, max=30))
    def classify(self, *, error: Dict[str, Any], matches: List[Match], augment_uncertainty: bool = False) -> Dict[str, Any]:
        sys = SYSTEM_PROMPT
        if augment_uncertainty:
            sys += (
                "\n\nIMPORTANT (retry mode): The first attempt produced low-confidence output. "
                "Reconsider edge cases, look for contradictions in the corpus, and lower your "
                "confidence if evidence is genuinely weak. It is acceptable to return only one "
                "honest proposal with low confidence rather than fabricating alternatives."
            )

        user_prompt = (
            f"=== Live SAP error ===\n{_format_error(error)}\n\n"
            f"=== Retrieved corpus (most relevant first) ===\n{_format_matches(matches)}\n\n"
            "Respond with a single JSON object as specified."
        )

        log.debug("classifier_call", model=self.model, n_matches=len(matches))
        resp = self.client.messages.create(
            model=self.model,
            max_tokens=2048,
            system=sys,
            messages=[{"role": "user", "content": user_prompt}],
        )
        # Concatenate text blocks
        text = "".join(b.text for b in resp.content if hasattr(b, "text"))
        try:
            payload = _extract_json(text)
        except Exception as e:
            log.warning("classifier_json_parse_failed", error=str(e), preview=text[:300])
            raise
        return payload
