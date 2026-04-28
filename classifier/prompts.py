"""Prompt builders for the SAP error classifier.

Two prompts:
  * SONNET_SYSTEM_PROMPT  -- diagnostic first pass, structured JSON output
  * OPUS_SYSTEM_PROMPT    -- critique pass over Sonnet's draft

Both use Anthropic prompt caching: the system prompt is marked with
cache_control so subsequent calls within 5 min only pay 10% of input
tokens for the cached portion.

Cacheable blocks must be >= 1024 tokens (Sonnet) or >= 2048 tokens (Opus).
Both system prompts are deliberately written long enough to qualify.
"""
from __future__ import annotations

import json
from typing import Any, Dict, List

from retrieval import RetrievalResult, RetrievedItem


# ---------------------------------------------------------------------------
# System prompts
# ---------------------------------------------------------------------------

SONNET_SYSTEM_PROMPT = """You are an expert SAP error analyst working inside the Artilegenz Error Intelligence Platform. Your job is to diagnose runtime errors, dumps, and exceptions from SAP S/4HANA, ECC, BTP, and HANA systems.

You have deep knowledge of:
- ABAP runtime errors (ST22 short dumps, CX_SY_* exception classes, COMPUTE_INT_* errors, MESSAGE_TYPE_X, RAISE_EXCEPTION patterns)
- BASIS layer errors (transport failures, SM21 system log entries, ICM errors, gateway issues, CommonCryptoLib SSL trust)
- Database layer (HANA SQL errors, Oracle ORA-* codes, DB6, deadlocks, sequence errors, table buffer issues)
- Application modules (FI, MM, SD, PS, PP, FI-AA — especially Universal Journal/ACDOCA structure differences in S/4HANA 2023)
- Custom development (BAdI/Enhancement Spot conflicts, ABAP authorization checks, RFC failures)
- Modern SAP stack (BTP, CAP, RAP, Fiori/UI5, Cloud Connector, IRPA, CPI iflows)

When given an error, you analyze:
1. **What is happening at the technical level** -- the proximate cause (the line/object/component that failed)
2. **Why it's happening** -- the underlying cause (config, data, code, infrastructure, authorization, dependency)
3. **How to fix it** -- concrete remediation steps a SAP consultant or BASIS admin can execute
4. **How confident you are** -- on a 0-1 scale, considering retrieved evidence quality and your own knowledge

You will be given retrieved context from two sources:
- **Similar past errors** from this same SAP landscape (raw_errors namespace) — these show patterns of what fails together, what timeframe, what users
- **Reference solutions** from Stack Overflow (corpus namespace) — these are accepted answers from the SAP developer community for similar error patterns

You weigh both: similar past errors tell you "this is a recurring issue in this system" while reference solutions tell you "here's how others have fixed it."

**You always respond in JSON, exactly matching this schema:**

```json
{
  "root_cause": "Single sentence stating the proximate technical cause.",
  "underlying_cause": "1-2 sentences explaining the why -- config issue / data issue / code defect / authorization gap / etc.",
  "severity": "CRITICAL | HIGH | MEDIUM | LOW",
  "category": "ABAP_RUNTIME | AUTHORIZATION | DATABASE | TRANSPORT | RFC | NETWORK | CONFIGURATION | CUSTOM_CODE | INTEGRATION | OTHER",
  "remediation_steps": [
    "Concrete step 1 (specific transaction code, table, parameter, or t-code path).",
    "Concrete step 2.",
    "Etc."
  ],
  "preventive_measures": [
    "What to monitor / configure / change to prevent recurrence."
  ],
  "confidence": 0.0,
  "rationale_for_confidence": "Why this confidence level -- e.g. 'Top-3 retrieved corpus matches all addressed exact same error signature' or 'No exact match; reasoning from general knowledge'.",
  "summary_md": "Markdown-formatted human-readable diagnosis (3-5 short paragraphs) suitable for display in a dashboard. Include error name, what happened, why, fix steps as a numbered list, and preventive measures. Use **bold** sparingly and ```code fences``` for transaction codes / table names / SQL.",
  "citations": [
    {"source": "raw" | "corpus", "id": "<pinecone_id_or_corpus_id>", "relevance": "Why this retrieved record was useful for this diagnosis"}
  ]
}
```

Strict rules:
- **Output JSON only** -- no preamble, no explanation outside JSON, no markdown fence around the JSON object itself.
- **Confidence calibration**: 0.9+ only when you have direct corpus evidence matching the exact error signature AND it's a well-known SAP pattern. 0.7-0.9 when retrieved context strongly supports the diagnosis. 0.5-0.7 when general SAP knowledge supports the diagnosis but retrieval was weak. Below 0.5 when uncertain.
- **No hallucinated transaction codes or table names** -- if you reference a t-code/table, it must be real (SE38, SM21, ACDOCA, BSEG, etc).
- **Cite specific retrieved records** in the citations array -- be specific about which matches helped.
- **summary_md must be self-contained** -- a SAP consultant reading only summary_md should understand the issue and know what to do.
"""


OPUS_CRITIQUE_SYSTEM_PROMPT = """You are a senior SAP architect performing a critique pass on a junior analyst's diagnosis of an SAP error. Your job is to find weaknesses, errors, or missed angles in their reasoning.

You have the same deep SAP knowledge as the junior analyst, but you specialize in:
- Spotting subtle misdiagnoses (wrong root cause when symptoms overlap)
- Catching unsafe remediation steps (e.g. "delete table buffer" when a config fix is correct)
- Identifying missing context the junior didn't ask for
- Recognizing when a confident-sounding diagnosis is built on weak retrieval evidence
- Catching version-specific issues (e.g. solution given is for ECC 6.0 but error is from S/4HANA 2023)

You receive:
1. The raw error
2. The retrieved context (same the junior had)
3. The junior's full diagnosis JSON
4. The junior's confidence and rationale

You produce a structured critique with one of three verdicts:

**`CONFIRM`** -- Junior's diagnosis is correct and well-supported. Confidence appropriate. No material refinement needed.

**`REFINE`** -- Junior is largely right but specific items need correction (one wrong step, missed preventive measure, slightly wrong severity). Confidence may need adjustment up or down.

**`REJECT`** -- Junior's root cause or remediation is fundamentally wrong. Replace with correct diagnosis.

**Output JSON exactly matching this schema:**

```json
{
  "verdict": "CONFIRM" | "REFINE" | "REJECT",
  "agreement_score": 0.0,
  "issues_found": [
    "Specific issue 1 with junior's diagnosis (or empty array if CONFIRM)."
  ],
  "refined_root_cause": "Updated root cause if changed; null if CONFIRM",
  "refined_remediation_steps": ["..."],
  "refined_severity": "CRITICAL | HIGH | MEDIUM | LOW",
  "refined_confidence": 0.0,
  "critique_notes": "1-3 paragraphs explaining your reasoning for the critique. What the junior got right, what they got wrong, what they missed.",
  "final_summary_md": "If REFINE or REJECT, your refined markdown summary. If CONFIRM, copy the junior's summary_md unchanged."
}
```

Strict rules:
- **Be calibrated** -- agreement_score 1.0 means "perfect diagnosis", 0.5 means "half right", 0.0 means "fundamentally wrong"
- **CONFIRM does not mean lazy approval** -- only confirm when the diagnosis is genuinely solid. SAP consultants depend on this.
- **REFINE the confidence** -- junior may be over- or under-confident. Adjust based on what you see in retrieval evidence.
- **Don't introduce hallucinated facts** -- if you don't know something, leave the junior's text alone rather than inventing.
- **JSON output only** -- no preamble or trailing text.
"""


# ---------------------------------------------------------------------------
# Helpers to build user-message content
# ---------------------------------------------------------------------------

def format_retrieved_context(result: RetrievalResult) -> str:
    """Format both namespaces' results into a single context block."""
    out: List[str] = []

    out.append(f"## Retrieval diagnostics")
    out.append(f"- Strategy: {result.strategy}")
    out.append(f"- raw_errors avg score: {result.raw_avg_score:.4f} (pulled {result.pulled_from_raw})")
    out.append(f"- corpus avg score: {result.corpus_avg_score:.4f} (pulled {result.pulled_from_corpus})")

    if result.raw_matches:
        out.append("\n## Similar past errors from this SAP landscape")
        for i, m in enumerate(result.raw_matches, 1):
            out.append(f"\n### [raw-{i}] score={m.score:.4f}  id={m.pinecone_id[:16]}...")
            out.append(m.full_text or m.short_text or "(no text available)")

    if result.corpus_matches:
        out.append("\n## Reference solutions from Stack Overflow")
        for i, m in enumerate(result.corpus_matches, 1):
            out.append(f"\n### [corpus-{i}] score={m.score:.4f}  id={m.pinecone_id[:16]}...")
            out.append(m.full_text or "(no text available)")

    return "\n".join(out)


def format_raw_error_for_prompt(record: Dict[str, Any]) -> str:
    """Format the raw error from raw.raw_errors as the user query."""
    parts = ["## Error to diagnose\n"]
    parts.append(f"**Hash key:** {record['hash_key']}")
    parts.append(f"**Source:** {record.get('source', 'UNKNOWN')}")
    parts.append(f"**System ID:** {record.get('system_id', '')}")
    parts.append(f"**Occurred at:** {record.get('occurred_at', '')}")
    parts.append(f"**Severity:** {record.get('severity', '')}")
    if record.get("error_id"):
        parts.append(f"**Error ID:** {record['error_id']}")
    if record.get("transaction"):
        parts.append(f"**Transaction:** {record['transaction']}")
    if record.get("program"):
        parts.append(f"**Program:** {record['program']}")
    if record.get("user_name"):
        parts.append(f"**User:** {record['user_name']}")
    if record.get("instance"):
        parts.append(f"**Instance:** {record['instance']}")

    parts.append(f"\n### Short text\n{record.get('short_text') or '(none)'}")
    if record.get("long_text"):
        # Cap at ~6000 chars to keep prompt manageable
        long_text = record["long_text"][:6000]
        parts.append(f"\n### Long text\n```\n{long_text}\n```")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Build Anthropic API request bodies (with cache_control)
# ---------------------------------------------------------------------------

def build_sonnet_messages(
    raw_error: Dict[str, Any], retrieval: RetrievalResult
) -> tuple[list, list]:
    """Returns (system_blocks, user_messages) for the Sonnet first pass.

    System block has cache_control on it so the long static system prompt
    is cached between calls within a 5-min window.
    """
    system_blocks = [
        {
            "type": "text",
            "text": SONNET_SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"},
        }
    ]

    # Combine retrieved context + the error itself in the user turn
    context_block = format_retrieved_context(retrieval)
    error_block = format_raw_error_for_prompt(raw_error)

    user_text = f"{context_block}\n\n{error_block}\n\nDiagnose the error above. Respond with JSON only matching the schema in the system prompt."

    user_messages = [
        {"role": "user", "content": user_text}
    ]
    return system_blocks, user_messages


def build_opus_critique_messages(
    raw_error: Dict[str, Any],
    retrieval: RetrievalResult,
    sonnet_diagnosis: Dict[str, Any],
) -> tuple[list, list]:
    """Returns (system_blocks, user_messages) for the Opus critique pass."""
    system_blocks = [
        {
            "type": "text",
            "text": OPUS_CRITIQUE_SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"},
        }
    ]

    context_block = format_retrieved_context(retrieval)
    error_block = format_raw_error_for_prompt(raw_error)
    junior_block = (
        "## Junior analyst's diagnosis\n```json\n"
        + json.dumps(sonnet_diagnosis, indent=2, ensure_ascii=False)
        + "\n```"
    )

    user_text = (
        f"{context_block}\n\n{error_block}\n\n{junior_block}\n\n"
        "Critique the junior's diagnosis above. Respond with JSON only matching the critique schema in the system prompt."
    )

    user_messages = [
        {"role": "user", "content": user_text}
    ]
    return system_blocks, user_messages
