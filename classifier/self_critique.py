"""Self-critique: a second Claude call evaluates and adjusts the first-pass output."""
from __future__ import annotations

import json
import re
from typing import Any, Dict, List

import structlog
from anthropic import Anthropic
from tenacity import retry, stop_after_attempt, wait_exponential

from config import settings
from embeddings.pinecone_index import Match

log = structlog.get_logger(__name__)


CRITIQUE_SYSTEM = """You are an SAP support reviewer. You will receive:
1. The original SAP error.
2. Retrieved corpus entries (with corpus_ids).
3. A first-pass diagnostic with ranked proposals.

Your job: critique the first-pass output. Specifically check:
- Does each proposal directly address the error symptom?
- Are confidence scores well-calibrated given the corpus evidence?
- Are there contradictions between proposals and the corpus?
- Are there safer or less invasive alternatives that were missed?
- Are any proposals risky enough that they should be down-ranked or removed?

Respond with a single JSON object (no prose, no markdown):

{
  "critique_notes": "1-3 sentences summarizing your assessment",
  "adjusted_proposals": [
    {
      "rank": 1,
      "title": "...",
      "steps": [...],
      "model_self_confidence": 0.0-1.0,
      "supporting_corpus_ids": [...],
      "risks": [...],
      "critique_changed": true|false,
      "change_reason": "if changed, why"
    }
  ]
}

Rules:
- You may add, remove, re-rank, or rewrite proposals.
- Keep at least 1 proposal in the output.
- Lower confidences when evidence is genuinely weak. Do not be polite — be honest.
- supporting_corpus_ids must be a subset of provided corpus_ids.
"""


def _extract_json(text: str) -> Dict[str, Any]:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.MULTILINE)
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON in critique output: {text[:200]}")
    return json.loads(cleaned[start : end + 1])


def _format_matches(matches: List[Match]) -> str:
    blocks = []
    for m in matches:
        md = m.metadata
        blocks.append(
            f"[corpus_id={m.corpus_id} sim={m.score:.3f} accepted={md.get('accepted_flag')}] "
            f"{md.get('error_signature','')[:200]} → {md.get('solution_excerpt','')[:400]}"
        )
    return "\n".join(blocks)


class SelfCritique:
    def __init__(self) -> None:
        self.client = Anthropic(api_key=settings.anthropic_api_key)
        self.model = settings.claude_critique_model

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=2, min=2, max=30))
    def critique(
        self,
        *,
        error: Dict[str, Any],
        matches: List[Match],
        first_pass: Dict[str, Any],
    ) -> Dict[str, Any]:
        user_prompt = (
            f"=== Original error ===\n"
            f"Source: {error.get('SOURCE')} | Severity: {error.get('SEVERITY')}\n"
            f"Short text: {error.get('SHORT_TEXT','')}\n\n"
            f"=== Corpus ===\n{_format_matches(matches)}\n\n"
            f"=== First-pass output ===\n{json.dumps(first_pass, indent=2)}\n\n"
            "Critique and respond with the JSON object as specified."
        )
        resp = self.client.messages.create(
            model=self.model,
            max_tokens=2048,
            system=CRITIQUE_SYSTEM,
            messages=[{"role": "user", "content": user_prompt}],
        )
        text = "".join(b.text for b in resp.content if hasattr(b, "text"))
        try:
            return _extract_json(text)
        except Exception as e:
            log.warning("critique_json_parse_failed", error=str(e), preview=text[:300])
            raise
