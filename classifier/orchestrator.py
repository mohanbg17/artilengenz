"""End-to-end orchestrator for a single error.

Flow:
  1. Embed the live error (Voyage query embedding)
  2. Retrieve top-K from Pinecone
  3. SinglePassClassifier produces ranked proposals
  4. SelfCritique adjusts proposals
  5. Composite confidence scored per proposal
  6. If top proposal composite < threshold, retry once with K=top_k_retry and uncertainty-augmented prompt
  7. Persist to Snowflake CLASSIFICATIONS
"""
from __future__ import annotations

import hashlib
import logging
from datetime import datetime
from typing import Any, Dict, List

import structlog

from config import settings
from embeddings.pinecone_index import Match, get_index
from embeddings.voyage_client import build_query_embedding_text, get_voyage
from scraper.sf_writer import write_classification

from .confidence import score_proposal
from .self_critique import SelfCritique
from .single_pass import SinglePassClassifier

log = structlog.get_logger(__name__)


def _classification_id(error_hash_key: str) -> str:
    h = hashlib.sha256(f"{error_hash_key}|{datetime.utcnow().isoformat()}".encode()).hexdigest()
    return f"cls-{h[:24]}"


def _score_all(proposals: List[Dict[str, Any]], matches: List[Match]) -> List[Dict[str, Any]]:
    """Attach composite scores to each proposal in-place; return enriched list."""
    enriched: List[Dict[str, Any]] = []
    for p in proposals:
        score = score_proposal(
            supporting_corpus_ids=p.get("supporting_corpus_ids", []),
            matches=matches,
            model_self_confidence=float(p.get("model_self_confidence", 0.5)),
        )
        enriched.append({
            **p,
            "vector_similarity": score.vector_similarity,
            "corpus_corroboration": score.corpus_corroboration,
            "composite_confidence": score.composite,
            "badge": score.badge,
        })
    # Re-rank by composite
    enriched.sort(key=lambda p: p["composite_confidence"], reverse=True)
    for i, p in enumerate(enriched, start=1):
        p["rank"] = i
    return enriched


class Orchestrator:
    def __init__(self) -> None:
        self.voyage = get_voyage()
        self.index = get_index()
        self.first_pass = SinglePassClassifier()
        self.critic = SelfCritique()

    async def run(self, error: Dict[str, Any]) -> Dict[str, Any]:
        # 1. Embed
        query_text = build_query_embedding_text(
            source=error.get("SOURCE", ""),
            short_text=error.get("SHORT_TEXT", ""),
            long_text=error.get("LONG_TEXT"),
            transaction=error.get("TRANSACTION"),
            program=error.get("PROGRAM"),
        )
        qvec = await self.voyage.embed_query_async(query_text)

        # 2. Retrieve
        matches = self.index.query(vector=qvec, top_k=settings.top_k_default)
        log.info("orchestrator_retrieved", n=len(matches), error=error.get("HASH_KEY"))

        # 3. First pass
        first = self.first_pass.classify(error=error, matches=matches, augment_uncertainty=False)

        # 4. Critique
        critique = self.critic.critique(error=error, matches=matches, first_pass=first)
        proposals = critique.get("adjusted_proposals", first.get("proposals", []))
        critique_notes = critique.get("critique_notes", "")

        # 5. Score
        scored = _score_all(proposals, matches)
        retried = False

        # 6. Retry if uncertain
        if scored and scored[0]["composite_confidence"] < settings.confidence_threshold:
            log.info(
                "orchestrator_retry",
                first_top_conf=scored[0]["composite_confidence"],
                threshold=settings.confidence_threshold,
            )
            matches = self.index.query(vector=qvec, top_k=settings.top_k_retry)
            first2 = self.first_pass.classify(error=error, matches=matches, augment_uncertainty=True)
            critique2 = self.critic.critique(error=error, matches=matches, first_pass=first2)
            proposals2 = critique2.get("adjusted_proposals", first2.get("proposals", []))
            critique_notes = critique2.get("critique_notes", "") + " [after retry]"
            scored = _score_all(proposals2, matches)
            retried = True

        # 7. Build final payload
        top = scored[0] if scored else {"title": "(no proposal)", "composite_confidence": 0.0, "badge": "UNCERTAIN"}
        citations = list({m.metadata.get("source_url") for m in matches if m.metadata.get("source_url")})
        cls_id = _classification_id(error["HASH_KEY"])
        payload = {
            "classification_id": cls_id,
            "error_hash_key": error["HASH_KEY"],
            "created_at": datetime.utcnow(),
            "top_proposal_title": top["title"][:1024],
            "composite_confidence": top["composite_confidence"],
            "badge": top["badge"],
            "retried": retried,
            "proposals": scored,
            "citations": citations,
            "critique_notes": critique_notes[:8000],
            "model_version": settings.claude_model,
            "module": first.get("module"),
            "root_cause_summary": first.get("root_cause_summary"),
        }

        # 8. Persist
        try:
            write_classification(payload)
        except Exception:  # noqa: BLE001
            log.exception("classification_persist_failed", error_hash=error["HASH_KEY"])

        log.info(
            "orchestrator_done",
            error=error["HASH_KEY"],
            top_conf=top["composite_confidence"],
            badge=top["badge"],
            retried=retried,
        )
        return payload
