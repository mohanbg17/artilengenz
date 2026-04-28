"""Composite confidence scoring.

composite = 0.45 * vector_similarity + 0.45 * model_self_confidence + 0.10 * corpus_corroboration

* vector_similarity: average cosine of top-3 retrieved corpus entries that align
  with the proposal (judged by Claude in the structured output via `supporting_corpus_ids`).
* model_self_confidence: Claude's own 0-1 estimate.
* corpus_corroboration: distinct supporting corpus entries normalized to [0,1] (3+ = 1.0).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List

from embeddings.pinecone_index import Match


W_VEC = 0.45
W_MODEL = 0.45
W_CORPUS = 0.10


@dataclass
class ProposalScore:
    vector_similarity: float
    model_self_confidence: float
    corpus_corroboration: float
    composite: float
    badge: str


def _badge(composite: float) -> str:
    if composite >= 0.85:
        return "HIGH"
    if composite >= 0.70:
        return "MEDIUM"
    return "UNCERTAIN"


def score_proposal(
    *,
    supporting_corpus_ids: List[str],
    matches: List[Match],
    model_self_confidence: float,
) -> ProposalScore:
    """Compute a single proposal's composite confidence."""
    by_id: Dict[str, Match] = {m.corpus_id: m for m in matches}

    # Filter to those that actually exist in retrieved set
    supporting = [by_id[c] for c in supporting_corpus_ids if c in by_id]

    if supporting:
        # Average top-3 cosine of supporting matches
        top3 = sorted(supporting, key=lambda m: m.score, reverse=True)[:3]
        vec_sim = sum(m.score for m in top3) / len(top3)
    else:
        # No supporting evidence → fall back to top-1 retrieved as ceiling
        vec_sim = matches[0].score if matches else 0.0

    distinct = len(set(supporting_corpus_ids) & set(by_id.keys()))
    corpus_corroboration = min(distinct / 3.0, 1.0)

    model_self_confidence = max(0.0, min(1.0, model_self_confidence))

    composite = (
        W_VEC * vec_sim
        + W_MODEL * model_self_confidence
        + W_CORPUS * corpus_corroboration
    )
    composite = max(0.0, min(1.0, composite))

    return ProposalScore(
        vector_similarity=round(vec_sim, 4),
        model_self_confidence=round(model_self_confidence, 4),
        corpus_corroboration=round(corpus_corroboration, 4),
        composite=round(composite, 4),
        badge=_badge(composite),
    )
