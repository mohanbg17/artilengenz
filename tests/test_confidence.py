"""Sanity tests for the composite confidence formula."""
from dataclasses import dataclass
from typing import Any, Dict

from classifier.confidence import score_proposal, _badge


@dataclass
class FakeMatch:
    pinecone_id: str
    corpus_id: str
    score: float
    hybrid_score: float
    metadata: Dict[str, Any]


def test_strong_signal_high_badge():
    matches = [
        FakeMatch("a", "c1", 0.95, 0.95, {}),
        FakeMatch("b", "c2", 0.92, 0.92, {}),
        FakeMatch("c", "c3", 0.90, 0.90, {}),
    ]
    s = score_proposal(supporting_corpus_ids=["c1", "c2", "c3"], matches=matches, model_self_confidence=0.95)
    assert s.composite > 0.85
    assert s.badge == "HIGH"


def test_weak_signal_uncertain_badge():
    matches = [
        FakeMatch("a", "c1", 0.45, 0.45, {}),
        FakeMatch("b", "c2", 0.40, 0.40, {}),
    ]
    s = score_proposal(supporting_corpus_ids=["c1"], matches=matches, model_self_confidence=0.40)
    assert s.composite < 0.70
    assert s.badge == "UNCERTAIN"


def test_no_supporting_falls_back_to_top_match():
    matches = [
        FakeMatch("a", "c1", 0.80, 0.80, {}),
    ]
    s = score_proposal(supporting_corpus_ids=[], matches=matches, model_self_confidence=0.70)
    # With 0 supporting → corpus_corroboration = 0
    # vec = 0.80 (top), model = 0.70, corp = 0
    # composite = 0.45*0.80 + 0.45*0.70 + 0.10*0 = 0.36 + 0.315 = 0.675
    assert abs(s.composite - 0.675) < 1e-3
    assert s.badge == "UNCERTAIN"


def test_badge_thresholds():
    assert _badge(0.90) == "HIGH"
    assert _badge(0.85) == "HIGH"
    assert _badge(0.84) == "MEDIUM"
    assert _badge(0.70) == "MEDIUM"
    assert _badge(0.69) == "UNCERTAIN"


def test_corpus_corroboration_caps_at_one():
    matches = [FakeMatch(f"i{i}", f"c{i}", 0.85, 0.85, {}) for i in range(6)]
    s = score_proposal(
        supporting_corpus_ids=[f"c{i}" for i in range(6)],
        matches=matches,
        model_self_confidence=0.80,
    )
    assert s.corpus_corroboration == 1.0
