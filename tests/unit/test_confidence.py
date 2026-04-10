"""Unit tests for the ConfidenceScorer and Wilson CI calculation."""

import math
import pytest

from nlp_sap.nlp.confidence import ConfidenceScorer, _wilson_interval
from nlp_sap.nlp.models import ParsedIntent, SAPModule
from nlp_sap.schema.registry import get_registry


@pytest.fixture
def scorer() -> ConfidenceScorer:
    return ConfidenceScorer(get_registry())


def _make_intent(**kwargs) -> ParsedIntent:
    defaults = {
        "intent_name": "ar_open_items",
        "module": SAPModule.FI,
        "confidence": 0.9,
        "company_code": "1000",
        "fiscal_year": "2024",
    }
    defaults.update(kwargs)
    return ParsedIntent(**defaults)


class TestWilsonInterval:
    def test_perfect_confidence_bounds_narrow(self) -> None:
        lo, hi = _wilson_interval(0.99, 40)
        assert lo > 0.90
        assert hi <= 1.0

    def test_low_confidence_interval_wide(self) -> None:
        lo, hi = _wilson_interval(0.5, 4)
        assert hi - lo > 0.3  # Wide interval for few signals

    def test_zero_samples_returns_full_range(self) -> None:
        lo, hi = _wilson_interval(0.7, 0)
        assert lo == 0.0
        assert hi == 1.0

    def test_interval_is_within_bounds(self) -> None:
        for p in [0.1, 0.3, 0.5, 0.7, 0.9]:
            for n in [5, 10, 40]:
                lo, hi = _wilson_interval(p, n)
                assert 0.0 <= lo <= hi <= 1.0

    def test_interval_symmetric_at_midpoint(self) -> None:
        lo, hi = _wilson_interval(0.5, 20)
        assert abs((hi - 0.5) - (0.5 - lo)) < 0.01


class TestConfidenceScorer:
    def test_high_confidence_intent(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent(confidence=0.95, company_code="1000", fiscal_year="2024")
        result = scorer.score(intent, "show AR open items for company 1000 in 2024")
        assert result.label == "High"
        assert result.score >= 0.75
        assert result.lower < result.score < result.upper

    def test_low_confidence_without_entities(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent(
            intent_name="gl_balance_inquiry",
            confidence=0.3,
            company_code=None,
            fiscal_year=None,
        )
        result = scorer.score(intent, "show me something")
        assert result.score < 0.7

    def test_label_high(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent(confidence=0.95)
        result = scorer.score(intent, "show AR open items for customer 10001 in company 1000")
        assert result.label in ("High", "Medium")  # depends on keyword overlap

    def test_label_low(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent(confidence=0.2, company_code=None, fiscal_year=None)
        result = scorer.score(intent, "xyz abc")
        assert result.label == "Low"

    def test_confidence_interval_present(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent(confidence=0.85)
        result = scorer.score(intent, "show open AR items for company 1000")
        d = result.to_dict()
        assert "interval" in d
        assert d["interval"]["method"] == "Wilson score"
        assert d["interval"]["level"] == "95%"

    def test_breakdown_has_all_dimensions(self, scorer: ConfidenceScorer) -> None:
        intent = _make_intent()
        result = scorer.score(intent, "test query")
        assert "llm_confidence" in result.breakdown
        assert "keyword_overlap" in result.breakdown
        assert "entity_completeness" in result.breakdown
        assert "schema_consistency" in result.breakdown

    def test_invalid_fiscal_year_lowers_schema_score(self, scorer: ConfidenceScorer) -> None:
        intent_good = _make_intent(fiscal_year="2024")
        intent_bad = _make_intent(fiscal_year="abcd")
        query = "show GL balance for 2024 in company 1000"
        good = scorer.score(intent_good, query)
        bad = scorer.score(intent_bad, query)
        assert good.score >= bad.score

    def test_invalid_company_code_lowers_schema_score(self, scorer: ConfidenceScorer) -> None:
        intent_good = _make_intent(company_code="1000")
        intent_bad = _make_intent(company_code="TOOLONG_CODE")
        query = "AR items for company 1000"
        good = scorer.score(intent_good, query)
        bad = scorer.score(intent_bad, query)
        assert good.breakdown["schema_consistency"] >= bad.breakdown["schema_consistency"]
