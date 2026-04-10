"""Confidence Scoring Module.

Produces a calibrated confidence score and Wilson-score confidence interval
for each NLP classification, answering: "How sure is the system about this
intent, and what is the statistical uncertainty of that estimate?"

## Theory

### Base confidence
The raw confidence is a blend of:
  - LLM self-reported probability (0–1)
  - Keyword overlap score (0–1)
  - Entity extraction completeness (were required filters found?)
  - Schema consistency (do the extracted entities match the intent's expected fields?)

### Confidence Interval
We model the classification as a Bernoulli trial and apply the Wilson score interval:
  p̂ = observed proportion (confidence)
  n = effective sample size (based on evidence signals)

  Wilson lower = (p̂ + z²/2n - z√(p̂(1-p̂)/n + z²/4n²)) / (1 + z²/n)
  Wilson upper = (p̂ + z²/2n + z√(p̂(1-p̂)/n + z²/4n²)) / (1 + z²/n)

We use z = 1.96 (95% CI).

The effective sample size n scales with the number of distinct evidence signals
the system observed (LLM vote, keyword hits, entity hits, schema consistency).

### Why Wilson over Wald?
The Wald interval p̂ ± z√(p̂(1-p̂)/n) fails near 0 and 1 (the most common
scenario in high-confidence NLP). Wilson is better-calibrated at extremes.

### Labels
  ≥ 0.85 → "High"   (system is confident — go ahead)
  0.65–0.84 → "Medium" (review suggested — ambiguous query)
  < 0.65  → "Low"   (clarification recommended — query is unclear)
"""

from __future__ import annotations

import math
from dataclasses import dataclass

from nlp_sap.nlp.models import ParsedIntent
from nlp_sap.schema.registry import SchemaRegistry, get_registry


@dataclass
class ConfidenceResult:
    score: float                     # Point estimate [0, 1]
    lower: float                     # Wilson CI lower bound [0, 1]
    upper: float                     # Wilson CI upper bound [0, 1]
    label: str                       # "High" | "Medium" | "Low"
    evidence_count: int              # Number of signals observed
    breakdown: dict[str, float]      # Scores per evidence dimension

    def to_dict(self) -> dict:
        return {
            "score": round(self.score, 4),
            "interval": {
                "lower": round(self.lower, 4),
                "upper": round(self.upper, 4),
                "level": "95%",
                "method": "Wilson score",
            },
            "label": self.label,
            "evidence_count": self.evidence_count,
            "breakdown": {k: round(v, 4) for k, v in self.breakdown.items()},
        }


def _wilson_interval(p: float, n: int, z: float = 1.96) -> tuple[float, float]:
    """Compute Wilson score confidence interval.

    Args:
        p: Point estimate (0–1)
        n: Effective sample size (number of evidence signals)
        z: Z-score for desired confidence level (default 1.96 → 95%)

    Returns:
        (lower, upper) bounds clamped to [0, 1]
    """
    if n == 0:
        return 0.0, 1.0

    z2 = z * z
    centre_adj = p + z2 / (2 * n)
    margin = z * math.sqrt(p * (1 - p) / n + z2 / (4 * n * n))
    denominator = 1 + z2 / n

    lower = max(0.0, (centre_adj - margin) / denominator)
    upper = min(1.0, (centre_adj + margin) / denominator)
    return lower, upper


class ConfidenceScorer:
    """Computes calibrated confidence scores and Wilson-score CIs."""

    # Weights for each evidence dimension (must sum to 1.0)
    _WEIGHTS = {
        "llm_confidence": 0.45,
        "keyword_overlap": 0.20,
        "entity_completeness": 0.20,
        "schema_consistency": 0.15,
    }

    def __init__(self, registry: SchemaRegistry | None = None) -> None:
        self._registry = registry or get_registry()

    def score(self, intent: ParsedIntent, query: str) -> ConfidenceResult:
        """Compute a ConfidenceResult for a classified intent."""
        breakdown: dict[str, float] = {}

        # ── Signal 1: LLM self-reported confidence ───────────────────────────
        breakdown["llm_confidence"] = intent.confidence

        # ── Signal 2: Keyword overlap ────────────────────────────────────────
        breakdown["keyword_overlap"] = self._keyword_overlap(query, intent.intent_name)

        # ── Signal 3: Entity completeness (required fields found?) ───────────
        breakdown["entity_completeness"] = self._entity_completeness(intent)

        # ── Signal 4: Schema consistency (extracted values look SAP-valid?) ──
        breakdown["schema_consistency"] = self._schema_consistency(intent)

        # ── Weighted blend ───────────────────────────────────────────────────
        score = sum(
            breakdown[dim] * weight for dim, weight in self._WEIGHTS.items()
        )
        score = max(0.0, min(1.0, score))

        # ── Effective sample size ─────────────────────────────────────────────
        # Each dimension with a non-zero observation counts as one trial.
        # Scale n by 10 so CI width is meaningful (avoids near-binary bounds).
        evidence_count = sum(1 for v in breakdown.values() if v > 0.0)
        n_effective = evidence_count * 10

        lower, upper = _wilson_interval(score, n_effective)

        label = self._label(score)
        return ConfidenceResult(
            score=score,
            lower=lower,
            upper=upper,
            label=label,
            evidence_count=evidence_count,
            breakdown=breakdown,
        )

    # ── Private helpers ──────────────────────────────────────────────────────

    def _keyword_overlap(self, query: str, intent_name: str) -> float:
        """Fraction of intent keywords found in query."""
        intent_def = self._registry.get_intent(intent_name)
        if not intent_def:
            return 0.0
        keywords: list[str] = intent_def.get("keywords", [])
        if not keywords:
            return 0.5  # Neutral when no keywords defined
        q_lower = query.lower()
        hits = sum(1 for kw in keywords if kw.lower() in q_lower)
        return hits / len(keywords)

    def _entity_completeness(self, intent: ParsedIntent) -> float:
        """Score how many required filter fields were extracted."""
        intent_def = self._registry.get_intent(intent.intent_name)
        if not intent_def:
            return 0.5

        required = set(intent_def.get("required_filters", []))
        if not required:
            return 0.9  # No required filters → easy pass

        found = 0
        entity_map = {
            "company_code": intent.company_code,
            "fiscal_year": intent.fiscal_year,
            "gl_account": intent.gl_account,
            "customer": intent.customer,
            "vendor": intent.vendor,
            "material": intent.material,
            "plant": intent.plant,
            "cost_center": intent.cost_center,
            "profit_center": intent.profit_center,
            "controlling_area": intent.company_code,  # often same
        }
        for req in required:
            if entity_map.get(req):
                found += 1

        return found / len(required)

    def _schema_consistency(self, intent: ParsedIntent) -> float:
        """Check extracted entity values look SAP-plausible."""
        checks: list[bool] = []

        if intent.company_code:
            checks.append(len(intent.company_code) <= 4 and intent.company_code.isalnum())

        if intent.fiscal_year:
            try:
                year = int(intent.fiscal_year)
                checks.append(1990 <= year <= 2099)
            except ValueError:
                checks.append(False)

        if intent.fiscal_period:
            try:
                period = int(intent.fiscal_period)
                checks.append(1 <= period <= 16)  # SAP allows up to 16 special periods
            except ValueError:
                checks.append(False)

        if intent.gl_account:
            checks.append(len(intent.gl_account) <= 10)

        if intent.po_number:
            checks.append(intent.po_number.startswith("45") or len(intent.po_number) == 10)

        if intent.date_from:
            try:
                from datetime import date

                date.fromisoformat(intent.date_from)
                checks.append(True)
            except ValueError:
                checks.append(False)

        if not checks:
            return 0.8  # No entities to validate → neutral
        return sum(checks) / len(checks)

    @staticmethod
    def _label(score: float) -> str:
        if score >= 0.85:
            return "High"
        elif score >= 0.65:
            return "Medium"
        else:
            return "Low"
