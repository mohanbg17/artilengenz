"""Query Orchestrator — the central entry point for the NLP-SAP engine.

Flow:
  User Query (string)
      ↓
  IntentClassifier  →  ParsedIntent (intent + entities + raw confidence)
      ↓
  ConfidenceScorer  →  ConfidenceResult (calibrated score + Wilson CI)
      ↓
  QueryBuilder      →  SAPQueryPlan (what to query and how)
      ↓
  BaseSAPConnector  →  SAPQueryResult (raw SAP data)
      ↓
  ResultFormatter   →  NLPQueryResponse (final JSON to caller)
"""

from __future__ import annotations

import logging
import time
from typing import Any

from nlp_sap.config import Settings, get_settings
from nlp_sap.connectors.base import BaseSAPConnector
from nlp_sap.connectors.factory import build_connector
from nlp_sap.nlp.confidence import ConfidenceScorer
from nlp_sap.nlp.intent_classifier import IntentClassifier
from nlp_sap.nlp.models import NLPQueryResponse, ParsedIntent, SAPQueryPlan
from nlp_sap.nlp.query_builder import QueryBuilder
from nlp_sap.schema.registry import SchemaRegistry, get_registry

logger = logging.getLogger(__name__)


class QueryOrchestrator:
    """Wires the full NLP-to-SAP pipeline together."""

    def __init__(
        self,
        settings: Settings | None = None,
        connector: BaseSAPConnector | None = None,
        registry: SchemaRegistry | None = None,
    ) -> None:
        self._settings = settings or get_settings()
        self._registry = registry or get_registry()
        self._connector = connector or build_connector(self._settings)
        self._classifier = IntentClassifier(self._settings, self._registry)
        self._scorer = ConfidenceScorer(self._registry)
        self._builder = QueryBuilder(self._settings, self._registry)

    async def query(self, user_query: str) -> NLPQueryResponse:
        """Process a natural-language query end-to-end and return structured results.

        Args:
            user_query: Any natural language question about SAP finance or logistics.

        Returns:
            NLPQueryResponse with data, confidence score, CI, and execution plan.
        """
        start = time.monotonic()
        logger.info("Processing query: %s", user_query)

        # ── Step 1: Intent Classification ─────────────────────────────────────
        try:
            intent: ParsedIntent = await self._classifier.classify(user_query)
        except Exception as exc:
            logger.error("Intent classification failed: %s", exc)
            return self._error_response(user_query, f"Classification error: {exc}")

        # ── Step 2: Confidence Scoring ────────────────────────────────────────
        confidence_result = self._scorer.score(intent, user_query)
        conf_dict = confidence_result.to_dict()

        # ── Step 3: Build Query Plan ──────────────────────────────────────────
        try:
            plan: SAPQueryPlan = self._builder.build_plan(intent)
        except Exception as exc:
            logger.error("Query plan build failed: %s", exc)
            return self._error_response(user_query, f"Query build error: {exc}")

        # ── Step 4: Execute Against SAP ───────────────────────────────────────
        request = self._builder.plan_to_request(plan)
        try:
            sap_result = await self._connector.execute(request)
        except Exception as exc:
            logger.error("SAP execution failed: %s", exc)
            return self._error_response(user_query, f"SAP execution error: {exc}")

        elapsed_ms = round((time.monotonic() - start) * 1000, 1)
        logger.info(
            "Query resolved: intent=%s confidence=%.2f rows=%d elapsed=%sms",
            intent.intent_name,
            confidence_result.score,
            len(sap_result.data),
            elapsed_ms,
        )

        source = sap_result.metadata.get("source", "unknown")
        if not sap_result.success:
            return self._error_response(user_query, sap_result.error or "Unknown SAP error")

        return NLPQueryResponse(
            query=user_query,
            intent=intent.intent_name,
            module=intent.module.value,
            confidence=round(confidence_result.score, 4),
            confidence_label=confidence_result.label,
            confidence_interval={
                "lower": conf_dict["interval"]["lower"],
                "upper": conf_dict["interval"]["upper"],
            },
            data=sap_result.data,
            total_count=sap_result.total_count,
            has_more=sap_result.has_more,
            execution_plan=self._serialize_plan(plan, elapsed_ms),
            warnings=sap_result.warnings,
            source=source,
        )

    async def close(self) -> None:
        await self._connector.close()

    # ── Async context manager support ─────────────────────────────────────────

    async def __aenter__(self) -> "QueryOrchestrator":
        return self

    async def __aexit__(self, *args: Any) -> None:
        await self.close()

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _error_response(query: str, error: str) -> NLPQueryResponse:
        return NLPQueryResponse(
            query=query,
            intent="unknown",
            module="UNKNOWN",
            confidence=0.0,
            confidence_label="Low",
            confidence_interval={"lower": 0.0, "upper": 0.0},
            data=[],
            total_count=0,
            has_more=False,
            execution_plan={},
            error=error,
        )

    @staticmethod
    def _serialize_plan(plan: SAPQueryPlan, elapsed_ms: float) -> dict[str, Any]:
        return {
            "primary_table": plan.primary_table,
            "join_table": plan.join_table,
            "query_type": plan.query_type,
            "odata_service": plan.odata_service,
            "odata_entity_set": plan.odata_entity_set,
            "bapi": plan.bapi_name,
            "filters": plan.filters,
            "fields": plan.fields,
            "aggregations": plan.aggregations,
            "max_rows": plan.max_rows,
            "elapsed_ms": elapsed_ms,
        }
