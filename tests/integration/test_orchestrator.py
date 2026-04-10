"""Integration tests for the QueryOrchestrator using mock SAP data.

These tests exercise the full pipeline (classifier → confidence → builder →
connector) without a live SAP system or a real Anthropic API key.
The IntentClassifier is mocked to return pre-built ParsedIntents.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, patch

import pytest

from nlp_sap.config import Settings
from nlp_sap.connectors.mock import MockSAPConnector
from nlp_sap.nlp.models import NLPQueryResponse, ParsedIntent, SAPModule
from nlp_sap.orchestrator import QueryOrchestrator
from nlp_sap.schema.registry import get_registry


def _mock_intent(intent_name: str, **kwargs) -> ParsedIntent:
    defaults = {
        "intent_name": intent_name,
        "module": SAPModule.FI,
        "confidence": 0.92,
        "company_code": "1000",
        "fiscal_year": "2024",
    }
    defaults.update(kwargs)
    return ParsedIntent(**defaults)


def _make_orchestrator() -> QueryOrchestrator:
    settings = Settings(MOCK_SAP=True, ANTHROPIC_API_KEY="sk-test-key")
    connector = MockSAPConnector()
    return QueryOrchestrator(settings=settings, connector=connector, registry=get_registry())


class TestOrchestratorIntegration:
    @pytest.mark.asyncio
    async def test_ar_open_items_end_to_end(self) -> None:
        orch = _make_orchestrator()
        mock_intent = _mock_intent("ar_open_items", customer="C10001")

        with patch.object(orch._classifier, "classify", new=AsyncMock(return_value=mock_intent)):
            result = await orch.query("Show AR open items for customer C10001")

        assert isinstance(result, NLPQueryResponse)
        assert result.error is None
        assert result.intent == "ar_open_items"
        assert result.module == "FI"
        assert len(result.data) > 0
        assert 0.0 <= result.confidence <= 1.0
        assert result.confidence_label in ("High", "Medium", "Low")
        assert "lower" in result.confidence_interval
        assert "upper" in result.confidence_interval

    @pytest.mark.asyncio
    async def test_purchase_order_end_to_end(self) -> None:
        orch = _make_orchestrator()
        mock_intent = _mock_intent(
            "purchase_order_status",
            module=SAPModule.MM,
            vendor="V20001",
        )

        with patch.object(orch._classifier, "classify", new=AsyncMock(return_value=mock_intent)):
            result = await orch.query("List open purchase orders for vendor V20001")

        assert result.error is None
        assert result.intent == "purchase_order_status"
        assert result.total_count > 0

    @pytest.mark.asyncio
    async def test_gl_balance_end_to_end(self) -> None:
        orch = _make_orchestrator()
        mock_intent = _mock_intent(
            "gl_balance_inquiry",
            gl_account="400000",
            fiscal_year="2024",
        )

        with patch.object(orch._classifier, "classify", new=AsyncMock(return_value=mock_intent)):
            result = await orch.query("What is the GL balance for account 400000 in 2024?")

        assert result.error is None
        assert result.data

    @pytest.mark.asyncio
    async def test_confidence_interval_is_valid(self) -> None:
        orch = _make_orchestrator()
        mock_intent = _mock_intent("ar_open_items")

        with patch.object(orch._classifier, "classify", new=AsyncMock(return_value=mock_intent)):
            result = await orch.query("show AR items")

        ci = result.confidence_interval
        assert ci["lower"] <= result.confidence <= ci["upper"]
        assert 0.0 <= ci["lower"] <= ci["upper"] <= 1.0

    @pytest.mark.asyncio
    async def test_execution_plan_in_response(self) -> None:
        orch = _make_orchestrator()
        mock_intent = _mock_intent("inventory_stock", module=SAPModule.MM, material="M-001")

        with patch.object(orch._classifier, "classify", new=AsyncMock(return_value=mock_intent)):
            result = await orch.query("What is the stock for material M-001?")

        plan = result.execution_plan
        assert "primary_table" in plan
        assert "query_type" in plan
        assert "elapsed_ms" in plan

    @pytest.mark.asyncio
    async def test_classifier_error_returns_error_response(self) -> None:
        orch = _make_orchestrator()

        with patch.object(
            orch._classifier,
            "classify",
            new=AsyncMock(side_effect=Exception("LLM unavailable")),
        ):
            result = await orch.query("show something")

        assert result.error is not None
        assert "Classification error" in result.error
        assert result.confidence == 0.0

    @pytest.mark.asyncio
    async def test_context_manager_support(self) -> None:
        settings = Settings(MOCK_SAP=True, ANTHROPIC_API_KEY="sk-test")
        async with QueryOrchestrator(settings=settings) as orch:
            assert orch is not None
