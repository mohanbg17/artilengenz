"""Unit tests for the QueryBuilder."""

import pytest

from nlp_sap.config import SAPSystemType, Settings
from nlp_sap.connectors.base import SAPQueryRequest
from nlp_sap.nlp.models import ParsedIntent, SAPModule
from nlp_sap.nlp.query_builder import QueryBuilder
from nlp_sap.schema.registry import get_registry


def _make_intent(**kwargs) -> ParsedIntent:
    defaults = {
        "intent_name": "ar_open_items",
        "module": SAPModule.FI,
        "confidence": 0.9,
        "company_code": "1000",
        "fiscal_year": "2024",
        "customer": "C10001",
    }
    defaults.update(kwargs)
    return ParsedIntent(**defaults)


def _builder(system_type: SAPSystemType = SAPSystemType.S4HANA) -> QueryBuilder:
    settings = Settings(SAP_SYSTEM_TYPE=system_type, MOCK_SAP=True)
    return QueryBuilder(settings=settings, registry=get_registry())


class TestQueryBuilder:
    def test_ar_open_items_uses_bapi(self) -> None:
        builder = _builder()
        intent = _make_intent(intent_name="ar_open_items")
        plan = builder.build_plan(intent)
        assert plan.query_type == "bapi"
        assert plan.bapi_name == "BAPI_AR_ACC_GETOPENITEMS"
        assert plan.primary_table == "BSID"

    def test_ap_open_items_uses_bapi(self) -> None:
        builder = _builder()
        intent = _make_intent(intent_name="ap_open_items", vendor="V20001", customer=None)
        plan = builder.build_plan(intent)
        assert plan.bapi_name == "BAPI_AP_ACC_GETOPENITEMS"

    def test_gl_balance_s4hana_uses_odata(self) -> None:
        builder = _builder(SAPSystemType.S4HANA)
        intent = _make_intent(intent_name="gl_balance_inquiry", customer=None)
        plan = builder.build_plan(intent)
        assert plan.query_type == "odata"
        assert plan.odata_service == "API_GLACCOUNTLINEITEM_SRV"

    def test_gl_balance_ecc_uses_table(self) -> None:
        builder = _builder(SAPSystemType.ECC)
        intent = _make_intent(intent_name="gl_balance_inquiry", customer=None)
        plan = builder.build_plan(intent)
        assert plan.query_type == "table"
        assert plan.primary_table == "FAGLFLEXT"

    def test_po_status_s4hana_uses_odata(self) -> None:
        builder = _builder(SAPSystemType.S4HANA)
        intent = _make_intent(intent_name="purchase_order_status", customer=None)
        plan = builder.build_plan(intent)
        assert plan.query_type == "odata"
        assert plan.odata_service == "API_PURCHASEORDER_PROCESS_SRV"

    def test_po_status_ecc_uses_table_join(self) -> None:
        builder = _builder(SAPSystemType.ECC)
        intent = _make_intent(intent_name="purchase_order_status", customer=None)
        plan = builder.build_plan(intent)
        assert plan.query_type == "table"
        assert plan.primary_table == "EKKO"
        assert plan.join_table == "EKPO"

    def test_filters_mapped_for_s4hana_odata(self) -> None:
        builder = _builder(SAPSystemType.S4HANA)
        intent = _make_intent(
            intent_name="gl_balance_inquiry",
            company_code="1000",
            fiscal_year="2024",
            customer=None,
        )
        plan = builder.build_plan(intent)
        # S4HANA OData maps BUKRS → CompanyCode, GJAHR → FiscalYear
        assert "CompanyCode" in plan.filters or "BUKRS" in plan.filters

    def test_overdue_filter_adds_date_constraint(self) -> None:
        builder = _builder()
        intent = _make_intent(intent_name="ar_open_items", overdue_only=True)
        plan = builder.build_plan(intent)
        assert "FAEDT" in plan.filters

    def test_plan_to_request_produces_valid_request(self) -> None:
        builder = _builder()
        intent = _make_intent(intent_name="ar_open_items")
        plan = builder.build_plan(intent)
        request = builder.plan_to_request(plan)
        assert isinstance(request, SAPQueryRequest)
        assert request.query_type == plan.query_type

    def test_inventory_intent(self) -> None:
        builder = _builder()
        intent = _make_intent(
            intent_name="inventory_stock",
            module=SAPModule.MM,
            customer=None,
            material="M-001",
            plant="1000",
        )
        plan = builder.build_plan(intent)
        assert plan.primary_table == "MCHB"
        assert "MATNR" in plan.filters
        assert "WERKS" in plan.filters

    def test_generic_fallback_for_unknown_intent(self) -> None:
        builder = _builder()
        intent = _make_intent(intent_name="nonexistent_intent_xyz")
        plan = builder.build_plan(intent)
        # Should not raise; returns some table
        assert plan.primary_table
