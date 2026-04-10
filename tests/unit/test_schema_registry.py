"""Unit tests for the SAP Schema Registry."""

import pytest

from nlp_sap.schema.registry import SchemaRegistry, get_registry


@pytest.fixture
def registry() -> SchemaRegistry:
    return get_registry()


class TestSchemaRegistry:
    def test_get_known_table(self, registry: SchemaRegistry) -> None:
        schema = registry.get_table_schema("BKPF")
        assert schema is not None
        assert "fields" in schema
        assert "BUKRS" in schema["fields"]

    def test_get_s4_only_table(self, registry: SchemaRegistry) -> None:
        schema = registry.get_table_schema("ACDOCA")
        assert schema is not None
        assert "S4HANA" in schema["system"]

    def test_get_unknown_table_returns_none(self, registry: SchemaRegistry) -> None:
        assert registry.get_table_schema("ZZNOTEXIST") is None

    def test_all_intents_returns_dict(self, registry: SchemaRegistry) -> None:
        intents = registry.all_intents()
        assert len(intents) > 5
        assert "gl_balance_inquiry" in intents
        assert "ar_open_items" in intents
        assert "purchase_order_status" in intents

    def test_get_intent(self, registry: SchemaRegistry) -> None:
        intent = registry.get_intent("ar_open_items")
        assert intent is not None
        assert intent["module"] == "FI"
        assert "bapi" in intent or "primary_table" in intent

    def test_filter_alias_document_type(self, registry: SchemaRegistry) -> None:
        result = registry.filter_alias("document_types", "invoice")
        assert isinstance(result, list)
        assert "RE" in result or "KR" in result

    def test_fiscal_period_month(self, registry: SchemaRegistry) -> None:
        result = registry.fiscal_period("january")
        assert result == "01"

    def test_fiscal_period_quarter(self, registry: SchemaRegistry) -> None:
        result = registry.fiscal_period("Q1")
        assert isinstance(result, list)
        assert "01" in result and "03" in result

    def test_intent_summary_for_prompt_is_string(self, registry: SchemaRegistry) -> None:
        summary = registry.intent_summary_for_prompt()
        assert isinstance(summary, str)
        assert "gl_balance_inquiry" in summary
        assert len(summary) > 100

    def test_table_fields_for_prompt(self, registry: SchemaRegistry) -> None:
        fields_str = registry.table_fields_for_prompt("BKPF")
        assert "Company Code" in fields_str or "BUKRS" in fields_str

    def test_get_all_tables_for_module(self, registry: SchemaRegistry) -> None:
        tables = registry.get_all_tables_for_module("MM")
        assert "EKKO" in tables
        assert "MARA" in tables

    def test_time_expression_today(self, registry: SchemaRegistry) -> None:
        result = registry.time_expression("today")
        assert result == "CURRENT_DATE"
