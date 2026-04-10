"""SAP Schema Registry.

Loads table/BAPI/OData metadata from YAML and provides a lookup interface
used by the NLP engine to generate valid SAP queries.
"""

from __future__ import annotations

import functools
from pathlib import Path
from typing import Any

import yaml

_SCHEMA_PATH = Path(__file__).parent.parent.parent.parent / "config" / "sap_schemas.yaml"
_INTENT_PATH = Path(__file__).parent.parent.parent.parent / "config" / "intent_patterns.yaml"


@functools.lru_cache(maxsize=1)
def _load_yaml(path: Path) -> dict:
    with path.open() as f:
        return yaml.safe_load(f)


class SchemaRegistry:
    """Provides structured access to SAP table/BAPI/OData schemas."""

    def __init__(self) -> None:
        self._schemas: dict = _load_yaml(_SCHEMA_PATH)
        self._intents: dict = _load_yaml(_INTENT_PATH)

    # ── Public API ────────────────────────────────────────────────────────

    def get_table_schema(self, table_name: str) -> dict | None:
        """Return field definitions for a SAP table, searching all modules."""
        for module_data in self._schemas.get("modules", {}).values():
            tables = module_data.get("tables", {})
            if table_name in tables:
                return tables[table_name]
        return None

    def get_intent(self, intent_name: str) -> dict | None:
        """Return intent definition by name."""
        return self._intents.get("intents", {}).get(intent_name)

    def all_intents(self) -> dict[str, dict]:
        """Return all intent definitions."""
        return self._intents.get("intents", {})

    def get_odata_service(self, module: str, service_name: str) -> dict | None:
        """Return OData service definition for a module."""
        module_data = self._schemas.get("modules", {}).get(module, {})
        return module_data.get("odata_services", {}).get(service_name)

    def get_bapi(self, module: str, bapi_name: str) -> dict | None:
        """Return BAPI definition for a module."""
        module_data = self._schemas.get("modules", {}).get(module, {})
        return module_data.get("bapis", {}).get(bapi_name)

    def filter_alias(self, category: str, term: str) -> list[str] | str | None:
        """Resolve user term to SAP code(s). E.g. 'invoice' → ['RE','KR',...]."""
        aliases = self._intents.get("filter_aliases", {}).get(category, {})
        return aliases.get(term.lower())

    def fiscal_period(self, term: str) -> list[str] | str | None:
        """Resolve month/quarter name to SAP period number(s)."""
        mapping = self._intents.get("filter_aliases", {}).get("fiscal_period_map", {})
        return mapping.get(term) or mapping.get(term.upper())

    def time_expression(self, term: str) -> str | None:
        """Resolve relative time like 'today' to a symbolic token."""
        mapping = self._intents.get("filter_aliases", {}).get("time_expressions", {})
        return mapping.get(term.lower().replace(" ", "_"))

    def intent_summary_for_prompt(self) -> str:
        """Return a compact JSON-like summary of all intents for LLM context."""
        lines = []
        for name, intent in self.all_intents().items():
            examples = intent.get("example_queries", [])[:2]
            lines.append(
                f'  "{name}": {{'
                f'"module": "{intent.get("module")}", '
                f'"description": "{intent.get("description")}", '
                f'"examples": {examples}'
                f"}}"
            )
        return "{\n" + ",\n".join(lines) + "\n}"

    def table_fields_for_prompt(self, table_name: str) -> str:
        """Return a compact field list for prompt injection."""
        schema = self.get_table_schema(table_name)
        if not schema:
            return f"(unknown table: {table_name})"
        fields = schema.get("fields", {})
        parts = [f"{f}: {v.get('description', '')}" for f, v in fields.items()]
        return ", ".join(parts[:20])  # cap at 20 fields to keep prompt compact

    def get_all_tables_for_module(self, module: str) -> dict[str, dict]:
        """Return all tables for a given SAP module."""
        module_data = self._schemas.get("modules", {}).get(module, {})
        return module_data.get("tables", {})

    def as_dict(self) -> dict[str, Any]:
        """Full schema dump (for debugging)."""
        return self._schemas


# Singleton
_registry: SchemaRegistry | None = None


def get_registry() -> SchemaRegistry:
    global _registry
    if _registry is None:
        _registry = SchemaRegistry()
    return _registry
