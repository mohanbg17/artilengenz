"""Pydantic models for NLP pipeline inputs and outputs."""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class SAPModule(str, Enum):
    FI = "FI"
    CO = "CO"
    MM = "MM"
    SD = "SD"
    WM = "WM"
    PP = "PP"
    UNKNOWN = "UNKNOWN"


class ParsedIntent(BaseModel):
    """Output of the intent classifier."""

    intent_name: str
    module: SAPModule
    confidence: float = Field(ge=0.0, le=1.0)
    description: str = ""

    # Extracted entities
    company_code: str | None = None
    fiscal_year: str | None = None
    fiscal_period: str | None = None
    fiscal_periods: list[str] = Field(default_factory=list)
    gl_account: str | None = None
    cost_center: str | None = None
    profit_center: str | None = None
    customer: str | None = None
    vendor: str | None = None
    material: str | None = None
    plant: str | None = None
    storage_location: str | None = None
    warehouse: str | None = None
    document_number: str | None = None
    document_type: str | None = None
    sales_org: str | None = None
    po_number: str | None = None
    date_from: str | None = None
    date_to: str | None = None
    amount_min: float | None = None
    amount_max: float | None = None
    overdue_only: bool = False
    max_rows: int = 500

    # Raw LLM output for debugging
    raw_llm_json: dict[str, Any] = Field(default_factory=dict)


class SAPQueryPlan(BaseModel):
    """Structured execution plan generated from a ParsedIntent."""

    intent: ParsedIntent
    primary_table: str
    join_table: str | None = None
    join_on: dict[str, str] | None = None
    query_type: str = "table"          # "table" | "bapi" | "odata"
    odata_service: str | None = None
    odata_entity_set: str | None = None
    bapi_name: str | None = None
    bapi_import_params: dict[str, Any] = Field(default_factory=dict)
    bapi_output_tables: list[str] = Field(default_factory=list)
    filters: dict[str, Any] = Field(default_factory=dict)
    fields: list[str] = Field(default_factory=list)
    aggregations: list[str] = Field(default_factory=list)
    order_by: list[str] = Field(default_factory=list)
    max_rows: int = 500


class NLPQueryResponse(BaseModel):
    """Final response returned to the caller."""

    query: str                          # Original user query
    intent: str
    module: str
    confidence: float
    confidence_label: str              # "High" | "Medium" | "Low"
    confidence_interval: dict[str, float]  # {"lower": 0.82, "upper": 0.96}
    data: list[dict[str, Any]]
    total_count: int
    has_more: bool
    execution_plan: dict[str, Any]     # The query plan (for transparency)
    warnings: list[str] = Field(default_factory=list)
    error: str | None = None
    source: str = "mock"               # "odata" | "rfc" | "mock"
