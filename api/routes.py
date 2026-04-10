"""API routes for the NLP-SAP Query Engine."""

from __future__ import annotations

import io
from typing import Annotated, Any

import pandas as pd
from fastapi import APIRouter, Depends, HTTPException, Query, Request, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from nlp_sap.connectors.probe import ProbeResult, probe
from nlp_sap.nlp.models import NLPQueryResponse
from nlp_sap.orchestrator import QueryOrchestrator

router = APIRouter()


# ── Request / Response models ─────────────────────────────────────────────────

class QueryRequest(BaseModel):
    """Payload for POST /query."""

    query: str = Field(
        ...,
        min_length=3,
        max_length=2000,
        examples=["Show me all open AR items for customer 10001 in company code 1000"],
    )
    max_rows: int = Field(default=500, ge=1, le=5000)
    export_format: str | None = Field(
        default=None,
        description="Optional: 'csv' or 'excel' to stream a file download",
    )


class HealthResponse(BaseModel):
    status: str
    mock_sap: bool
    version: str


class IntentListResponse(BaseModel):
    intents: list[dict[str, Any]]
    total: int


# ── Dependency ────────────────────────────────────────────────────────────────

def get_orchestrator(request: Request) -> QueryOrchestrator:
    return request.app.state.orchestrator


OrchestratorDep = Annotated[QueryOrchestrator, Depends(get_orchestrator)]


# ── Routes ────────────────────────────────────────────────────────────────────

@router.post(
    "/query",
    response_model=NLPQueryResponse,
    summary="Execute a natural-language SAP query",
    tags=["query"],
)
async def post_query(
    body: QueryRequest,
    orchestrator: OrchestratorDep,
) -> NLPQueryResponse:
    """Submit a natural language question and get SAP data back.

    Examples:
    - "What is the GL balance for account 400000 in company code 1000 for 2024?"
    - "Show all open AR items for customer C10001 overdue as of today"
    - "List purchase orders for vendor V20001 created this month"
    - "What is the inventory level for material M-001 in plant 1000?"
    """
    result = await orchestrator.query(body.query)
    if result.error:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=result.error,
        )
    return result


@router.get(
    "/query",
    response_model=NLPQueryResponse,
    summary="Execute a natural-language SAP query (GET)",
    tags=["query"],
)
async def get_query(
    q: Annotated[str, Query(description="Natural language SAP query", min_length=3)],
    orchestrator: OrchestratorDep,
    max_rows: int = Query(default=500, ge=1, le=5000),
) -> NLPQueryResponse:
    """GET version of the query endpoint — convenient for browser/curl testing."""
    result = await orchestrator.query(q)
    if result.error:
        raise HTTPException(status_code=422, detail=result.error)
    return result


@router.post(
    "/query/export",
    summary="Execute query and download result as CSV or Excel",
    tags=["query"],
)
async def export_query(
    body: QueryRequest,
    orchestrator: OrchestratorDep,
) -> StreamingResponse:
    """Run a query and stream the result as a downloadable file."""
    result = await orchestrator.query(body.query)
    if result.error:
        raise HTTPException(status_code=422, detail=result.error)

    fmt = (body.export_format or "csv").lower()
    df = pd.DataFrame(result.data)

    if fmt == "excel":
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="SAP_Report")
            # Write metadata to second sheet
            meta = pd.DataFrame([{
                "query": result.query,
                "intent": result.intent,
                "confidence": result.confidence,
                "confidence_label": result.confidence_label,
                "total_count": result.total_count,
                "source": result.source,
            }])
            meta.to_excel(writer, index=False, sheet_name="Metadata")
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="sap_report.xlsx"'},
        )
    else:
        content = df.to_csv(index=False)
        return StreamingResponse(
            iter([content]),
            media_type="text/csv",
            headers={"Content-Disposition": 'attachment; filename="sap_report.csv"'},
        )


@router.get(
    "/intents",
    response_model=IntentListResponse,
    summary="List all supported SAP query intents",
    tags=["metadata"],
)
async def list_intents(orchestrator: OrchestratorDep) -> IntentListResponse:
    """Return the full catalogue of intents this system can handle."""
    from nlp_sap.schema.registry import get_registry

    registry = get_registry()
    intents = []
    for name, defn in registry.all_intents().items():
        intents.append({
            "name": name,
            "module": defn.get("module"),
            "description": defn.get("description"),
            "keywords": defn.get("keywords", []),
            "example_queries": defn.get("example_queries", []),
            "required_filters": defn.get("required_filters", []),
        })
    return IntentListResponse(intents=intents, total=len(intents))


@router.get(
    "/schema/{table_name}",
    summary="Get SAP table field definitions",
    tags=["metadata"],
)
async def get_table_schema(table_name: str) -> dict:
    """Return field definitions for a given SAP table."""
    from nlp_sap.schema.registry import get_registry

    registry = get_registry()
    schema = registry.get_table_schema(table_name.upper())
    if not schema:
        raise HTTPException(
            status_code=404,
            detail=f"Table '{table_name}' not found in schema registry",
        )
    return {"table": table_name.upper(), "schema": schema}


@router.get(
    "/ping",
    summary="Check SAP system connectivity",
    tags=["health"],
)
async def ping_sap(orchestrator: OrchestratorDep) -> dict:
    """Ping the SAP system to verify connectivity."""
    reachable = await orchestrator._connector.ping()
    return {
        "sap_reachable": reachable,
        "mock_mode": orchestrator._settings.mock_sap,
    }


@router.get(
    "/connect",
    summary="Full SAP S/4HANA connection probe (TCP + OData + Auth + Services)",
    tags=["health"],
)
async def connect_probe(orchestrator: OrchestratorDep) -> dict:
    """Run a deep connectivity check against the configured SAP system.

    Returns:
    - TCP reachability
    - OData endpoint status
    - Authentication result
    - Which NLP engine OData services are activated
    - Suggested .env block to paste

    Safe to call without admin rights — only uses HTTP(S).
    """
    settings = orchestrator._settings
    if settings.mock_sap:
        return {
            "mode": "mock",
            "ready": True,
            "message": "Running in mock mode — no live SAP connection needed.",
            "to_use_live_sap": "Set MOCK_SAP=false in .env and restart.",
        }

    result: ProbeResult = await probe(settings)
    activation_hints = []
    for svc in result.services_missing:
        activation_hints.append(
            f"Activate '{svc}' in SAP: transaction /IWFND/MAINT_SERVICE → Add Service"
        )

    return {
        "host":             result.host,
        "port":             result.port,
        "tcp_reachable":    result.tcp_reachable,
        "odata_reachable":  result.odata_reachable,
        "authenticated":    result.authenticated,
        "services_ok":      result.services_ok,
        "services_missing": result.services_missing,
        "elapsed_ms":       result.elapsed_ms,
        "ready":            result.ready,
        "error":            result.error,
        "activation_hints": activation_hints,
        "suggested_env":    result.suggested_env,
    }
