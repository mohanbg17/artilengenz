"""FastAPI backend for the SAP Error Intelligence dashboard.

Endpoints:
  GET  /healthz
  GET  /errors                — paginated triage feed with classification join
  GET  /errors/{hash_key}     — full error + proposals + citations
  POST /errors/{hash_key}/classify   — force (re)classification
  POST /errors/{hash_key}/feedback   — accept/reject + corrections
  GET  /analytics/summary     — counts by source/severity/module/badge
  POST /pipeline/sweep        — classify all unclassified errors (limit param)
"""
from __future__ import annotations

import asyncio
import hashlib
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

import structlog
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import snowflake.connector
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import serialization
from pathlib import Path

from config import settings
from classifier.orchestrator import Orchestrator
from scraper.sf_writer import fetch_unclassified_errors
from api.ingest import router as ingest_router

log = structlog.get_logger(__name__)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
    wrapper_class=structlog.make_filtering_bound_logger(getattr(logging, settings.log_level.upper(), logging.INFO)),
)


# --- Snowflake helpers (read-only paths) ------------------------------------

def _load_pkey() -> bytes:
    p = Path(settings.sf_private_key_path).expanduser()
    with open(p, "rb") as fh:
        passphrase = settings.sf_private_key_passphrase.encode() if settings.sf_private_key_passphrase else None
        key = serialization.load_pem_private_key(fh.read(), password=passphrase, backend=default_backend())
    return key.private_bytes(
        encoding=serialization.Encoding.DER,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )


def _connect():
    return snowflake.connector.connect(
        account=settings.sf_account,
        user=settings.sf_user,
        private_key=_load_pkey(),
        warehouse=settings.sf_warehouse,
        database=settings.sf_database,
        schema=settings.sf_intel_schema,
        role=settings.sf_role,
    )


# --- Pydantic response models ------------------------------------------------

class ErrorListItem(BaseModel):
    hash_key: str
    source: str
    severity: str
    occurred_at: datetime
    short_text: str
    transaction: Optional[str] = None
    classification_id: Optional[str] = None
    top_proposal_title: Optional[str] = None
    composite_confidence: Optional[float] = None
    badge: Optional[str] = None


class ErrorDetail(BaseModel):
    hash_key: str
    source: str
    severity: str
    system_id: str
    occurred_at: datetime
    short_text: str
    long_text: Optional[str] = None
    transaction: Optional[str] = None
    program: Optional[str] = None
    classification: Optional[Dict[str, Any]] = None


class FeedbackPayload(BaseModel):
    reviewer: str
    accepted: bool
    corrected_solution: Optional[str] = None
    comments: Optional[str] = None


class AnalyticsSummary(BaseModel):
    by_source: Dict[str, int]
    by_severity: Dict[str, int]
    by_badge: Dict[str, int]
    by_module: Dict[str, int]
    total_errors: int
    total_classified: int
    retry_rate: float


# --- App ---------------------------------------------------------------------

app = FastAPI(title="SAP Error Intelligence API", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(ingest_router)

orchestrator: Optional[Orchestrator] = None


@app.on_event("startup")
def _startup() -> None:
    global orchestrator
    orchestrator = Orchestrator()
    log.info("api_startup", model=settings.claude_model)


@app.get("/healthz")
def healthz() -> Dict[str, str]:
    return {"status": "ok", "model": settings.claude_model}


@app.get("/errors", response_model=List[ErrorListItem])
def list_errors(
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
    source: Optional[str] = None,
    severity: Optional[str] = None,
    badge: Optional[str] = None,
) -> List[ErrorListItem]:
    where: List[str] = []
    params: List[Any] = []
    if source:
        where.append("r.SOURCE = %s")
        params.append(source)
    if severity:
        where.append("r.SEVERITY = %s")
        params.append(severity)
    if badge:
        where.append("c.BADGE = %s")
        params.append(badge)
    where_sql = ("WHERE " + " AND ".join(where)) if where else ""
    sql = f"""
      SELECT r.HASH_KEY, r.SOURCE, r.SEVERITY, r.OCCURRED_AT, r.SHORT_TEXT, r.TRANSACTION,
             c.CLASSIFICATION_ID, c.TOP_PROPOSAL_TITLE, c.COMPOSITE_CONFIDENCE, c.BADGE
      FROM {settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS r
      LEFT JOIN {settings.sf_database}.{settings.sf_intel_schema}.CLASSIFICATIONS c
        ON r.HASH_KEY = c.ERROR_HASH_KEY
      {where_sql}
      ORDER BY COALESCE(c.COMPOSITE_CONFIDENCE, 0) DESC, r.OCCURRED_AT DESC
      LIMIT {int(limit)} OFFSET {int(offset)}
    """
    conn = _connect()
    cur = conn.cursor(snowflake.connector.DictCursor)
    try:
        cur.execute(sql, params)
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()
    return [
        ErrorListItem(
            hash_key=r["HASH_KEY"],
            source=r["SOURCE"],
            severity=r["SEVERITY"] or "ERROR",
            occurred_at=r["OCCURRED_AT"],
            short_text=r["SHORT_TEXT"] or "",
            transaction=r.get("TRANSACTION"),
            classification_id=r.get("CLASSIFICATION_ID"),
            top_proposal_title=r.get("TOP_PROPOSAL_TITLE"),
            composite_confidence=r.get("COMPOSITE_CONFIDENCE"),
            badge=r.get("BADGE"),
        )
        for r in rows
    ]


@app.get("/errors/{hash_key}", response_model=ErrorDetail)
def get_error(hash_key: str) -> ErrorDetail:
    sql = f"""
      SELECT r.HASH_KEY, r.SOURCE, r.SEVERITY, r.SYSTEM_ID, r.OCCURRED_AT,
             r.SHORT_TEXT, r.LONG_TEXT, r.TRANSACTION, r.PROGRAM,
             c.CLASSIFICATION_ID, c.TOP_PROPOSAL_TITLE, c.COMPOSITE_CONFIDENCE,
             c.BADGE, c.RETRIED, c.PROPOSALS, c.CITATIONS, c.CRITIQUE_NOTES,
             c.MODEL_VERSION, c.CREATED_AT
      FROM {settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS r
      LEFT JOIN {settings.sf_database}.{settings.sf_intel_schema}.CLASSIFICATIONS c
        ON r.HASH_KEY = c.ERROR_HASH_KEY
      WHERE r.HASH_KEY = %s
    """
    conn = _connect()
    cur = conn.cursor(snowflake.connector.DictCursor)
    try:
        cur.execute(sql, (hash_key,))
        row = cur.fetchone()
    finally:
        cur.close()
        conn.close()
    if not row:
        raise HTTPException(status_code=404, detail="error not found")

    classification = None
    if row.get("CLASSIFICATION_ID"):
        classification = {
            "classification_id": row["CLASSIFICATION_ID"],
            "top_proposal_title": row.get("TOP_PROPOSAL_TITLE"),
            "composite_confidence": row.get("COMPOSITE_CONFIDENCE"),
            "badge": row.get("BADGE"),
            "retried": row.get("RETRIED"),
            "proposals": row.get("PROPOSALS"),
            "citations": row.get("CITATIONS"),
            "critique_notes": row.get("CRITIQUE_NOTES"),
            "model_version": row.get("MODEL_VERSION"),
            "created_at": row.get("CREATED_AT"),
        }

    return ErrorDetail(
        hash_key=row["HASH_KEY"],
        source=row["SOURCE"],
        severity=row["SEVERITY"] or "ERROR",
        system_id=row["SYSTEM_ID"],
        occurred_at=row["OCCURRED_AT"],
        short_text=row["SHORT_TEXT"] or "",
        long_text=row.get("LONG_TEXT"),
        transaction=row.get("TRANSACTION"),
        program=row.get("PROGRAM"),
        classification=classification,
    )


@app.post("/errors/{hash_key}/classify")
async def force_classify(hash_key: str) -> Dict[str, Any]:
    if orchestrator is None:
        raise HTTPException(status_code=503, detail="orchestrator not ready")
    # Pull the error row
    detail = get_error(hash_key)
    error_dict = {
        "HASH_KEY": detail.hash_key,
        "SOURCE": detail.source,
        "SEVERITY": detail.severity,
        "SYSTEM_ID": detail.system_id,
        "OCCURRED_AT": detail.occurred_at,
        "SHORT_TEXT": detail.short_text,
        "LONG_TEXT": detail.long_text,
        "TRANSACTION": detail.transaction,
        "PROGRAM": detail.program,
    }
    payload = await orchestrator.run(error_dict)
    return {"ok": True, "classification_id": payload["classification_id"], "badge": payload["badge"]}


@app.post("/errors/{hash_key}/feedback")
def submit_feedback(hash_key: str, body: FeedbackPayload) -> Dict[str, str]:
    # Find latest classification id for this error
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT CLASSIFICATION_ID FROM CLASSIFICATIONS WHERE ERROR_HASH_KEY = %s ORDER BY CREATED_AT DESC LIMIT 1",
            (hash_key,),
        )
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="no classification to give feedback on")
        cls_id = row[0]
        fb_id = "fb-" + hashlib.sha256(f"{cls_id}|{datetime.utcnow().isoformat()}".encode()).hexdigest()[:24]
        cur.execute(
            """
            INSERT INTO HUMAN_FEEDBACK (FEEDBACK_ID, CLASSIFICATION_ID, REVIEWER, REVIEWED_AT,
                                        ACCEPTED, CORRECTED_SOLUTION, COMMENTS)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            (fb_id, cls_id, body.reviewer, datetime.utcnow(), body.accepted,
             body.corrected_solution, body.comments),
        )
        conn.commit()
    finally:
        cur.close()
        conn.close()
    return {"ok": "true", "feedback_id": fb_id}


@app.get("/analytics/summary", response_model=AnalyticsSummary)
def analytics_summary() -> AnalyticsSummary:
    conn = _connect()
    cur = conn.cursor(snowflake.connector.DictCursor)
    try:
        cur.execute(f"""
          WITH joined AS (
            SELECT r.SOURCE, r.SEVERITY, c.BADGE, c.RETRIED,
                   PARSE_JSON(c.PROPOSALS):0:module::STRING AS MODULE
            FROM {settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS r
            LEFT JOIN CLASSIFICATIONS c ON r.HASH_KEY = c.ERROR_HASH_KEY
          )
          SELECT
            OBJECT_AGG(SOURCE, src_count) AS by_source,
            OBJECT_AGG(SEVERITY, sev_count) AS by_severity,
            OBJECT_AGG(BADGE, badge_count) AS by_badge,
            OBJECT_AGG(MODULE, mod_count) AS by_module,
            (SELECT COUNT(*) FROM {settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS) AS total_errors,
            (SELECT COUNT(*) FROM CLASSIFICATIONS) AS total_classified,
            (SELECT AVG(IFF(RETRIED, 1, 0)) FROM CLASSIFICATIONS) AS retry_rate
          FROM (
            SELECT SOURCE, SEVERITY, BADGE, MODULE,
                   COUNT(*) OVER (PARTITION BY SOURCE) AS src_count,
                   COUNT(*) OVER (PARTITION BY SEVERITY) AS sev_count,
                   COUNT(*) OVER (PARTITION BY BADGE) AS badge_count,
                   COUNT(*) OVER (PARTITION BY MODULE) AS mod_count
            FROM joined
          )
          LIMIT 1
        """)
        row = cur.fetchone() or {}
    finally:
        cur.close()
        conn.close()
    return AnalyticsSummary(
        by_source=row.get("BY_SOURCE") or {},
        by_severity=row.get("BY_SEVERITY") or {},
        by_badge=row.get("BY_BADGE") or {},
        by_module=row.get("BY_MODULE") or {},
        total_errors=row.get("TOTAL_ERRORS") or 0,
        total_classified=row.get("TOTAL_CLASSIFIED") or 0,
        retry_rate=float(row.get("RETRY_RATE") or 0.0),
    )


@app.post("/pipeline/sweep")
async def sweep_unclassified(limit: int = Query(50, ge=1, le=500)) -> Dict[str, Any]:
    """Classify up to `limit` unclassified errors. Runs in-process (foreground)."""
    if orchestrator is None:
        raise HTTPException(status_code=503, detail="orchestrator not ready")
    rows = fetch_unclassified_errors(limit=limit)
    log.info("sweep_start", count=len(rows))
    results: List[Dict[str, Any]] = []
    for r in rows:
        try:
            payload = await orchestrator.run(r)
            results.append({
                "hash_key": r["HASH_KEY"],
                "badge": payload["badge"],
                "confidence": payload["composite_confidence"],
            })
        except Exception:  # noqa: BLE001
            log.exception("sweep_error", hash_key=r.get("HASH_KEY"))
    return {"processed": len(results), "results": results}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=settings.api_host, port=settings.api_port)
