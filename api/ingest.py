"""ABAP-push ingest endpoint."""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import structlog
from fastapi import APIRouter, Header, HTTPException, status
from pydantic import BaseModel, Field

from db import get_backend

log = structlog.get_logger(__name__)

INGEST_TOKEN = os.getenv("ARTILEGENZ_INGEST_TOKEN", "")

router = APIRouter(prefix="/ingest", tags=["ingest"])


class AbapErrorRecord(BaseModel):
    source: str
    error_id: Optional[str] = None
    occurred_at: str
    severity: str = "ERROR"
    short_text: str = ""
    long_text: Optional[str] = None
    user_name: Optional[str] = None
    transaction: Optional[str] = None
    program: Optional[str] = None
    instance: Optional[str] = None
    work_process: Optional[str] = None
    object: Optional[str] = None
    sub_object: Optional[str] = None
    job_name: Optional[str] = None
    raw: Optional[str] = None


class AbapBatch(BaseModel):
    system_id: Optional[str] = None
    client: Optional[str] = None
    host: Optional[str] = None
    batch_size: Optional[int] = Field(default=None, ge=0)
    records: List[AbapErrorRecord]


def _parse_abap_timestamp(ts: str) -> datetime:
    if not ts:
        return datetime.utcnow()
    try:
        return datetime.strptime(ts.rstrip("Z"), "%Y%m%dT%H%M%S")
    except ValueError:
        try:
            return datetime.fromisoformat(ts.rstrip("Z"))
        except ValueError:
            log.warning("ingest_bad_timestamp", value=ts)
            return datetime.utcnow()


def _hash_key(record: AbapErrorRecord, system_full_id: str, occurred: datetime) -> str:
    short_norm = " ".join((record.short_text or "").lower().split())[:200]
    parts = [
        record.source,
        system_full_id,
        record.error_id or "",
        occurred.isoformat(timespec="seconds"),
        short_norm,
    ]
    return hashlib.sha256("|".join(parts).encode()).hexdigest()


def _to_row(record: AbapErrorRecord, *, system_full_id: str) -> Dict[str, Any]:
    occurred = _parse_abap_timestamp(record.occurred_at)
    raw_payload: Any = {}
    if record.raw:
        try:
            raw_payload = json.loads(record.raw)
        except json.JSONDecodeError:
            raw_payload = {"abap_raw": record.raw}

    return {
        "hash_key": _hash_key(record, system_full_id, occurred),
        "source": record.source,
        "system_id": system_full_id,
        "occurred_at": occurred,
        "extracted_at": datetime.utcnow(),
        "severity": record.severity,
        "short_text": (record.short_text or "")[:2000],
        "long_text": record.long_text,
        "error_id": record.error_id,
        "user_name": record.user_name,
        "transaction": record.transaction,
        "program": record.program,
        "instance": record.instance,
        "work_process": record.work_process,
        "object": record.object,
        "sub_object": record.sub_object,
        "job_name": record.job_name,
        "raw": raw_payload,
    }


@router.post("/errors")
async def ingest_errors(
    batch: AbapBatch,
    authorization: str = Header(""),
    x_sap_system_id: str = Header(""),
    x_sap_client: str = Header(""),
) -> Dict[str, Any]:
    if not INGEST_TOKEN:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="ARTILEGENZ_INGEST_TOKEN not configured on server",
        )
    if authorization != f"Bearer {INGEST_TOKEN}":
        log.warning("ingest_auth_failed", got_prefix=authorization[:10] if authorization else "")
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="invalid bearer token")

    if not batch.records:
        return {"status": "ok", "received": 0, "inserted": 0, "duplicates": 0}

    # Resolve identity from body, then headers, then defaults
    host = batch.host or x_sap_system_id or "unknown"
    client = batch.client or x_sap_client or "100"
    sysnr = x_sap_system_id or batch.system_id or "00"
    system_full_id = f"{host}:{sysnr}:{client}"

    log.info(
        "ingest_batch_received",
        system_id=system_full_id,
        host=host,
        size=len(batch.records),
    )

    rows = [_to_row(r, system_full_id=system_full_id) for r in batch.records]
    inserted = get_backend().upsert_raw_errors(rows)
    duplicates = len(rows) - inserted

    log.info("ingest_batch_persisted", inserted=inserted, duplicates=duplicates)
    return {
        "status": "ok",
        "received": len(rows),
        "inserted": inserted,
        "duplicates": duplicates,
    }


@router.get("/errors/health")
def ingest_health() -> Dict[str, Any]:
    backend = get_backend()
    return {
        "status": "ok",
        "endpoint": "ready",
        "auth_configured": "yes" if INGEST_TOKEN else "no",
        "db_engine": backend.name,
    }