"""IDoc poller — pulls failed IDocs from a connector and POSTs them to the
platform's ``/ingest/errors`` endpoint.

The poller does not write to the DB itself; it reuses the existing ABAP
ingest endpoint so the de-duplication, hashing, and audit logic stays in
one place. From the platform's perspective an IDoc record is just another
``raw.raw_errors`` row with ``source='IDOC'``.
"""
from __future__ import annotations

import os
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence

import httpx
import structlog

from .connector import IDocConnector
from .severity import DEFAULT_FAILED_STATUSES
from .transformer import idoc_to_ingest_record
from .watermark import MemoryWatermarkStore, WatermarkStore

log = structlog.get_logger(__name__)

SOURCE = "IDOC"


@dataclass
class PollResult:
    fetched: int
    inserted: int
    duplicates: int
    watermark_before: Optional[datetime]
    watermark_after: Optional[datetime]


class IDocPoller:
    """Single-tick poller. Wrap with ``run_forever`` for continuous mode."""

    def __init__(
        self,
        connector: IDocConnector,
        *,
        ingest_url: Optional[str] = None,
        ingest_token: Optional[str] = None,
        sap_host: Optional[str] = None,
        sap_sysnr: Optional[str] = None,
        sap_client: Optional[str] = None,
        statuses: Optional[Sequence[str]] = None,
        batch_size: int = 50,
        watermark_store: Optional[WatermarkStore] = None,
        http_client: Optional[httpx.Client] = None,
    ) -> None:
        self.connector = connector
        self.ingest_url = ingest_url or os.getenv(
            "IDOC_INGEST_URL", "http://localhost:8000/ingest/errors"
        )
        self.ingest_token = ingest_token or os.getenv("ARTILEGENZ_INGEST_TOKEN", "")
        self.sap_host = sap_host or os.getenv("IDOC_SAP_HOST", "")
        self.sap_sysnr = sap_sysnr or os.getenv("IDOC_SAP_SYSNR", "00")
        self.sap_client = sap_client or os.getenv("IDOC_SAP_CLIENT", "100")
        self.statuses = list(statuses) if statuses else list(DEFAULT_FAILED_STATUSES)
        self.batch_size = batch_size
        self.watermark = watermark_store or MemoryWatermarkStore()
        self.http = http_client or httpx.Client(timeout=30.0)

    @property
    def system_full_id(self) -> str:
        return f"{self.sap_host}:{self.sap_sysnr}:{self.sap_client}"

    def run_once(self) -> PollResult:
        if not self.ingest_token:
            raise RuntimeError(
                "ARTILEGENZ_INGEST_TOKEN is not set; refusing to POST without auth"
            )

        wm_before = self.watermark.get(SOURCE, self.system_full_id)
        idocs = self.connector.list_failed_idocs(
            since=wm_before,
            statuses=self.statuses,
            max_results=self.batch_size,
        )
        log.info(
            "idoc_poller_fetched",
            count=len(idocs),
            since=str(wm_before),
            system=self.system_full_id,
        )

        if not idocs:
            return PollResult(
                fetched=0,
                inserted=0,
                duplicates=0,
                watermark_before=wm_before,
                watermark_after=wm_before,
            )

        records = [idoc_to_ingest_record(d) for d in idocs]
        body = self._build_body(records)
        headers = self._build_headers()

        resp = self.http.post(self.ingest_url, json=body, headers=headers)
        resp.raise_for_status()
        result = resp.json()

        wm_after = max(d.created_at for d in idocs)
        self.watermark.set(SOURCE, self.system_full_id, wm_after)

        log.info(
            "idoc_poller_ingested",
            inserted=result.get("inserted", 0),
            duplicates=result.get("duplicates", 0),
            watermark_after=str(wm_after),
        )
        return PollResult(
            fetched=len(records),
            inserted=int(result.get("inserted", 0)),
            duplicates=int(result.get("duplicates", 0)),
            watermark_before=wm_before,
            watermark_after=wm_after,
        )

    def run_forever(self, interval_seconds: int = 60) -> None:
        log.info(
            "idoc_poller_starting",
            interval=interval_seconds,
            system=self.system_full_id,
        )
        while True:
            try:
                self.run_once()
            except KeyboardInterrupt:
                log.info("idoc_poller_interrupted")
                return
            except Exception:  # noqa: BLE001
                log.exception("idoc_poller_run_failed")
            time.sleep(interval_seconds)

    # ── helpers ──

    def _build_body(self, records: List[Dict[str, Any]]) -> Dict[str, Any]:
        return {
            "system_id": self.sap_sysnr,
            "client": self.sap_client,
            "host": self.sap_host,
            "batch_size": len(records),
            "records": records,
        }

    def _build_headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.ingest_token}",
            "X-SAP-System-Id": self.sap_sysnr,
            "X-SAP-Client": self.sap_client,
            "Content-Type": "application/json",
        }
