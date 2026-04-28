"""OData v2 client for SAP standard IDoc service.

The actual S/4 service path varies by release. Common options:

  * ``/sap/opu/odata/sap/API_IDOC_SRV``       (S/4HANA Cloud / Public Cloud APIs)
  * ``/sap/opu/odata/sap/IDOC_DISPLAY_SRV``   (some on-prem releases)
  * a custom Z-service exposing CDS view ZI_FAILED_IDOC

Configure ``IDOC_ODATA_BASE_URL`` plus the entity set names if your service
uses non-standard names. Defaults assume API_IDOC_SRV.

This module is a thin client: build query → GET → parse JSON → emit IDoc.
HTTP transport is injectable so tests use ``httpx.MockTransport`` instead of
hitting a network.
"""
from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence

import httpx
import structlog

from .connector import IDocConnector
from .models import INBOUND, OUTBOUND, IDoc, IDocSegment, IDocStatus

log = structlog.get_logger(__name__)


# ---------------------------------------------------------------------------
# Defaults — overridable via env or constructor args
# ---------------------------------------------------------------------------
DEFAULT_BASE_URL = os.getenv(
    "IDOC_ODATA_BASE_URL",
    "https://your-s4-host/sap/opu/odata/sap/API_IDOC_SRV",
)
DEFAULT_ENTITY_HEADER = os.getenv("IDOC_ENTITY_HEADER", "Idoc")
DEFAULT_ENTITY_STATUS = os.getenv("IDOC_ENTITY_STATUS", "IdocStatusRecord")
DEFAULT_ENTITY_SEGMENT = os.getenv("IDOC_ENTITY_SEGMENT", "IdocSegment")

# Field names on the OData entity. Real S/4 services vary; override via env.
DEFAULT_FIELD_DOCNUM = os.getenv("IDOC_FIELD_DOCNUM", "IdocNumber")
DEFAULT_FIELD_STATUS = os.getenv("IDOC_FIELD_STATUS", "Status")
DEFAULT_FIELD_CREATED = os.getenv("IDOC_FIELD_CREATED", "CreationDateTime")
DEFAULT_FIELD_DIRECTION = os.getenv("IDOC_FIELD_DIRECTION", "Direction")
DEFAULT_FIELD_MESTYP = os.getenv("IDOC_FIELD_MESTYP", "MessageType")
DEFAULT_FIELD_IDOCTP = os.getenv("IDOC_FIELD_IDOCTP", "BasicType")
DEFAULT_FIELD_SNDPRN = os.getenv("IDOC_FIELD_SNDPRN", "SenderPartnerNumber")
DEFAULT_FIELD_RCVPRN = os.getenv("IDOC_FIELD_RCVPRN", "ReceiverPartnerNumber")
DEFAULT_FIELD_CLIENT = os.getenv("IDOC_FIELD_CLIENT", "Client")


def _odata_v2_datetime(ts: datetime) -> str:
    """Format a Python datetime as the OData v2 ``datetime'YYYY-MM-DDTHH:MM:SS'`` literal."""
    return f"datetime'{ts.strftime('%Y-%m-%dT%H:%M:%S')}'"


def _parse_odata_datetime(raw: Any) -> datetime:
    """Parse the various datetime shapes SAP OData services return.

    Common forms:
      * ``"/Date(1700000000000)/"``           (OData v2 JSON ticks)
      * ``"2025-04-15T10:23:45"``             (ISO-8601, no zone)
      * ``"2025-04-15T10:23:45Z"``            (ISO-8601 UTC)
      * already-a-datetime                    (when callers hydrate first)
    """
    if isinstance(raw, datetime):
        return raw
    if not raw:
        return datetime.utcnow()
    if isinstance(raw, str):
        if raw.startswith("/Date(") and raw.endswith(")/"):
            try:
                ms = int(raw[6:-2].split("+", 1)[0].split("-", 1)[0])
                return datetime.utcfromtimestamp(ms / 1000.0)
            except (ValueError, IndexError):
                return datetime.utcnow()
        try:
            return datetime.fromisoformat(raw.rstrip("Z"))
        except ValueError:
            return datetime.utcnow()
    return datetime.utcnow()


def _normalize_direction(value: Any) -> str:
    """Map the various SAP direction representations to INBOUND / OUTBOUND."""
    if value in (1, "1", "OUTBOUND", "out", "Outbound"):
        return OUTBOUND
    if value in (2, "2", "INBOUND", "in", "Inbound"):
        return INBOUND
    return INBOUND  # safe default; transformer still reads the raw value


class ODataIDocClient(IDocConnector):
    """Pull failed IDocs from a SAP OData v2 service."""

    def __init__(
        self,
        *,
        base_url: str = DEFAULT_BASE_URL,
        username: Optional[str] = None,
        password: Optional[str] = None,
        entity_header: str = DEFAULT_ENTITY_HEADER,
        entity_status: str = DEFAULT_ENTITY_STATUS,
        entity_segment: str = DEFAULT_ENTITY_SEGMENT,
        field_docnum: str = DEFAULT_FIELD_DOCNUM,
        field_status: str = DEFAULT_FIELD_STATUS,
        field_created: str = DEFAULT_FIELD_CREATED,
        field_direction: str = DEFAULT_FIELD_DIRECTION,
        field_mestyp: str = DEFAULT_FIELD_MESTYP,
        field_idoctp: str = DEFAULT_FIELD_IDOCTP,
        field_sndprn: str = DEFAULT_FIELD_SNDPRN,
        field_rcvprn: str = DEFAULT_FIELD_RCVPRN,
        field_client: str = DEFAULT_FIELD_CLIENT,
        http_client: Optional[httpx.Client] = None,
        verify_ssl: bool = True,
        timeout: float = 30.0,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.entity_header = entity_header
        self.entity_status = entity_status
        self.entity_segment = entity_segment
        self.f_docnum = field_docnum
        self.f_status = field_status
        self.f_created = field_created
        self.f_direction = field_direction
        self.f_mestyp = field_mestyp
        self.f_idoctp = field_idoctp
        self.f_sndprn = field_sndprn
        self.f_rcvprn = field_rcvprn
        self.f_client = field_client

        if http_client is not None:
            self.http = http_client
        else:
            auth = (username, password) if (username and password) else None
            self.http = httpx.Client(
                auth=auth,
                timeout=timeout,
                verify=verify_ssl,
                headers={"Accept": "application/json"},
            )

    # ── public API ────────────────────────────────────────────────────────

    def list_failed_idocs(
        self,
        *,
        since: Optional[datetime] = None,
        statuses: Optional[Sequence[str]] = None,
        max_results: int = 100,
    ) -> List[IDoc]:
        params = self._build_query_params(
            since=since, statuses=statuses, top=max_results
        )
        url = f"{self.base_url}/{self.entity_header}"
        log.info("idoc_odata_query", url=url, since=str(since), statuses=list(statuses or []))
        resp = self.http.get(url, params=params)
        resp.raise_for_status()
        records = self._extract_results(resp.json())
        idocs = [self._parse_header(rec) for rec in records]
        # Stable monotone order so watermark advances correctly
        idocs.sort(key=lambda d: (d.created_at, d.docnum))
        return idocs

    # ── internals ─────────────────────────────────────────────────────────

    def _build_query_params(
        self,
        *,
        since: Optional[datetime],
        statuses: Optional[Sequence[str]],
        top: int,
    ) -> Dict[str, Any]:
        filters: List[str] = []
        if statuses:
            joined = " or ".join(f"{self.f_status} eq '{s}'" for s in statuses)
            filters.append(f"({joined})")
        if since is not None:
            filters.append(f"{self.f_created} gt {_odata_v2_datetime(since)}")

        params: Dict[str, Any] = {
            "$format": "json",
            "$top": str(int(top)),
            "$orderby": f"{self.f_created} asc",
            "$expand": f"{self.entity_status},{self.entity_segment}",
        }
        if filters:
            params["$filter"] = " and ".join(filters)
        return params

    @staticmethod
    def _extract_results(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
        """OData v2 wraps results in {'d': {'results': [...]}}; v4 in {'value': [...]}."""
        if not isinstance(payload, dict):
            return []
        if "value" in payload and isinstance(payload["value"], list):
            return payload["value"]
        d = payload.get("d")
        if isinstance(d, dict) and isinstance(d.get("results"), list):
            return d["results"]
        if isinstance(d, list):
            return d
        return []

    def _parse_header(self, rec: Dict[str, Any]) -> IDoc:
        statuses = self._parse_inline(rec.get(self.entity_status))
        segments = self._parse_inline(rec.get(self.entity_segment))

        return IDoc(
            docnum=str(rec.get(self.f_docnum, "")),
            mestyp=str(rec.get(self.f_mestyp, "") or ""),
            idoctp=str(rec.get(self.f_idoctp, "") or ""),
            direction=_normalize_direction(rec.get(self.f_direction)),
            status=str(rec.get(self.f_status, "") or ""),
            created_at=_parse_odata_datetime(rec.get(self.f_created)),
            sender_partner=str(rec.get(self.f_sndprn, "") or ""),
            receiver_partner=str(rec.get(self.f_rcvprn, "") or ""),
            client=str(rec.get(self.f_client, "") or ""),
            statuses=[self._parse_status(s) for s in statuses],
            segments=[self._parse_segment(s) for s in segments],
        )

    @staticmethod
    def _parse_inline(value: Any) -> List[Dict[str, Any]]:
        """OData v2 expands embed as {'results': [...]}; v4 returns the list directly."""
        if value is None:
            return []
        if isinstance(value, dict) and isinstance(value.get("results"), list):
            return value["results"]
        if isinstance(value, list):
            return value
        return []

    @staticmethod
    def _parse_status(rec: Dict[str, Any]) -> IDocStatus:
        return IDocStatus(
            counter=int(rec.get("Counter") or rec.get("StatusCounter") or 0),
            status=str(rec.get("Status", "") or ""),
            status_text=str(rec.get("StatusText") or rec.get("Text") or ""),
            log_at=_parse_odata_datetime(
                rec.get("LogDateTime") or rec.get("ChangedAt") or rec.get("CreationDateTime")
            ),
            msg_id=rec.get("MessageId") or rec.get("MsgId"),
            msg_no=rec.get("MessageNumber") or rec.get("MsgNumber"),
            msg_v1=rec.get("MessageV1") or rec.get("MsgV1"),
            msg_v2=rec.get("MessageV2") or rec.get("MsgV2"),
            msg_v3=rec.get("MessageV3") or rec.get("MsgV3"),
            msg_v4=rec.get("MessageV4") or rec.get("MsgV4"),
            seg_num=rec.get("SegmentNumber") or rec.get("SegNum"),
            seg_fld=rec.get("SegmentField") or rec.get("SegFld"),
        )

    @staticmethod
    def _parse_segment(rec: Dict[str, Any]) -> IDocSegment:
        return IDocSegment(
            counter=int(rec.get("Counter") or rec.get("SegmentCounter") or 0),
            seg_name=str(rec.get("SegmentName") or rec.get("SegmentType") or ""),
            data=str(rec.get("SegmentData") or rec.get("Sdata") or ""),
        )
