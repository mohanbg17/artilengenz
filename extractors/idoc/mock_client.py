"""Mock IDoc connector backed by JSON fixtures.

Used by tests and local development without a live SAP system. Reads JSON
fixtures from a directory, parses them through the same OData parser used in
production so the fixtures double as wire-format documentation.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence

from .connector import IDocConnector
from .models import IDoc
from .odata_client import ODataIDocClient


class MockIDocClient(IDocConnector):
    """Read IDocs from a directory of OData-shaped JSON fixtures.

    Each fixture must be either:
      * a single entity dict (representing one IDoc), or
      * a v2-style ``{"d": {"results": [...]}}`` payload, or
      * a v4-style ``{"value": [...]}`` payload.

    The parser matches the production OData parser so fixtures stay realistic.
    """

    def __init__(self, fixtures_dir: Path | str) -> None:
        self.fixtures_dir = Path(fixtures_dir)
        if not self.fixtures_dir.is_dir():
            raise FileNotFoundError(f"fixtures dir not found: {self.fixtures_dir}")
        self._parser = ODataIDocClient(http_client=_NoopHttp())  # only used for parse helpers

    def list_failed_idocs(
        self,
        *,
        since: Optional[datetime] = None,
        statuses: Optional[Sequence[str]] = None,
        max_results: int = 100,
    ) -> List[IDoc]:
        all_records: List[IDoc] = []
        for path in sorted(self.fixtures_dir.glob("*.json")):
            payload = json.loads(path.read_text(encoding="utf-8"))
            for rec in self._records_in(payload):
                all_records.append(self._parser._parse_header(rec))

        filtered = [
            d for d in all_records
            if (since is None or d.created_at > since)
            and (not statuses or d.status in set(statuses))
        ]
        filtered.sort(key=lambda d: (d.created_at, d.docnum))
        return filtered[:max_results]

    @staticmethod
    def _records_in(payload):
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            if "value" in payload and isinstance(payload["value"], list):
                return payload["value"]
            d = payload.get("d")
            if isinstance(d, dict) and isinstance(d.get("results"), list):
                return d["results"]
            if isinstance(d, list):
                return d
            return [payload]
        return []


class _NoopHttp:
    """Stand-in httpx client for the mock — never called, parser-only path."""

    def get(self, *args, **kwargs):
        raise RuntimeError("MockIDocClient must not perform HTTP requests")
