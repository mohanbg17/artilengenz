"""Watermark store for IDoc polling.

Tracks the last successfully ingested ``occurred_at`` per (source, system_id)
in the existing ``raw.watermarks`` table. The DB-backed implementation uses
the platform's existing backend abstraction (``db.get_backend``) so we don't
hard-code Postgres SQL into the extractor.
"""
from __future__ import annotations

from datetime import datetime
from typing import Dict, Optional, Protocol, Tuple


class WatermarkStore(Protocol):
    def get(self, source: str, system_id: str) -> Optional[datetime]: ...
    def set(self, source: str, system_id: str, last_occurred_at: datetime) -> None: ...


class MemoryWatermarkStore:
    """In-memory store for tests and dry-runs."""

    def __init__(self) -> None:
        self._wm: Dict[Tuple[str, str], datetime] = {}

    def get(self, source: str, system_id: str) -> Optional[datetime]:
        return self._wm.get((source, system_id))

    def set(self, source: str, system_id: str, last_occurred_at: datetime) -> None:
        self._wm[(source, system_id)] = last_occurred_at


class DBWatermarkStore:
    """Watermark store backed by ``raw.watermarks``.

    Imports ``db.get_backend`` lazily so unit tests that don't touch the DB
    can construct the poller without psycopg / a live Postgres.
    """

    def __init__(self, backend=None) -> None:
        if backend is None:
            from db import get_backend
            backend = get_backend()
        self.backend = backend

    def get(self, source: str, system_id: str) -> Optional[datetime]:
        sql = (
            f"SELECT last_occurred_at FROM {self.backend.watermarks_table} "
            f"WHERE source = %s AND system_id = %s"
        )
        with self.backend.cursor() as cur:
            cur.execute(sql, (source, system_id))
            row = cur.fetchone()
        if not row:
            return None
        return row[0] if not isinstance(row, dict) else row.get("last_occurred_at")

    def set(self, source: str, system_id: str, last_occurred_at: datetime) -> None:
        sql = (
            f"INSERT INTO {self.backend.watermarks_table} "
            f"(source, system_id, last_occurred_at, updated_at) "
            f"VALUES (%s, %s, %s, CURRENT_TIMESTAMP) "
            f"ON CONFLICT (source, system_id) DO UPDATE "
            f"  SET last_occurred_at = EXCLUDED.last_occurred_at, "
            f"      updated_at = CURRENT_TIMESTAMP"
        )
        with self.backend.cursor() as cur:
            cur.execute(sql, (source, system_id, last_occurred_at))
            cur.connection.commit()
