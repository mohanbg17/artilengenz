"""Database backend abstraction.

The application talks to Postgres or Snowflake via a single interface.
Selection is driven by `DB_ENGINE` in `.env` (postgres | snowflake).

Why this exists: we want demo-ability today (Postgres on laptop) without
forking the codebase. Same SQL where possible; engine-specific where not.

Usage::

    from db.backend import get_backend
    with get_backend().cursor() as cur:
        cur.execute("SELECT 1")
        print(cur.fetchone())
"""
from __future__ import annotations

import json
from abc import ABC, abstractmethod
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Sequence

import structlog

from config import settings

log = structlog.get_logger(__name__)


# ── Public abstraction ──────────────────────────────────────────────────────

class DBBackend(ABC):
    """Minimal interface every backend must implement."""

    name: str = "abstract"

    @abstractmethod
    def connect(self) -> Any: ...

    @contextmanager
    def cursor(self, *, dict_rows: bool = False) -> Iterator[Any]:
        conn = self.connect()
        cur = self._make_cursor(conn, dict_rows=dict_rows)
        try:
            yield cur
        finally:
            cur.close()
            conn.close()

    @abstractmethod
    def _make_cursor(self, conn: Any, *, dict_rows: bool) -> Any: ...

    # ── SQL fragments that differ between engines ───────────────────────────

    @property
    @abstractmethod
    def raw_schema(self) -> str: ...

    @property
    @abstractmethod
    def intel_schema(self) -> str: ...

    @property
    @abstractmethod
    def raw_errors_table(self) -> str: ...

    @property
    @abstractmethod
    def corpus_table(self) -> str: ...

    @property
    @abstractmethod
    def classifications_table(self) -> str: ...

    @property
    @abstractmethod
    def feedback_table(self) -> str: ...

    @property
    @abstractmethod
    def watermarks_table(self) -> str: ...

    @abstractmethod
    def upsert_raw_errors(self, rows: List[Dict[str, Any]]) -> int:
        """Idempotent insert into raw_errors. Returns count of new rows."""

    @abstractmethod
    def upsert_corpus(self, rows: List[Dict[str, Any]]) -> int: ...

    @abstractmethod
    def insert_classification(self, payload: Dict[str, Any]) -> None: ...

    @abstractmethod
    def insert_feedback(self, payload: Dict[str, Any]) -> None: ...


# ── Postgres ────────────────────────────────────────────────────────────────

class PostgresBackend(DBBackend):
    name = "postgres"

    def __init__(self) -> None:
        # Lazy import so Snowflake-only deployments don't need psycopg
        import psycopg  # noqa: F401
        self._psycopg = psycopg

    def connect(self) -> Any:
        return self._psycopg.connect(
            host=settings.pg_host,
            port=settings.pg_port,
            user=settings.pg_user,
            password=settings.pg_password,
            dbname=settings.pg_database,
        )

    def _make_cursor(self, conn: Any, *, dict_rows: bool) -> Any:
        if dict_rows:
            from psycopg.rows import dict_row
            return conn.cursor(row_factory=dict_row)
        return conn.cursor()

    @property
    def raw_schema(self) -> str:
        return "raw"

    @property
    def intel_schema(self) -> str:
        return "intel"

    @property
    def raw_errors_table(self) -> str:
        return "raw.raw_errors"

    @property
    def corpus_table(self) -> str:
        return "intel.corpus"

    @property
    def classifications_table(self) -> str:
        return "intel.classifications"

    @property
    def feedback_table(self) -> str:
        return "intel.human_feedback"

    @property
    def watermarks_table(self) -> str:
        return "raw.watermarks"

    def upsert_raw_errors(self, rows: List[Dict[str, Any]]) -> int:
        if not rows:
            return 0
        cols = [
            "hash_key", "source", "system_id", "occurred_at", "extracted_at", "severity",
            "short_text", "long_text", "error_id", "user_name", "transaction", "program",
            "instance", "work_process", "object", "sub_object", "job_name", "raw",
        ]
        with self.cursor() as cur:
            placeholders = "(" + ",".join(["%s"] * len(cols)) + ")"
            values_sql = ",".join([placeholders] * len(rows))
            params: List[Any] = []
            for r in rows:
                for c in cols:
                    v = r.get(c.upper(), r.get(c))
                    # JSONB column: ensure we pass dict/None, not pre-stringified JSON
                    if c == "raw" and isinstance(v, str):
                        try:
                            v = json.loads(v)
                        except (TypeError, json.JSONDecodeError):
                            v = {"abap_raw": v}
                    if c == "raw":
                        v = self._psycopg.types.json.Jsonb(v) if v is not None else None
                    params.append(v)

            sql = f"""
                INSERT INTO {self.raw_errors_table}
                  ({", ".join(cols)})
                VALUES {values_sql}
                ON CONFLICT (hash_key) DO NOTHING
            """
            cur.execute(sql, params)
            inserted = cur.rowcount or 0
            cur.connection.commit()
            return inserted

    def upsert_corpus(self, rows: List[Dict[str, Any]]) -> int:
        if not rows:
            return 0
        cols = [
            "corpus_id", "source_platform", "source_url", "error_signature",
            "error_text", "proposed_solution", "system_module",
            "upvote_score", "accepted_flag", "scraped_at",
            "embedding_version", "pinecone_id",
        ]
        with self.cursor() as cur:
            placeholders = "(" + ",".join(["%s"] * len(cols)) + ")"
            values_sql = ",".join([placeholders] * len(rows))
            params: List[Any] = []
            for r in rows:
                for c in cols:
                    params.append(r.get(c.upper(), r.get(c)))
            sql = f"""
                INSERT INTO {self.corpus_table}
                  ({", ".join(cols)})
                VALUES {values_sql}
                ON CONFLICT (corpus_id) DO NOTHING
            """
            cur.execute(sql, params)
            inserted = cur.rowcount or 0
            cur.connection.commit()
            return inserted

    def insert_classification(self, payload: Dict[str, Any]) -> None:
        with self.cursor() as cur:
            cur.execute(
                f"""
                INSERT INTO {self.classifications_table} (
                  classification_id, error_hash_key, created_at, top_proposal_title,
                  composite_confidence, badge, retried, proposals, citations,
                  critique_notes, model_version
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    payload["classification_id"],
                    payload["error_hash_key"],
                    payload["created_at"],
                    payload["top_proposal_title"],
                    payload["composite_confidence"],
                    payload["badge"],
                    payload["retried"],
                    self._psycopg.types.json.Jsonb(payload["proposals"]),
                    self._psycopg.types.json.Jsonb(payload["citations"]),
                    payload.get("critique_notes", ""),
                    payload["model_version"],
                ),
            )
            cur.connection.commit()

    def insert_feedback(self, payload: Dict[str, Any]) -> None:
        with self.cursor() as cur:
            cur.execute(
                f"""
                INSERT INTO {self.feedback_table} (
                  feedback_id, classification_id, reviewer, reviewed_at,
                  accepted, corrected_solution, comments
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    payload["feedback_id"],
                    payload["classification_id"],
                    payload["reviewer"],
                    payload["reviewed_at"],
                    payload["accepted"],
                    payload.get("corrected_solution"),
                    payload.get("comments"),
                ),
            )
            cur.connection.commit()


# ── Snowflake (kept for portability) ────────────────────────────────────────

class SnowflakeBackend(DBBackend):
    name = "snowflake"

    def __init__(self) -> None:
        import snowflake.connector  # noqa: F401
        self._snowflake = __import__("snowflake.connector", fromlist=["connector"])

    def _load_pkey(self) -> bytes:
        from cryptography.hazmat.backends import default_backend
        from cryptography.hazmat.primitives import serialization
        p = Path(settings.sf_private_key_path).expanduser()
        with open(p, "rb") as fh:
            passphrase = settings.sf_private_key_passphrase.encode() if settings.sf_private_key_passphrase else None
            key = serialization.load_pem_private_key(fh.read(), password=passphrase, backend=default_backend())
        return key.private_bytes(
            encoding=serialization.Encoding.DER,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )

    def connect(self) -> Any:
        return self._snowflake.connect(
            account=settings.sf_account,
            user=settings.sf_user,
            private_key=self._load_pkey(),
            warehouse=settings.sf_warehouse,
            database=settings.sf_database,
            schema=settings.sf_intel_schema,
            role=settings.sf_role,
        )

    def _make_cursor(self, conn: Any, *, dict_rows: bool) -> Any:
        if dict_rows:
            return conn.cursor(self._snowflake.DictCursor)
        return conn.cursor()

    @property
    def raw_schema(self) -> str:
        return settings.sf_raw_schema

    @property
    def intel_schema(self) -> str:
        return settings.sf_intel_schema

    @property
    def raw_errors_table(self) -> str:
        return f"{settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS"

    @property
    def corpus_table(self) -> str:
        return f"{settings.sf_database}.{settings.sf_intel_schema}.CORPUS"

    @property
    def classifications_table(self) -> str:
        return f"{settings.sf_database}.{settings.sf_intel_schema}.CLASSIFICATIONS"

    @property
    def feedback_table(self) -> str:
        return f"{settings.sf_database}.{settings.sf_intel_schema}.HUMAN_FEEDBACK"

    @property
    def watermarks_table(self) -> str:
        return f"{settings.sf_database}.{settings.sf_raw_schema}.WATERMARKS"

    def upsert_raw_errors(self, rows: List[Dict[str, Any]]) -> int:
        # Snowflake path uses MERGE via temp table — kept in api/ingest.py for now.
        # When enabling Snowflake, lift that logic here. For the demo this method is unused.
        raise NotImplementedError(
            "SnowflakeBackend.upsert_raw_errors is not migrated yet. "
            "Use Postgres for the current demo, or restore the original ingest.py MERGE path."
        )

    def upsert_corpus(self, rows: List[Dict[str, Any]]) -> int:
        raise NotImplementedError("Snowflake corpus upsert not yet migrated to backend abstraction.")

    def insert_classification(self, payload: Dict[str, Any]) -> None:
        raise NotImplementedError

    def insert_feedback(self, payload: Dict[str, Any]) -> None:
        raise NotImplementedError


# ── Factory ─────────────────────────────────────────────────────────────────

_singleton: Optional[DBBackend] = None


def get_backend() -> DBBackend:
    """Return the configured backend (cached)."""
    global _singleton
    if _singleton is not None:
        return _singleton
    engine = (settings.db_engine or "postgres").lower()
    if engine == "postgres":
        _singleton = PostgresBackend()
    elif engine == "snowflake":
        _singleton = SnowflakeBackend()
    else:
        raise ValueError(f"Unknown DB_ENGINE: {engine!r}. Use 'postgres' or 'snowflake'.")
    log.info("db_backend_selected", engine=_singleton.name)
    return _singleton
