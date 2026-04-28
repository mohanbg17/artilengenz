"""Database access for the raw-error embedding worker.

Wraps Postgres connections, provides SELECT for pending records, and
UPSERT for the embedding tracking log. Uses psycopg2 with autocommit
disabled and explicit COMMITs for batch atomicity.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterator, List, Optional

import psycopg2
import psycopg2.extras
import structlog

log = structlog.get_logger(__name__)


# ---------------------------------------------------------------------------
# Connection helpers
# ---------------------------------------------------------------------------

def _build_dsn(
    host: str, port: int, user: str, password: str, database: str
) -> str:
    return (
        f"host={host} port={port} user={user} "
        f"password={password} dbname={database}"
    )


@contextmanager
def get_conn(dsn: str) -> Iterator[psycopg2.extensions.connection]:
    """Yield a Postgres connection. Always closed on exit."""
    conn = psycopg2.connect(dsn)
    try:
        yield conn
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Pending-record query
# ---------------------------------------------------------------------------

PENDING_SQL = """
SELECT
  r.hash_key,
  r.source,
  r.system_id,
  r.occurred_at,
  r.severity,
  r.short_text,
  r.long_text,
  r.error_id,
  r.transaction,
  r.program,
  r.user_name,
  r.instance,
  COALESCE(l.retry_count, 0) AS retry_count
FROM raw.raw_errors r
LEFT JOIN intel.raw_embeddings_log l
  ON l.hash_key = r.hash_key
 AND l.embedding_version = %(target_version)s
WHERE
  l.hash_key IS NULL
  OR (l.embed_error IS NOT NULL AND COALESCE(l.retry_count, 0) < %(max_retries)s)
ORDER BY r.occurred_at ASC
LIMIT %(limit)s;
"""


def fetch_pending_records(
    dsn: str,
    target_version: str,
    limit: int = 32,
    max_retries: int = 3,
) -> List[Dict[str, Any]]:
    """Return up to `limit` records that need embedding under the current
    target version, including those that failed but have retries left."""
    with get_conn(dsn) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                PENDING_SQL,
                {
                    "target_version": target_version,
                    "max_retries": max_retries,
                    "limit": limit,
                },
            )
            return [dict(row) for row in cur.fetchall()]


# ---------------------------------------------------------------------------
# Log writers
# ---------------------------------------------------------------------------

UPSERT_SUCCESS_SQL = """
INSERT INTO intel.raw_embeddings_log (
  hash_key, embedding_version, pinecone_id, pinecone_namespace,
  embedded_at, text_length, truncated, embed_error,
  retry_count, last_attempt_at
)
VALUES (
  %(hash_key)s, %(embedding_version)s, %(pinecone_id)s, %(pinecone_namespace)s,
  CURRENT_TIMESTAMP, %(text_length)s, %(truncated)s, NULL,
  0, CURRENT_TIMESTAMP
)
ON CONFLICT (hash_key) DO UPDATE SET
  embedding_version  = EXCLUDED.embedding_version,
  pinecone_id        = EXCLUDED.pinecone_id,
  pinecone_namespace = EXCLUDED.pinecone_namespace,
  embedded_at        = CURRENT_TIMESTAMP,
  text_length        = EXCLUDED.text_length,
  truncated          = EXCLUDED.truncated,
  embed_error        = NULL,
  retry_count        = 0,
  last_attempt_at    = CURRENT_TIMESTAMP;
"""

UPSERT_FAILURE_SQL = """
INSERT INTO intel.raw_embeddings_log (
  hash_key, embedding_version, pinecone_id, pinecone_namespace,
  embed_error, retry_count, last_attempt_at
)
VALUES (
  %(hash_key)s, %(embedding_version)s, '', '',
  %(embed_error)s, 1, CURRENT_TIMESTAMP
)
ON CONFLICT (hash_key) DO UPDATE SET
  embed_error     = EXCLUDED.embed_error,
  retry_count     = intel.raw_embeddings_log.retry_count + 1,
  last_attempt_at = CURRENT_TIMESTAMP;
"""


def log_success_batch(
    dsn: str,
    rows: List[Dict[str, Any]],
) -> int:
    """Upsert success rows. Returns row count."""
    if not rows:
        return 0
    with get_conn(dsn) as conn:
        with conn.cursor() as cur:
            psycopg2.extras.execute_batch(cur, UPSERT_SUCCESS_SQL, rows)
            conn.commit()
    return len(rows)


def log_failure(
    dsn: str,
    hash_key: str,
    embedding_version: str,
    embed_error: str,
) -> None:
    """Mark a single record as failed, increment retry count."""
    with get_conn(dsn) as conn:
        with conn.cursor() as cur:
            cur.execute(
                UPSERT_FAILURE_SQL,
                {
                    "hash_key": hash_key,
                    "embedding_version": embedding_version,
                    "embed_error": (embed_error or "")[:2000],
                },
            )
            conn.commit()


# ---------------------------------------------------------------------------
# Status
# ---------------------------------------------------------------------------

STATUS_SQL = """
SELECT embedded, pending, failed, total
FROM intel.v_embedding_status;
"""


def fetch_status(dsn: str) -> Dict[str, int]:
    with get_conn(dsn) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(STATUS_SQL)
            row = cur.fetchone()
            return dict(row) if row else {"embedded": 0, "pending": 0, "failed": 0, "total": 0}


# ---------------------------------------------------------------------------
# Convenience: build DSN from environment
# ---------------------------------------------------------------------------

def dsn_from_env() -> str:
    import os

    return _build_dsn(
        host=os.getenv("PG_HOST", "localhost"),
        port=int(os.getenv("PG_PORT", "5432")),
        user=os.getenv("PG_USER", "artilegenz"),
        password=os.getenv("PG_PASSWORD", "artilegenz_local_dev"),
        database=os.getenv("PG_DATABASE", "sap_errors"),
    )
