"""Snowflake helpers for the intelligence layer.

Persists CorpusEntry to INTEL.CORPUS and provides a reader for new RAW_ERRORS
that haven't been classified yet.
"""
from __future__ import annotations

import asyncio
import json
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import structlog
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import serialization
import snowflake.connector

from config import settings
from .base import CorpusEntry

log = structlog.get_logger(__name__)


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


def _connect() -> snowflake.connector.SnowflakeConnection:
    return snowflake.connector.connect(
        account=settings.sf_account,
        user=settings.sf_user,
        private_key=_load_pkey(),
        warehouse=settings.sf_warehouse,
        database=settings.sf_database,
        schema=settings.sf_intel_schema,
        role=settings.sf_role,
    )


def _write_corpus_sync(entries: List[CorpusEntry]) -> int:
    conn = _connect()
    cur = conn.cursor()
    try:
        # Stage rows in TEMP table, MERGE on corpus_id
        tmp = f"TMP_CORPUS_{int(datetime.utcnow().timestamp() * 1000)}"
        cur.execute(f"CREATE TEMPORARY TABLE {tmp} LIKE CORPUS")
        rows = [
            (
                e.corpus_id, e.source_platform, e.source_url, e.error_signature,
                e.error_text, e.proposed_solution, e.system_module,
                e.upvote_score, e.accepted_flag, e.scraped_at,
                None,  # embedding_version filled by indexer
                None,  # pinecone_id filled by indexer
            )
            for e in entries
        ]
        placeholders = ",".join(["(" + ",".join(["%s"] * 12) + ")"] * len(rows))
        params: list = [v for row in rows for v in row]
        cur.execute(
            f"""
            INSERT INTO {tmp}
              (CORPUS_ID, SOURCE_PLATFORM, SOURCE_URL, ERROR_SIGNATURE, ERROR_TEXT,
               PROPOSED_SOLUTION, SYSTEM_MODULE, UPVOTE_SCORE, ACCEPTED_FLAG, SCRAPED_AT,
               EMBEDDING_VERSION, PINECONE_ID)
            VALUES {placeholders}
            """,
            params,
        )
        cur.execute(f"""
            MERGE INTO CORPUS tgt USING {tmp} src
            ON tgt.CORPUS_ID = src.CORPUS_ID
            WHEN NOT MATCHED THEN INSERT (
              CORPUS_ID, SOURCE_PLATFORM, SOURCE_URL, ERROR_SIGNATURE, ERROR_TEXT,
              PROPOSED_SOLUTION, SYSTEM_MODULE, UPVOTE_SCORE, ACCEPTED_FLAG, SCRAPED_AT,
              EMBEDDING_VERSION, PINECONE_ID
            ) VALUES (
              src.CORPUS_ID, src.SOURCE_PLATFORM, src.SOURCE_URL, src.ERROR_SIGNATURE, src.ERROR_TEXT,
              src.PROPOSED_SOLUTION, src.SYSTEM_MODULE, src.UPVOTE_SCORE, src.ACCEPTED_FLAG, src.SCRAPED_AT,
              src.EMBEDDING_VERSION, src.PINECONE_ID
            )
        """)
        inserted = cur.rowcount or 0
        conn.commit()
        cur.execute(f"DROP TABLE IF EXISTS {tmp}")
        return inserted
    finally:
        cur.close()
        conn.close()


async def write_corpus(entries: List[CorpusEntry]) -> int:
    """Async wrapper around blocking Snowflake call."""
    return await asyncio.to_thread(_write_corpus_sync, entries)


def fetch_corpus_for_embedding(*, only_unembedded: bool = True, limit: Optional[int] = None) -> List[dict]:
    conn = _connect()
    cur = conn.cursor(snowflake.connector.DictCursor)
    try:
        sql = """
          SELECT CORPUS_ID, ERROR_SIGNATURE, ERROR_TEXT, PROPOSED_SOLUTION,
                 SYSTEM_MODULE, UPVOTE_SCORE, ACCEPTED_FLAG, SOURCE_URL, SOURCE_PLATFORM
          FROM CORPUS
        """
        if only_unembedded:
            sql += " WHERE EMBEDDING_VERSION IS NULL"
        if limit:
            sql += f" LIMIT {int(limit)}"
        cur.execute(sql)
        return cur.fetchall()
    finally:
        cur.close()
        conn.close()


def mark_embedded(corpus_id: str, *, embedding_version: str, pinecone_id: str) -> None:
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute(
            "UPDATE CORPUS SET EMBEDDING_VERSION = %s, PINECONE_ID = %s WHERE CORPUS_ID = %s",
            (embedding_version, pinecone_id, corpus_id),
        )
        conn.commit()
    finally:
        cur.close()
        conn.close()


def fetch_unclassified_errors(limit: int = 500) -> List[dict]:
    """Pull RAW_ERRORS that don't yet have a row in CLASSIFICATIONS."""
    conn = _connect()
    cur = conn.cursor(snowflake.connector.DictCursor)
    try:
        cur.execute(f"""
            SELECT r.HASH_KEY, r.SOURCE, r.SYSTEM_ID, r.OCCURRED_AT, r.SEVERITY,
                   r.SHORT_TEXT, r.LONG_TEXT, r.TRANSACTION, r.PROGRAM, r.OBJECT, r.SUB_OBJECT
            FROM {settings.sf_database}.{settings.sf_raw_schema}.RAW_ERRORS r
            LEFT JOIN {settings.sf_database}.{settings.sf_intel_schema}.CLASSIFICATIONS c
              ON r.HASH_KEY = c.ERROR_HASH_KEY
            WHERE c.CLASSIFICATION_ID IS NULL
            ORDER BY r.OCCURRED_AT DESC
            LIMIT {int(limit)}
        """)
        return cur.fetchall()
    finally:
        cur.close()
        conn.close()


def write_classification(payload: dict) -> None:
    """Insert one row into CLASSIFICATIONS."""
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute(
            """
            INSERT INTO CLASSIFICATIONS (
              CLASSIFICATION_ID, ERROR_HASH_KEY, CREATED_AT, TOP_PROPOSAL_TITLE,
              COMPOSITE_CONFIDENCE, BADGE, RETRIED, PROPOSALS, CITATIONS,
              CRITIQUE_NOTES, MODEL_VERSION
            )
            SELECT %s, %s, %s, %s, %s, %s, %s, PARSE_JSON(%s), PARSE_JSON(%s), %s, %s
            """,
            (
                payload["classification_id"],
                payload["error_hash_key"],
                payload["created_at"],
                payload["top_proposal_title"],
                payload["composite_confidence"],
                payload["badge"],
                payload["retried"],
                json.dumps(payload["proposals"]),
                json.dumps(payload["citations"]),
                payload.get("critique_notes", ""),
                payload["model_version"],
            ),
        )
        conn.commit()
    finally:
        cur.close()
        conn.close()
