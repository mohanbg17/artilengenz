"""Load scraped SO records into intel.corpus.

Takes the JSON output from fetch_so.py, normalizes into intel.corpus schema,
generates corpus_id (sha256 of source_platform + source_id), and inserts
with ON CONFLICT (source_platform, source_id) DO UPDATE so re-runs are idempotent.

Usage:
    python load_corpus.py --input scraped_so.json
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import logging
import os
import re
import sys
from datetime import datetime, timezone
from typing import Any, Dict, List

import psycopg2
import psycopg2.extras
import structlog

logging.basicConfig(format="%(message)s", stream=sys.stdout, level=logging.INFO)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.dev.ConsoleRenderer(),
    ],
)
log = structlog.get_logger("load_corpus")


PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Text cleaning
# ---------------------------------------------------------------------------
def strip_html(text: str) -> str:
    """Remove HTML tags. SO bodies may contain <p>, <code>, <pre>, etc."""
    if not text:
        return ""
    # Replace <br>, <p>, </p> with newlines
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?p[^>]*>", "\n", text, flags=re.IGNORECASE)
    # Strip remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode HTML entities
    text = html.unescape(text)
    # Collapse whitespace
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def extract_error_signature(question_body: str, title: str) -> str:
    """Best-effort signature extraction from question content.

    Looks for ABAP exception classes (CX_*), runtime errors (UPPERCASE_WITH_UNDERSCORES),
    or quoted error strings. Falls back to first 200 chars of title.
    """
    sig_patterns = [
        r"\b(CX_[A-Z_]+(?:_[A-Z]+)*)\b",       # CX_SY_ZERODIVIDE
        r"\b([A-Z]{3,}_[A-Z_]+)\b",            # COMPUTE_INT_ZERODIVIDE
        r"\b(SY-SUBRC\s*=\s*\d+)\b",            # SY-SUBRC = 4
        r"\b(MESSAGE\s+E?\d+\s*\([^)]+\))",    # MESSAGE E000(SY)
    ]
    for pat in sig_patterns:
        m = re.search(pat, question_body or "", re.IGNORECASE)
        if m:
            return m.group(1)[:512]
    # Fallback: title
    return (title or "")[:512]


def make_corpus_id(source_platform: str, source_id: str) -> str:
    raw = f"{source_platform}:{source_id}"
    return hashlib.sha256(raw.encode()).hexdigest()


# ---------------------------------------------------------------------------
# DB upsert
# ---------------------------------------------------------------------------
UPSERT_SQL = """
INSERT INTO intel.corpus (
  corpus_id, source_platform, source_url, source_id,
  error_signature, error_text, proposed_solution,
  system_module, upvote_score, accepted_flag,
  scraped_at, fetched_at,
  tags, view_count, answer_count, question_score, language
)
VALUES (
  %(corpus_id)s, %(source_platform)s, %(source_url)s, %(source_id)s,
  %(error_signature)s, %(error_text)s, %(proposed_solution)s,
  %(system_module)s, %(upvote_score)s, %(accepted_flag)s,
  %(scraped_at)s, %(fetched_at)s,
  %(tags)s, %(view_count)s, %(answer_count)s, %(question_score)s, %(language)s
)
ON CONFLICT (source_platform, source_id) WHERE source_id IS NOT NULL
DO UPDATE SET
  error_signature   = EXCLUDED.error_signature,
  error_text        = EXCLUDED.error_text,
  proposed_solution = EXCLUDED.proposed_solution,
  system_module     = EXCLUDED.system_module,
  upvote_score      = EXCLUDED.upvote_score,
  accepted_flag     = EXCLUDED.accepted_flag,
  scraped_at        = EXCLUDED.scraped_at,
  fetched_at        = EXCLUDED.fetched_at,
  tags              = EXCLUDED.tags,
  view_count        = EXCLUDED.view_count,
  answer_count      = EXCLUDED.answer_count,
  question_score    = EXCLUDED.question_score,
  -- Reset embedding on content change so embedder picks it up again
  embedding_version = NULL,
  pinecone_id       = NULL;
"""


def to_corpus_row(rec: Dict[str, Any]) -> Dict[str, Any]:
    source_platform = rec["source_platform"]
    source_id = rec["source_id"]
    title = rec.get("title", "")
    q_body = strip_html(rec.get("question_body", ""))
    a_body = strip_html(rec.get("answer_body", ""))

    # error_text is the question body + title context for retrieval
    error_text = f"{title}\n\n{q_body}" if title else q_body

    # creation_date_unix -> scraped_at (when SO published the question)
    scraped_at = None
    if rec.get("creation_date_unix"):
        scraped_at = datetime.fromtimestamp(
            rec["creation_date_unix"], tz=timezone.utc
        ).replace(tzinfo=None)

    fetched_at = None
    if rec.get("fetched_at"):
        try:
            fetched_at = datetime.fromisoformat(rec["fetched_at"].replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            fetched_at = datetime.utcnow()

    return {
        "corpus_id": make_corpus_id(source_platform, source_id),
        "source_platform": source_platform,
        "source_url": rec.get("source_url"),
        "source_id": source_id,
        "error_signature": extract_error_signature(q_body, title),
        "error_text": error_text[:50000],
        "proposed_solution": a_body[:50000],
        "system_module": rec.get("module") or "UNKNOWN",
        "upvote_score": rec.get("answer_score", 0),
        "accepted_flag": bool(rec.get("is_accepted", False)),
        "scraped_at": scraped_at,
        "fetched_at": fetched_at or datetime.utcnow(),
        "tags": rec.get("tags") or [],
        "view_count": rec.get("view_count", 0),
        "answer_count": rec.get("answer_count", 0),
        "question_score": rec.get("question_score", 0),
        "language": "en",
    }


def load(input_path: str) -> int:
    with open(input_path, "r", encoding="utf-8") as f:
        records = json.load(f)

    log.info("load_starting", input=input_path, raw_count=len(records))

    rows = [to_corpus_row(r) for r in records]
    log.info("normalized", count=len(rows))

    inserted = 0
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            psycopg2.extras.execute_batch(cur, UPSERT_SQL, rows, page_size=50)
            inserted = cur.rowcount
            conn.commit()

    log.info("load_complete", inserted_or_updated=inserted)
    return inserted


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="JSON file from fetch_so.py")
    args = parser.parse_args()
    load(args.input)


if __name__ == "__main__":
    main()
