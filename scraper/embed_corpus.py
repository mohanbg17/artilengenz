"""Embed intel.corpus rows into Pinecone (namespace: corpus_v1).

Mirrors raw_worker.py but for scraped corpus content (not live SAP errors).
Run after load_corpus.py finishes.

Usage:
    python embed_corpus.py
    python embed_corpus.py --batch-size 16
    python embed_corpus.py --reembed-all   # re-embed everything (e.g. after model upgrade)
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from datetime import datetime
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
log = structlog.get_logger("embed_corpus")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
VOYAGE_API_KEY     = os.getenv("VOYAGE_API_KEY", "")
VOYAGE_MODEL       = os.getenv("VOYAGE_MODEL", "voyage-3-large")
VOYAGE_DIM         = int(os.getenv("VOYAGE_DIM", "1024"))
PINECONE_API_KEY   = os.getenv("PINECONE_API_KEY", "")
PINECONE_INDEX     = os.getenv("PINECONE_INDEX", "artilegenz-sap-errors")
PINECONE_NAMESPACE = os.getenv("PINECONE_CORPUS_NAMESPACE", "corpus_v1")
PINECONE_CLOUD     = os.getenv("PINECONE_CLOUD", "aws")
PINECONE_REGION    = os.getenv("PINECONE_REGION", "us-east-1")

MAX_TEXT_CHARS = int(os.getenv("MAX_TEXT_CHARS", "100000"))

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Lazy imports
# ---------------------------------------------------------------------------
def _import_runtime():
    import voyageai
    from pinecone import Pinecone, ServerlessSpec
    from tenacity import retry, stop_after_attempt, wait_exponential
    return voyageai, Pinecone, ServerlessSpec, retry, stop_after_attempt, wait_exponential


# ---------------------------------------------------------------------------
# Pending fetch
# ---------------------------------------------------------------------------
PENDING_SQL = """
SELECT corpus_id, source_platform, source_url, source_id,
       error_signature, error_text, proposed_solution,
       system_module, upvote_score, accepted_flag,
       tags, view_count, question_score
FROM intel.corpus
WHERE %(reembed_all)s = TRUE
   OR embedding_version IS NULL
   OR embedding_version != %(target_version)s
ORDER BY upvote_score DESC NULLS LAST, scraped_at DESC NULLS LAST
LIMIT %(limit)s;
"""


def fetch_pending(limit: int, target_version: str, reembed_all: bool) -> List[Dict[str, Any]]:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                PENDING_SQL,
                {"limit": limit, "target_version": target_version, "reembed_all": reembed_all},
            )
            return [dict(r) for r in cur.fetchall()]


MARK_EMBEDDED_SQL = """
UPDATE intel.corpus
SET embedding_version = %(embedding_version)s,
    pinecone_id       = %(pinecone_id)s
WHERE corpus_id = %(corpus_id)s;
"""


def mark_embedded_batch(rows: List[Dict[str, Any]]) -> None:
    if not rows:
        return
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            psycopg2.extras.execute_batch(cur, MARK_EMBEDDED_SQL, rows)
            conn.commit()


# ---------------------------------------------------------------------------
# Text builder for corpus -- different shape than raw_errors
# ---------------------------------------------------------------------------
def build_corpus_embed_text(record: Dict[str, Any]) -> tuple[str, bool]:
    """Build the document text for Voyage. Includes signature, error context,
    AND the solution -- so retrieval surfaces the right answer for a given
    error pattern."""
    parts: List[str] = []
    parts.append(f"Module: {record.get('system_module', 'UNKNOWN')}")
    if record.get("error_signature"):
        parts.append(f"Signature: {record['error_signature']}")
    if record.get("tags"):
        tags = record["tags"] if isinstance(record["tags"], list) else []
        parts.append(f"Tags: {', '.join(tags)}")

    parts.append(f"\nError context:\n{record.get('error_text', '')}")
    parts.append(f"\nSolution:\n{record.get('proposed_solution', '')}")

    full = "\n".join(parts)
    if len(full) > MAX_TEXT_CHARS:
        return full[:MAX_TEXT_CHARS] + "\n[...truncated]", True
    return full, False


def build_corpus_metadata(record: Dict[str, Any]) -> Dict[str, Any]:
    md: Dict[str, Any] = {
        "corpus_id": record["corpus_id"],
        "source_platform": record.get("source_platform", ""),
        "source_url": record.get("source_url") or "",
        "system_module": record.get("system_module", ""),
        "upvote_score": int(record.get("upvote_score") or 0),
        "accepted_flag": bool(record.get("accepted_flag", False)),
        "view_count": int(record.get("view_count") or 0),
        "question_score": int(record.get("question_score") or 0),
    }
    if record.get("error_signature"):
        md["error_signature"] = str(record["error_signature"])[:500]
    if record.get("tags"):
        tags = record["tags"] if isinstance(record["tags"], list) else []
        md["tags"] = [t[:50] for t in tags[:10]]
    # Solution preview for quick UI display
    sol = record.get("proposed_solution") or ""
    if sol:
        md["solution_preview"] = sol[:1000]
    return md


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def run(batch_size: int, reembed_all: bool) -> None:
    if not VOYAGE_API_KEY or not PINECONE_API_KEY:
        sys.exit("Set VOYAGE_API_KEY and PINECONE_API_KEY env vars first")

    voyageai, Pinecone, ServerlessSpec, retry, stop_after_attempt, wait_exponential = _import_runtime()

    log.info(
        "embed_corpus_starting",
        voyage_model=VOYAGE_MODEL, pinecone_index=PINECONE_INDEX,
        namespace=PINECONE_NAMESPACE, batch_size=batch_size, reembed_all=reembed_all,
    )

    voyage = voyageai.Client(api_key=VOYAGE_API_KEY)
    pc = Pinecone(api_key=PINECONE_API_KEY)

    # Index should already exist (created by raw_worker). Verify.
    existing = [i["name"] for i in pc.list_indexes()]
    if PINECONE_INDEX not in existing:
        log.info("creating_pinecone_index", name=PINECONE_INDEX)
        pc.create_index(
            name=PINECONE_INDEX,
            dimension=VOYAGE_DIM,
            metric="cosine",
            spec=ServerlessSpec(cloud=PINECONE_CLOUD, region=PINECONE_REGION),
        )
        for _ in range(60):
            if pc.describe_index(PINECONE_INDEX).status.get("ready"):
                break
            time.sleep(1)
    index = pc.Index(PINECONE_INDEX)

    @retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=2, min=2, max=60), reraise=True)
    def embed_with_retry(texts: List[str]) -> List[List[float]]:
        result = voyage.embed(texts=texts, model=VOYAGE_MODEL, input_type="document")
        return result.embeddings

    total_embedded = 0
    while True:
        pending = fetch_pending(limit=batch_size, target_version=VOYAGE_MODEL, reembed_all=reembed_all)
        if not pending:
            log.info("no_more_pending", total_embedded=total_embedded)
            break

        log.info("batch_fetched", size=len(pending))

        texts: List[str] = []
        ids: List[str] = []
        metas: List[Dict[str, Any]] = []
        for rec in pending:
            text, _ = build_corpus_embed_text(rec)
            texts.append(text)
            ids.append(rec["corpus_id"])
            metas.append(build_corpus_metadata(rec))

        try:
            vectors = embed_with_retry(texts)
        except Exception as exc:  # noqa: BLE001
            log.exception("voyage_failed", err=str(exc))
            break

        items = [{"id": i, "values": v, "metadata": m} for i, v, m in zip(ids, vectors, metas)]
        try:
            index.upsert(vectors=items, namespace=PINECONE_NAMESPACE)
        except Exception as exc:  # noqa: BLE001
            log.exception("pinecone_upsert_failed", err=str(exc))
            break

        # Mark embedded in Postgres
        mark_rows = [
            {"embedding_version": VOYAGE_MODEL, "pinecone_id": cid, "corpus_id": cid}
            for cid in ids
        ]
        mark_embedded_batch(mark_rows)
        total_embedded += len(pending)
        log.info("batch_done", embedded=len(pending), running_total=total_embedded)

        # Once reembed_all is set, the SQL keeps returning everything in a loop.
        # Switch reembed_all off after the first pass so loop exits naturally.
        if reembed_all:
            reembed_all = False

    log.info("embed_corpus_complete", total=total_embedded)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=32)
    parser.add_argument("--reembed-all", action="store_true",
                        help="Re-embed everything (e.g. after model change)")
    args = parser.parse_args()
    run(batch_size=args.batch_size, reembed_all=args.reembed_all)


if __name__ == "__main__":
    main()
