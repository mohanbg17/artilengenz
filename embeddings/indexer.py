"""Embedding indexer.

Reads INTEL.CORPUS, embeds via Voyage, uploads to Pinecone, marks rows as embedded.
"""
from __future__ import annotations

import logging
from typing import Dict, List

import click
import structlog

from config import settings
from scraper.sf_writer import fetch_corpus_for_embedding, mark_embedded
from .pinecone_index import get_index
from .voyage_client import build_corpus_embedding_text, get_voyage

log = structlog.get_logger(__name__)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
    wrapper_class=structlog.make_filtering_bound_logger(getattr(logging, settings.log_level.upper(), logging.INFO)),
)

EMBED_BATCH = 64


def _build_inputs(rows: List[Dict]) -> tuple[List[str], List[str], List[Dict]]:
    ids: List[str] = []
    texts: List[str] = []
    metas: List[Dict] = []
    for r in rows:
        text = build_corpus_embedding_text(
            signature=r.get("ERROR_SIGNATURE"),
            error_text=r.get("ERROR_TEXT", ""),
            solution=r.get("PROPOSED_SOLUTION", ""),
            module=r.get("SYSTEM_MODULE"),
        )
        pinecone_id = r["CORPUS_ID"]
        ids.append(pinecone_id)
        texts.append(text)
        metas.append({
            "corpus_id": r["CORPUS_ID"],
            "source_platform": r.get("SOURCE_PLATFORM"),
            "source_url": r.get("SOURCE_URL"),
            "system_module": r.get("SYSTEM_MODULE") or "",
            "upvote_score": int(r.get("UPVOTE_SCORE") or 0),
            "accepted_flag": bool(r.get("ACCEPTED_FLAG")),
            "error_signature": (r.get("ERROR_SIGNATURE") or "")[:500],
            "solution_excerpt": (r.get("PROPOSED_SOLUTION") or "")[:1000],
        })
    return ids, texts, metas


def run(*, reindex_all: bool, limit: int | None) -> None:
    voyage = get_voyage()
    index = get_index()

    rows = fetch_corpus_for_embedding(only_unembedded=not reindex_all, limit=limit)
    log.info("indexer_start", rows=len(rows))

    for i in range(0, len(rows), EMBED_BATCH):
        batch = rows[i : i + EMBED_BATCH]
        ids, texts, metas = _build_inputs(batch)
        try:
            vectors = voyage.embed_documents(texts)
        except Exception:  # noqa: BLE001
            log.exception("voyage_batch_failed", offset=i)
            continue
        try:
            index.upsert_batch(ids=ids, vectors=vectors, metadatas=metas)
        except Exception:  # noqa: BLE001
            log.exception("pinecone_upsert_failed", offset=i)
            continue
        for cid in ids:
            mark_embedded(cid, embedding_version=settings.voyage_model, pinecone_id=cid)
        log.info("indexer_batch_done", offset=i, count=len(batch))

    log.info("indexer_done")


@click.command()
@click.option("--reindex-all", is_flag=True, help="Re-embed everything (ignore EMBEDDING_VERSION)")
@click.option("--limit", type=int, default=None)
def main(reindex_all: bool, limit: int | None) -> None:
    run(reindex_all=reindex_all, limit=limit)


if __name__ == "__main__":
    main()
