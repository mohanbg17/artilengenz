"""Corpus builder: runs all scrapers, dedups, filters, persists to Snowflake INTEL.CORPUS.

Volume targets per source:
  sap_community: 4500
  stack_overflow: 2500
  reddit: 1500
  blog: 1500
Total target: 10,000
"""
from __future__ import annotations

import asyncio
import logging
from typing import Dict, List

import click
import structlog

from config import settings
from .base import CorpusEntry
from .sap_community import SapCommunityScraper
from .stack_overflow import StackOverflowScraper
from .reddit_sap import RedditScraper
from .blog_crawler import BlogScraper

log = structlog.get_logger(__name__)

structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
    wrapper_class=structlog.make_filtering_bound_logger(getattr(logging, settings.log_level.upper(), logging.INFO)),
)


PER_SOURCE_TARGETS: Dict[str, int] = {
    "sap_community": 4500,
    "stack_overflow": 2500,
    "reddit": 1500,
    "blog": 1500,
}


def _quality_filter(entry: CorpusEntry) -> bool:
    """Drop low-quality entries before persisting."""
    if len(entry.error_text) < 50:
        return False
    if len(entry.proposed_solution) < 30:
        return False
    if entry.upvote_score < 0:
        return False
    return True


async def _run_one(scraper, target: int) -> List[CorpusEntry]:
    out: List[CorpusEntry] = []
    try:
        async for e in scraper.scrape(target=target):
            if _quality_filter(e):
                out.append(e)
    except Exception:  # noqa: BLE001
        log.exception("scraper_crashed", platform=scraper.platform)
    finally:
        await scraper.close()
    return out


async def _persist_to_snowflake(entries: List[CorpusEntry]) -> int:
    """Idempotent MERGE into INTEL.CORPUS by corpus_id."""
    from .sf_writer import write_corpus  # local import to avoid mandatory SF dep at parse time
    return await write_corpus(entries)


async def build(target: int) -> None:
    scrapers = [
        (SapCommunityScraper(), int(PER_SOURCE_TARGETS["sap_community"] * target / 10000)),
        (StackOverflowScraper(), int(PER_SOURCE_TARGETS["stack_overflow"] * target / 10000)),
        (RedditScraper(), int(PER_SOURCE_TARGETS["reddit"] * target / 10000)),
        (BlogScraper(), int(PER_SOURCE_TARGETS["blog"] * target / 10000)),
    ]
    log.info("corpus_build_start", target=target,
             plan={s.platform: t for s, t in scrapers})

    # Run scrapers concurrently
    results = await asyncio.gather(*[_run_one(s, t) for s, t in scrapers], return_exceptions=False)

    # Flatten, dedup by corpus_id
    seen: set[str] = set()
    flat: List[CorpusEntry] = []
    for batch in results:
        for e in batch:
            if e.corpus_id in seen:
                continue
            seen.add(e.corpus_id)
            flat.append(e)

    log.info("corpus_dedup_done", total=len(flat))

    if flat:
        inserted = await _persist_to_snowflake(flat)
        log.info("corpus_persisted", inserted=inserted)


@click.command()
@click.option("--target", type=int, default=10000, help="Total corpus target across all sources")
def main(target: int) -> None:
    asyncio.run(build(target))


if __name__ == "__main__":
    main()
