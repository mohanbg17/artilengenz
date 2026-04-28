"""Blog crawler — discovers SAP-focused posts via sitemaps and RSS feeds.

We use a curated seed list of high-signal blogs. For each, fetch the sitemap, filter
URLs that mention SAP/error/troubleshooting keywords, and parse the page.
"""
from __future__ import annotations

from typing import AsyncIterator, List
import xml.etree.ElementTree as ET

import structlog
from bs4 import BeautifulSoup

from .base import BaseScraper, CorpusEntry, detect_module, make_corpus_id, redact_pii

log = structlog.get_logger(__name__)


# Curated seed list - extend as needed
SEED_BLOGS = [
    "https://blogs.sap.com/sitemap.xml",
    "https://saplearners.com/sitemap.xml",
    "https://www.sapspot.com/sitemap.xml",
]
KEYWORDS = ("error", "dump", "fail", "issue", "troubleshoot", "fix", "resolve", "exception")


class BlogScraper(BaseScraper):
    platform = "blog"
    rate_limit_seconds = 2.0

    async def _read_sitemap(self, url: str) -> List[str]:
        try:
            r = await self.fetch(url)
        except Exception:  # noqa: BLE001
            return []
        try:
            root = ET.fromstring(r.text)
        except ET.ParseError:
            return []
        ns = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        urls = [u.text for u in root.findall("sm:url/sm:loc", ns)]
        # Sub-sitemaps?
        sub = [u.text for u in root.findall("sm:sitemap/sm:loc", ns)]
        for s in sub:
            if s:
                urls.extend(await self._read_sitemap(s))
        return [u for u in urls if u]

    async def _parse_post(self, url: str) -> "CorpusEntry | None":
        try:
            r = await self.fetch(url)
        except Exception:  # noqa: BLE001
            return None
        soup = BeautifulSoup(r.text, "lxml")
        # Strip nav/footer/aside
        for tag in soup.select("nav, footer, aside, script, style, .navigation, .sidebar"):
            tag.decompose()

        title_el = soup.select_one("h1, title")
        title = title_el.get_text(strip=True) if title_el else ""

        # Best-effort body: <article> or <main> or largest <div class*="content">
        body_el = (
            soup.select_one("article")
            or soup.select_one("main")
            or soup.select_one("[class*='content']")
            or soup.body
        )
        if body_el is None:
            return None
        body_text = body_el.get_text(" ", strip=True)
        if len(body_text) < 200:
            return None

        # Heuristic: if body contains "Solution:" / "Resolution:" split on it
        for marker in ["Resolution:", "Solution:", "Fix:", "Workaround:", "How to fix"]:
            if marker.lower() in body_text.lower():
                idx = body_text.lower().index(marker.lower())
                problem = body_text[:idx]
                solution = body_text[idx + len(marker):]
                if len(problem) > 100 and len(solution) > 50:
                    return CorpusEntry(
                        corpus_id=make_corpus_id(self.platform, url),
                        source_platform=self.platform,
                        source_url=url,
                        error_signature=title,
                        error_text=redact_pii(problem[:8000]),
                        proposed_solution=redact_pii(solution[:8000]),
                        system_module=detect_module(title + " " + problem),
                        upvote_score=0,
                        accepted_flag=False,
                    )
        # No clear split → treat full body as solution writeup keyed by title
        return CorpusEntry(
            corpus_id=make_corpus_id(self.platform, url),
            source_platform=self.platform,
            source_url=url,
            error_signature=title,
            error_text=redact_pii(title)[:1000],
            proposed_solution=redact_pii(body_text[:8000]),
            system_module=detect_module(title + " " + body_text[:1000]),
            upvote_score=0,
            accepted_flag=False,
        )

    async def scrape(self, *, target: int) -> AsyncIterator[CorpusEntry]:
        emitted = 0
        for sitemap in SEED_BLOGS:
            if emitted >= target:
                break
            urls = await self._read_sitemap(sitemap)
            urls = [u for u in urls if any(k in u.lower() for k in KEYWORDS)]
            log.info("blog_sitemap_filtered", sitemap=sitemap, count=len(urls))
            for url in urls:
                if emitted >= target:
                    break
                entry = await self._parse_post(url)
                if entry is None:
                    continue
                emitted += 1
                yield entry
        log.info("blog_done", emitted=emitted)
