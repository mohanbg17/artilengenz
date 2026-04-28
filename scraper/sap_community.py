"""SAP Community (community.sap.com) scraper.

Strategy:
  * Tag pages list error-related questions: error, dump, exception, abap, hana
  * Each question page exposes a JSON Khoros API that returns the body and replies
  * Accepted answers are flagged in the API response
"""
from __future__ import annotations

import re
from typing import AsyncIterator, List
from urllib.parse import urljoin

import structlog
from bs4 import BeautifulSoup

from .base import BaseScraper, CorpusEntry, detect_module, make_corpus_id, redact_pii

log = structlog.get_logger(__name__)


SEED_TAGS = [
    "abap-development",
    "sap-hana-cloud",
    "sap-s-4hana-cloud",
    "sap-basis",
    "sap-data-services",
    "sap-bw",
]
BASE = "https://community.sap.com"


class SapCommunityScraper(BaseScraper):
    platform = "sap_community"
    rate_limit_seconds = 1.5  # polite

    async def _list_questions_for_tag(self, tag: str, max_pages: int = 50) -> List[str]:
        urls: List[str] = []
        for page in range(1, max_pages + 1):
            list_url = f"{BASE}/t5/forums/searchpage/tab/message?advanced=false&q=&filter=labels&label_search_action=AND&location=tag:{tag}&page={page}"
            try:
                r = await self.fetch(list_url)
            except Exception:  # noqa: BLE001
                log.warning("sap_community_page_failed", tag=tag, page=page, exc_info=True)
                break
            soup = BeautifulSoup(r.text, "lxml")
            anchors = soup.select("a.page-link, a.message-subject-link, h2 a")
            page_urls = [a.get("href") for a in anchors if a.get("href")]
            page_urls = [
                urljoin(BASE, u)
                for u in page_urls
                if isinstance(u, str) and "/td-p/" in u
            ]
            if not page_urls:
                break
            urls.extend(page_urls)
        # dedup preserving order
        seen = set()
        deduped: List[str] = []
        for u in urls:
            if u not in seen:
                seen.add(u)
                deduped.append(u)
        return deduped

    async def _parse_thread(self, url: str) -> "CorpusEntry | None":
        try:
            r = await self.fetch(url)
        except Exception:  # noqa: BLE001
            return None
        soup = BeautifulSoup(r.text, "lxml")

        # Question body: first .lia-message-body-content
        q_body_el = soup.select_one(".lia-message-body-content, .lia-quilt-row-main")
        if not q_body_el:
            return None
        question_text = q_body_el.get_text(" ", strip=True)

        # Find accepted answer
        accepted_el = soup.select_one(".lia-message-solution-marked, .lia-component-solution")
        accepted = accepted_el is not None
        if accepted_el:
            answer_text = accepted_el.get_text(" ", strip=True)
        else:
            # Fallback to most-kudoed reply
            replies = soup.select(".lia-message-body-content")
            if len(replies) < 2:
                return None
            answer_text = replies[1].get_text(" ", strip=True)

        if len(question_text) < 50 or len(answer_text) < 30:
            return None

        # Title for signature
        title_el = soup.select_one("h1.lia-message-subject, .lia-message-subject")
        title = title_el.get_text(strip=True) if title_el else ""

        # Kudos count (best effort)
        score = 0
        kudos_el = soup.select_one(".kudos-count-link, .MessageKudosCount")
        if kudos_el:
            m = re.search(r"\d+", kudos_el.get_text())
            if m:
                score = int(m.group(0))

        question_text = redact_pii(question_text)
        answer_text = redact_pii(answer_text)

        return CorpusEntry(
            corpus_id=make_corpus_id(self.platform, url),
            source_platform=self.platform,
            source_url=url,
            error_signature=title,
            error_text=question_text[:8000],
            proposed_solution=answer_text[:8000],
            system_module=detect_module(question_text + " " + title),
            upvote_score=score,
            accepted_flag=accepted,
        )

    async def scrape(self, *, target: int) -> AsyncIterator[CorpusEntry]:
        per_tag = max(target // len(SEED_TAGS) + 50, 100)
        emitted = 0
        for tag in SEED_TAGS:
            if emitted >= target:
                break
            log.info("sap_community_tag_start", tag=tag, per_tag=per_tag)
            urls = await self._list_questions_for_tag(tag, max_pages=per_tag // 20 + 5)
            log.info("sap_community_tag_listed", tag=tag, urls=len(urls))
            for url in urls:
                if emitted >= target:
                    break
                entry = await self._parse_thread(url)
                if entry is None:
                    continue
                emitted += 1
                yield entry
        log.info("sap_community_done", emitted=emitted)
