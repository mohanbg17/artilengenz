"""Stack Overflow scraper using the Stack Exchange API v2.3.

Uses public API (no key needed for low volume). For 2,500 entries we'll respect the
quota by paging at the max page size of 100.

Tags: abap, sap, sap-erp, sap-hana, sap-cloud-platform.
"""
from __future__ import annotations

from typing import AsyncIterator, List

import structlog

from .base import BaseScraper, CorpusEntry, detect_module, make_corpus_id, redact_pii

log = structlog.get_logger(__name__)

API = "https://api.stackexchange.com/2.3"
TAGS = ["abap", "sap", "sap-erp", "sap-hana", "sap-cloud-platform", "sap-bw", "sap-fiori"]


class StackOverflowScraper(BaseScraper):
    platform = "stack_overflow"
    rate_limit_seconds = 0.4  # API is generous up to 30/sec

    async def _list_questions(self, tag: str, *, max_pages: int) -> List[dict]:
        questions: List[dict] = []
        for page in range(1, max_pages + 1):
            try:
                r = await self.fetch(
                    f"{API}/questions",
                    params={
                        "order": "desc",
                        "sort": "votes",
                        "tagged": tag,
                        "site": "stackoverflow",
                        "pagesize": 100,
                        "page": page,
                        "filter": "withbody",
                    },
                )
            except Exception:  # noqa: BLE001
                log.warning("so_questions_failed", tag=tag, page=page, exc_info=True)
                break
            payload = r.json()
            items = payload.get("items", [])
            if not items:
                break
            questions.extend(items)
            if not payload.get("has_more"):
                break
        return questions

    async def _accepted_answer(self, q: dict) -> "dict | None":
        if not q.get("accepted_answer_id"):
            return None
        try:
            r = await self.fetch(
                f"{API}/answers/{q['accepted_answer_id']}",
                params={"site": "stackoverflow", "filter": "withbody"},
            )
        except Exception:  # noqa: BLE001
            return None
        items = r.json().get("items", [])
        return items[0] if items else None

    async def scrape(self, *, target: int) -> AsyncIterator[CorpusEntry]:
        per_tag = max(target // len(TAGS) + 50, 100)
        emitted = 0
        for tag in TAGS:
            if emitted >= target:
                break
            questions = await self._list_questions(tag, max_pages=per_tag // 100 + 2)
            log.info("so_tag_listed", tag=tag, count=len(questions))
            for q in questions:
                if emitted >= target:
                    break
                if not q.get("accepted_answer_id"):
                    continue
                ans = await self._accepted_answer(q)
                if ans is None:
                    continue
                question_text = redact_pii(q.get("body", "")[:8000])
                answer_text = redact_pii(ans.get("body", "")[:8000])
                if len(question_text) < 50 or len(answer_text) < 30:
                    continue
                title = q.get("title", "")
                entry = CorpusEntry(
                    corpus_id=make_corpus_id(self.platform, q["link"]),
                    source_platform=self.platform,
                    source_url=q["link"],
                    error_signature=title,
                    error_text=question_text,
                    proposed_solution=answer_text,
                    system_module=detect_module(title + " " + question_text),
                    upvote_score=int(q.get("score", 0)),
                    accepted_flag=True,
                )
                emitted += 1
                yield entry
        log.info("so_done", emitted=emitted)
