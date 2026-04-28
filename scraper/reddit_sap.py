"""Reddit scraper for r/SAP and r/SAPBASIS using the public .json endpoints.

No auth needed for read-only at low volume. Filters to posts with at least one
top-level reply and a positive score.
"""
from __future__ import annotations

from typing import AsyncIterator, List

import structlog

from .base import BaseScraper, CorpusEntry, detect_module, make_corpus_id, redact_pii

log = structlog.get_logger(__name__)


SUBREDDITS = ["SAP", "SAPBASIS", "ABAP"]


class RedditScraper(BaseScraper):
    platform = "reddit"
    rate_limit_seconds = 2.0
    user_agent = "ArtilegenzCorpusBot/1.0 by u/artilegenz_research"

    async def _list_posts(self, subreddit: str, *, target: int) -> List[dict]:
        posts: List[dict] = []
        after = None
        while len(posts) < target:
            params = {"limit": 100, "t": "all"}
            if after:
                params["after"] = after
            url = f"https://www.reddit.com/r/{subreddit}/top.json"
            try:
                r = await self.fetch(url, params=params)
            except Exception:  # noqa: BLE001
                log.warning("reddit_list_failed", subreddit=subreddit, exc_info=True)
                break
            payload = r.json().get("data", {})
            children = payload.get("children", [])
            if not children:
                break
            posts.extend(c["data"] for c in children)
            after = payload.get("after")
            if not after:
                break
        return posts[:target]

    async def _top_comment(self, permalink: str) -> "str | None":
        try:
            r = await self.fetch(f"https://www.reddit.com{permalink}.json", params={"limit": 5, "sort": "top"})
        except Exception:  # noqa: BLE001
            return None
        payload = r.json()
        if not isinstance(payload, list) or len(payload) < 2:
            return None
        comments = payload[1].get("data", {}).get("children", [])
        for c in comments:
            data = c.get("data", {})
            body = data.get("body", "")
            if data.get("score", 0) > 1 and len(body) > 50:
                return body
        return None

    async def scrape(self, *, target: int) -> AsyncIterator[CorpusEntry]:
        per_sub = target // len(SUBREDDITS) + 50
        emitted = 0
        for sub in SUBREDDITS:
            if emitted >= target:
                break
            posts = await self._list_posts(sub, target=per_sub)
            log.info("reddit_sub_listed", sub=sub, count=len(posts))
            for p in posts:
                if emitted >= target:
                    break
                title = p.get("title", "")
                body = p.get("selftext", "")
                permalink = p.get("permalink")
                score = int(p.get("score", 0))
                if score < 2 or not permalink or len(body) < 100:
                    continue
                # heuristic filter: posts that mention error / dump / failed
                text_lower = (title + " " + body).lower()
                if not any(k in text_lower for k in ["error", "dump", "fail", "exception", "issue", "problem"]):
                    continue
                top = await self._top_comment(permalink)
                if not top:
                    continue
                entry = CorpusEntry(
                    corpus_id=make_corpus_id(self.platform, permalink),
                    source_platform=self.platform,
                    source_url=f"https://www.reddit.com{permalink}",
                    error_signature=title,
                    error_text=redact_pii((title + "\n\n" + body)[:8000]),
                    proposed_solution=redact_pii(top[:8000]),
                    system_module=detect_module(title + " " + body),
                    upvote_score=score,
                    accepted_flag=False,
                )
                emitted += 1
                yield entry
        log.info("reddit_done", emitted=emitted)
