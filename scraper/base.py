"""Shared schema and base class for corpus scrapers."""
from __future__ import annotations

import asyncio
import hashlib
import re
from datetime import datetime
from typing import AsyncIterator, Optional

import httpx
import structlog
from pydantic import BaseModel, Field
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

log = structlog.get_logger(__name__)


class CorpusEntry(BaseModel):
    """One scraped error/resolution pair."""

    corpus_id: str
    source_platform: str
    source_url: str
    error_signature: Optional[str] = None
    error_text: str
    proposed_solution: str
    system_module: Optional[str] = None
    upvote_score: int = 0
    accepted_flag: bool = False
    scraped_at: datetime = Field(default_factory=datetime.utcnow)


def make_corpus_id(platform: str, url: str) -> str:
    h = hashlib.sha256(f"{platform}|{url}".encode()).hexdigest()
    return f"{platform[:6]}-{h[:32]}"


_PII_PATTERNS = [
    (re.compile(r"\b[\w._%+-]+@[\w.-]+\.[A-Za-z]{2,}\b"), "<email>"),
    (re.compile(r"\b\d{3}-\d{2}-\d{4}\b"), "<ssn>"),
    (re.compile(r"\b(?:\d[ -]*?){13,16}\b"), "<card>"),
    (re.compile(r"\b\+?\d{1,3}[ -]?\(?\d{2,4}\)?[ -]?\d{3,4}[ -]?\d{3,4}\b"), "<phone>"),
]


def redact_pii(text: str) -> str:
    if not text:
        return text
    out = text
    for pat, repl in _PII_PATTERNS:
        out = pat.sub(repl, out)
    return out


def detect_module(text: str) -> Optional[str]:
    """Cheap heuristic to tag SAP module from keywords. Refined later by LLM."""
    if not text:
        return None
    t = text.lower()
    if any(k in t for k in ["abap", "syntax error", "dump", "short dump"]):
        return "BASIS"
    if any(k in t for k in ["sales order", "vbak", "vbap", "billing", "delivery"]):
        return "SD"
    if any(k in t for k in ["purchase order", "ekko", "ekpo", "goods receipt"]):
        return "MM"
    if any(k in t for k in ["invoice", "fb01", "bseg", "general ledger"]):
        return "FI"
    if any(k in t for k in ["controlling", "ksb1", "cost center", "internal order"]):
        return "CO"
    if any(k in t for k in ["calculation view", "hana", "calc view", "scenario"]):
        return "HANA"
    if any(k in t for k in ["bods", "data services", "dataflow"]):
        return "BODS"
    if any(k in t for k in ["btp", "cap", "ui5", "fiori"]):
        return "BTP"
    return None


class BaseScraper:
    """Shared HTTP client + politeness + retry."""

    platform: str = "base"
    user_agent: str = "ArtilegenzCorpusBot/1.0 (research; contact: hello@artilegenz.com)"
    rate_limit_seconds: float = 1.0

    def __init__(self) -> None:
        self.client = httpx.AsyncClient(
            timeout=30.0,
            follow_redirects=True,
            headers={"User-Agent": self.user_agent},
        )
        self._last_call: float = 0.0

    async def close(self) -> None:
        await self.client.aclose()

    async def _polite_wait(self) -> None:
        loop = asyncio.get_event_loop()
        now = loop.time()
        elapsed = now - self._last_call
        if elapsed < self.rate_limit_seconds:
            await asyncio.sleep(self.rate_limit_seconds - elapsed)
        self._last_call = loop.time()

    @retry(
        stop=stop_after_attempt(4),
        wait=wait_exponential(multiplier=2, min=2, max=60),
        retry=retry_if_exception_type((httpx.HTTPError,)),
        reraise=True,
    )
    async def fetch(self, url: str, **kwargs) -> httpx.Response:
        await self._polite_wait()
        log.debug("scraper_fetch", platform=self.platform, url=url)
        r = await self.client.get(url, **kwargs)
        if r.status_code == 429:
            log.warning("scraper_rate_limited", platform=self.platform, url=url)
            await asyncio.sleep(30)
            raise httpx.HTTPError("rate limited")
        r.raise_for_status()
        return r

    async def scrape(self, *, target: int) -> AsyncIterator[CorpusEntry]:
        """Subclasses override. Yield CorpusEntry instances."""
        raise NotImplementedError
        yield  # type: ignore[unreachable]
