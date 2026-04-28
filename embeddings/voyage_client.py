"""Voyage AI client wrapper.

* Batches up to 128 inputs per call
* Retries with exponential backoff on rate-limit / transient errors
* Returns 1024-dim float lists (voyage-3)
"""
from __future__ import annotations

import asyncio
from typing import List

import structlog
import voyageai
from tenacity import retry, stop_after_attempt, wait_exponential

from config import settings

log = structlog.get_logger(__name__)


class VoyageClient:
    BATCH_SIZE = 128

    def __init__(self) -> None:
        self.client = voyageai.Client(api_key=settings.voyage_api_key)
        self.model = settings.voyage_model
        self.dim = settings.voyage_dim

    @retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=2, min=2, max=60))
    def _embed_batch(self, texts: List[str], input_type: str) -> List[List[float]]:
        # input_type: 'document' for corpus, 'query' for live errors
        result = self.client.embed(texts=texts, model=self.model, input_type=input_type)
        return result.embeddings

    def embed_documents(self, texts: List[str]) -> List[List[float]]:
        out: List[List[float]] = []
        for i in range(0, len(texts), self.BATCH_SIZE):
            batch = texts[i : i + self.BATCH_SIZE]
            log.debug("voyage_embed_batch", size=len(batch), input_type="document")
            out.extend(self._embed_batch(batch, "document"))
        return out

    def embed_query(self, text: str) -> List[float]:
        result = self._embed_batch([text], "query")
        return result[0]

    async def embed_query_async(self, text: str) -> List[float]:
        return await asyncio.to_thread(self.embed_query, text)


_singleton: VoyageClient | None = None


def get_voyage() -> VoyageClient:
    global _singleton
    if _singleton is None:
        _singleton = VoyageClient()
    return _singleton


def build_corpus_embedding_text(*, signature: str | None, error_text: str, solution: str, module: str | None) -> str:
    """Concatenate relevant fields into the text we embed for retrieval."""
    parts: List[str] = []
    if signature:
        parts.append(f"Signature: {signature}")
    if module:
        parts.append(f"Module: {module}")
    parts.append(f"Error: {error_text}")
    parts.append(f"Solution: {solution}")
    return "\n".join(parts)[:30000]


def build_query_embedding_text(*, source: str, short_text: str, long_text: str | None, transaction: str | None, program: str | None) -> str:
    parts = [f"Source: {source}"]
    if transaction:
        parts.append(f"Transaction: {transaction}")
    if program:
        parts.append(f"Program: {program}")
    parts.append(f"Error: {short_text}")
    if long_text:
        parts.append(f"Detail: {long_text[:4000]}")
    return "\n".join(parts)[:30000]
