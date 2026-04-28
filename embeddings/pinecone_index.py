"""Pinecone serverless index wrapper.

* Lazy index creation if missing
* Hybrid score combines cosine similarity + accepted_flag bonus + log(upvote)
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import structlog
from pinecone import Pinecone, ServerlessSpec

from config import settings

log = structlog.get_logger(__name__)


@dataclass
class Match:
    pinecone_id: str
    corpus_id: str
    score: float            # cosine similarity from Pinecone
    hybrid_score: float     # boosted score (accepted/upvote)
    metadata: Dict[str, Any]


class PineconeIndex:
    def __init__(self) -> None:
        self.pc = Pinecone(api_key=settings.pinecone_api_key)
        self.index_name = settings.pinecone_index
        self.namespace = settings.pinecone_namespace
        self._ensure_index()
        self.index = self.pc.Index(self.index_name)

    def _ensure_index(self) -> None:
        existing = [i["name"] for i in self.pc.list_indexes()]
        if self.index_name not in existing:
            log.info("pinecone_create_index", name=self.index_name, dim=settings.voyage_dim)
            self.pc.create_index(
                name=self.index_name,
                dimension=settings.voyage_dim,
                metric="cosine",
                spec=ServerlessSpec(cloud=settings.pinecone_cloud, region=settings.pinecone_region),
            )

    def upsert_batch(self, *, ids: List[str], vectors: List[List[float]], metadatas: List[Dict[str, Any]]) -> None:
        items = [
            {"id": i, "values": v, "metadata": m}
            for i, v, m in zip(ids, vectors, metadatas)
        ]
        log.debug("pinecone_upsert", batch=len(items))
        self.index.upsert(vectors=items, namespace=self.namespace)

    def query(self, *, vector: List[float], top_k: int, filter: Optional[Dict[str, Any]] = None) -> List[Match]:
        resp = self.index.query(
            vector=vector,
            top_k=top_k,
            namespace=self.namespace,
            include_metadata=True,
            filter=filter,
        )
        out: List[Match] = []
        for m in resp.matches:
            md = dict(m.metadata or {})
            score = float(m.score)
            hybrid = self._hybrid(score, md)
            out.append(
                Match(
                    pinecone_id=m.id,
                    corpus_id=md.get("corpus_id", m.id),
                    score=score,
                    hybrid_score=hybrid,
                    metadata=md,
                )
            )
        # Re-sort by hybrid
        out.sort(key=lambda x: x.hybrid_score, reverse=True)
        return out

    @staticmethod
    def _hybrid(cosine: float, md: Dict[str, Any]) -> float:
        accepted_bonus = 0.05 if md.get("accepted_flag") else 0.0
        upvote = float(md.get("upvote_score") or 0.0)
        upvote_bonus = 0.02 * math.log(1.0 + max(upvote, 0.0))
        return min(cosine + accepted_bonus + upvote_bonus, 1.0)


_singleton: PineconeIndex | None = None


def get_index() -> PineconeIndex:
    global _singleton
    if _singleton is None:
        _singleton = PineconeIndex()
    return _singleton
