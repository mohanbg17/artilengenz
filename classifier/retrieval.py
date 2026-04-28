"""Adaptive retrieval for the classifier.

Given a raw error, embed it as a Voyage 'query' and probe both Pinecone
namespaces (raw_errors_v1 and corpus_v1). Choose how many to keep from
each based on which scored higher on the probe -- but always retain a
minimum from each namespace so the LLM gets balanced context.

Returns hydrated records (with full text from Postgres), not just vectors.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from statistics import mean
from typing import Any, Dict, List, Optional

import psycopg2
import psycopg2.extras


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
VOYAGE_API_KEY     = os.getenv("VOYAGE_API_KEY", "")
VOYAGE_MODEL       = os.getenv("VOYAGE_MODEL", "voyage-3-large")
PINECONE_API_KEY   = os.getenv("PINECONE_API_KEY", "")
PINECONE_INDEX     = os.getenv("PINECONE_INDEX", "artilegenz-sap-errors")
NS_RAW             = os.getenv("PINECONE_NAMESPACE", "raw_errors_v1")
NS_CORPUS          = os.getenv("PINECONE_CORPUS_NAMESPACE", "corpus_v1")

# Total context budget (across both namespaces)
TOTAL_K_DEFAULT = 10
# Hard floor per namespace (don't go below this even if adaptive logic wants to)
MIN_PER_NAMESPACE = 2

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------
@dataclass
class RetrievedItem:
    namespace: str
    pinecone_id: str
    score: float
    metadata: Dict[str, Any]
    full_text: str = ""           # hydrated from Postgres
    source_url: Optional[str] = None
    short_text: Optional[str] = None


@dataclass
class RetrievalResult:
    raw_matches: List[RetrievedItem] = field(default_factory=list)
    corpus_matches: List[RetrievedItem] = field(default_factory=list)
    # Diagnostics
    raw_avg_score: float = 0.0
    corpus_avg_score: float = 0.0
    pulled_from_raw: int = 0
    pulled_from_corpus: int = 0
    strategy: str = ""             # 'corpus_heavy' | 'raw_heavy' | 'balanced'

    def adaptive_split(self) -> Dict[str, Any]:
        """Diagnostics blob written to intel.classifications.adaptive_split"""
        return {
            "raw_avg_score": round(self.raw_avg_score, 4),
            "corpus_avg_score": round(self.corpus_avg_score, 4),
            "pulled_from_raw": self.pulled_from_raw,
            "pulled_from_corpus": self.pulled_from_corpus,
            "strategy": self.strategy,
        }


# ---------------------------------------------------------------------------
# Voyage query embedding
# ---------------------------------------------------------------------------
def embed_query(text: str) -> List[float]:
    import voyageai
    client = voyageai.Client(api_key=VOYAGE_API_KEY)
    result = client.embed(texts=[text], model=VOYAGE_MODEL, input_type="query")
    return result.embeddings[0]


# ---------------------------------------------------------------------------
# Pinecone probe
# ---------------------------------------------------------------------------
def _probe(index, namespace: str, vector: List[float], k: int) -> List[RetrievedItem]:
    resp = index.query(
        vector=vector, top_k=k, namespace=namespace, include_metadata=True,
    )
    out: List[RetrievedItem] = []
    for m in resp.matches or []:
        out.append(RetrievedItem(
            namespace=namespace,
            pinecone_id=m.id,
            score=float(m.score),
            metadata=dict(m.metadata or {}),
        ))
    return out


# ---------------------------------------------------------------------------
# Adaptive selection
# ---------------------------------------------------------------------------
def _adaptive_split(
    raw_avg: float, corpus_avg: float, total_k: int, min_each: int
) -> tuple[int, int, str]:
    """Decide how many slots each namespace gets."""
    diff = corpus_avg - raw_avg
    if diff > 0.05:
        # corpus much better -> 70/30 split
        n_corpus = int(round(total_k * 0.7))
        n_raw = total_k - n_corpus
        strategy = "corpus_heavy"
    elif diff < -0.05:
        n_raw = int(round(total_k * 0.7))
        n_corpus = total_k - n_raw
        strategy = "raw_heavy"
    else:
        n_corpus = total_k // 2
        n_raw = total_k - n_corpus
        strategy = "balanced"

    # Enforce floor
    n_raw = max(n_raw, min_each)
    n_corpus = max(n_corpus, min_each)
    return n_raw, n_corpus, strategy


# ---------------------------------------------------------------------------
# Postgres hydration
# ---------------------------------------------------------------------------
def _hydrate_raw(items: List[RetrievedItem]) -> None:
    """Fill in full short_text + long_text from raw.raw_errors."""
    if not items:
        return
    ids = [i.pinecone_id for i in items]
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT hash_key, short_text, long_text, error_id, transaction, program,
                          severity, system_id, occurred_at
                   FROM raw.raw_errors WHERE hash_key = ANY(%s)""",
                (ids,),
            )
            by_hash = {r["hash_key"]: dict(r) for r in cur.fetchall()}
    for item in items:
        rec = by_hash.get(item.pinecone_id, {})
        item.short_text = rec.get("short_text")
        # Build a compact text representation for the LLM
        parts = []
        if rec.get("error_id"):
            parts.append(f"Error ID: {rec['error_id']}")
        if rec.get("severity"):
            parts.append(f"Severity: {rec['severity']}")
        if rec.get("system_id"):
            parts.append(f"System: {rec['system_id']}")
        if rec.get("occurred_at"):
            parts.append(f"Occurred: {rec['occurred_at']}")
        if rec.get("transaction"):
            parts.append(f"Transaction: {rec['transaction']}")
        if rec.get("program"):
            parts.append(f"Program: {rec['program']}")
        if rec.get("short_text"):
            parts.append(f"\nShort: {rec['short_text']}")
        if rec.get("long_text"):
            # Truncate long_text to keep prompt manageable
            parts.append(f"\nLong:\n{rec['long_text'][:3000]}")
        item.full_text = "\n".join(parts)


def _hydrate_corpus(items: List[RetrievedItem]) -> None:
    """Fill in full error_text + proposed_solution from intel.corpus."""
    if not items:
        return
    ids = [i.pinecone_id for i in items]
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT corpus_id, source_url, error_signature, system_module,
                          error_text, proposed_solution, upvote_score, tags
                   FROM intel.corpus WHERE corpus_id = ANY(%s)""",
                (ids,),
            )
            by_id = {r["corpus_id"]: dict(r) for r in cur.fetchall()}
    for item in items:
        rec = by_id.get(item.pinecone_id, {})
        item.source_url = rec.get("source_url")
        parts = []
        if rec.get("error_signature"):
            parts.append(f"Signature: {rec['error_signature']}")
        if rec.get("system_module"):
            parts.append(f"Module: {rec['system_module']}")
        if rec.get("upvote_score"):
            parts.append(f"Upvotes: {rec['upvote_score']}")
        if rec.get("source_url"):
            parts.append(f"URL: {rec['source_url']}")
        if rec.get("error_text"):
            parts.append(f"\nQuestion / Error context:\n{(rec['error_text'] or '')[:4000]}")
        if rec.get("proposed_solution"):
            parts.append(f"\nAccepted Solution:\n{(rec['proposed_solution'] or '')[:4000]}")
        item.full_text = "\n".join(parts)


# ---------------------------------------------------------------------------
# Main entrypoint
# ---------------------------------------------------------------------------
def adaptive_retrieve(
    query_text: str,
    total_k: int = TOTAL_K_DEFAULT,
    min_each: int = MIN_PER_NAMESPACE,
    probe_size: int = 5,
) -> RetrievalResult:
    """
    1. Embed query.
    2. Probe both namespaces (top probe_size from each).
    3. Compute avg score for each namespace.
    4. Decide split (n_raw + n_corpus = total_k, each >= min_each).
    5. Re-query if asked for more than probe_size; otherwise slice.
    6. Hydrate full text from Postgres.
    """
    from pinecone import Pinecone

    qvec = embed_query(query_text)

    pc = Pinecone(api_key=PINECONE_API_KEY)
    index = pc.Index(PINECONE_INDEX)

    # Probe both namespaces
    raw_probe = _probe(index, NS_RAW, qvec, probe_size)
    corpus_probe = _probe(index, NS_CORPUS, qvec, probe_size)

    raw_avg = mean([m.score for m in raw_probe]) if raw_probe else 0.0
    corpus_avg = mean([m.score for m in corpus_probe]) if corpus_probe else 0.0

    n_raw, n_corpus, strategy = _adaptive_split(raw_avg, corpus_avg, total_k, min_each)

    # If we need more than probe_size from a namespace, re-query
    if n_raw > probe_size:
        raw_probe = _probe(index, NS_RAW, qvec, n_raw)
    if n_corpus > probe_size:
        corpus_probe = _probe(index, NS_CORPUS, qvec, n_corpus)

    raw_matches = raw_probe[:n_raw]
    corpus_matches = corpus_probe[:n_corpus]

    # Hydrate full text from Postgres
    _hydrate_raw(raw_matches)
    _hydrate_corpus(corpus_matches)

    return RetrievalResult(
        raw_matches=raw_matches,
        corpus_matches=corpus_matches,
        raw_avg_score=raw_avg,
        corpus_avg_score=corpus_avg,
        pulled_from_raw=len(raw_matches),
        pulled_from_corpus=len(corpus_matches),
        strategy=strategy,
    )
