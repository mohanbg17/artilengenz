"""Quick query against corpus_v1 namespace -- verify corpus retrieval works.

Usage:
    python query_corpus.py --text "ABAP zero divide error"
    python query_corpus.py --text "authorization failure" --top-k 10 --module ABAP
"""
from __future__ import annotations

import argparse
import os
import sys
from typing import Any, Dict, List, Optional

import psycopg2
import psycopg2.extras


VOYAGE_API_KEY     = os.getenv("VOYAGE_API_KEY", "")
VOYAGE_MODEL       = os.getenv("VOYAGE_MODEL", "voyage-3-large")
PINECONE_API_KEY   = os.getenv("PINECONE_API_KEY", "")
PINECONE_INDEX     = os.getenv("PINECONE_INDEX", "artilegenz-sap-errors")
PINECONE_NAMESPACE = os.getenv("PINECONE_CORPUS_NAMESPACE", "corpus_v1")

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


def fetch_corpus_details(corpus_ids: List[str]) -> Dict[str, Dict[str, Any]]:
    if not corpus_ids:
        return {}
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT corpus_id, source_url, error_signature, system_module,
                          upvote_score, proposed_solution
                   FROM intel.corpus
                   WHERE corpus_id = ANY(%s)""",
                (corpus_ids,),
            )
            return {r["corpus_id"]: dict(r) for r in cur.fetchall()}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--text", required=True)
    parser.add_argument("--top-k", type=int, default=5)
    parser.add_argument("--module", help="Filter by system_module (e.g. ABAP, BASIS)")
    args = parser.parse_args()

    if not VOYAGE_API_KEY or not PINECONE_API_KEY:
        sys.exit("Set VOYAGE_API_KEY and PINECONE_API_KEY env vars first")

    import voyageai
    from pinecone import Pinecone

    voyage = voyageai.Client(api_key=VOYAGE_API_KEY)
    result = voyage.embed(texts=[args.text], model=VOYAGE_MODEL, input_type="query")
    qvec = result.embeddings[0]

    pc_filter: Dict[str, Any] = {}
    if args.module:
        pc_filter["system_module"] = {"$eq": args.module}

    pc = Pinecone(api_key=PINECONE_API_KEY)
    idx = pc.Index(PINECONE_INDEX)
    resp = idx.query(
        vector=qvec, top_k=args.top_k, namespace=PINECONE_NAMESPACE,
        include_metadata=True, filter=pc_filter or None,
    )
    matches = resp.matches or []
    details = fetch_corpus_details([m.id for m in matches])

    print(f"\n=== Top {len(matches)} corpus matches in '{PINECONE_NAMESPACE}' ===\n")
    for i, m in enumerate(matches, 1):
        md = dict(m.metadata or {})
        d = details.get(m.id, {})
        print(f"[{i}] score={m.score:.4f}")
        print(f"    module:    {md.get('system_module', '?')}")
        print(f"    upvote:    {md.get('upvote_score', '?')}")
        print(f"    signature: {md.get('error_signature', '?')[:80]}")
        print(f"    url:       {md.get('source_url', '?')}")
        sol = (d.get('proposed_solution') or md.get('solution_preview') or '')[:300]
        print(f"    solution:  {sol}")
        print()


if __name__ == "__main__":
    main()
