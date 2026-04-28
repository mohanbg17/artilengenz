"""Quick CLI to query similar SAP errors via Pinecone.

After raw_worker.py has embedded records, use this to verify retrieval works.

Usage:
    # Free-text search
    python query_similar.py --text "ABAP zerodivide CL_SADL"

    # Find errors similar to an existing one (paste hash_key)
    python query_similar.py --hash-key abc123def456...

    # Filter by source/severity/system
    python query_similar.py --text "auth failure" --source ST22 --top-k 10
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from typing import Any, Dict, List, Optional

import psycopg2
import psycopg2.extras


# ---------------------------------------------------------------------------
# Lazy imports
# ---------------------------------------------------------------------------

def _import_runtime():
    import voyageai
    from pinecone import Pinecone
    return voyageai, Pinecone


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

VOYAGE_API_KEY = os.getenv("VOYAGE_API_KEY", "")
VOYAGE_MODEL = os.getenv("VOYAGE_MODEL", "voyage-3-large")
PINECONE_API_KEY = os.getenv("PINECONE_API_KEY", "")
PINECONE_INDEX = os.getenv("PINECONE_INDEX", "artilegenz-sap-errors")
PINECONE_NAMESPACE = os.getenv("PINECONE_NAMESPACE", "raw_errors_v1")

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fetch_record_by_hash(hash_key: str) -> Optional[Dict[str, Any]]:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """
                SELECT hash_key, source, system_id, occurred_at, severity,
                       short_text, long_text, error_id, transaction, program
                FROM raw.raw_errors
                WHERE hash_key = %s
                """,
                (hash_key,),
            )
            row = cur.fetchone()
            return dict(row) if row else None


def build_query_text_from_record(rec: Dict[str, Any]) -> str:
    parts = [f"Source: {rec.get('source', 'UNKNOWN')}"]
    if rec.get("severity"):
        parts.append(f"Severity: {rec['severity']}")
    if rec.get("transaction"):
        parts.append(f"Transaction: {rec['transaction']}")
    if rec.get("program"):
        parts.append(f"Program: {rec['program']}")
    if rec.get("short_text"):
        parts.append(f"Short text: {rec['short_text']}")
    if rec.get("long_text"):
        parts.append(f"Long text:\n{rec['long_text'][:8000]}")
    return "\n".join(parts)[:30000]


def fetch_postgres_details(hash_keys: List[str]) -> Dict[str, Dict[str, Any]]:
    if not hash_keys:
        return {}
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """
                SELECT hash_key, source, system_id, occurred_at, severity,
                       short_text, error_id, transaction, program
                FROM raw.raw_errors
                WHERE hash_key = ANY(%s)
                """,
                (hash_keys,),
            )
            return {row["hash_key"]: dict(row) for row in cur.fetchall()}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run() -> None:
    parser = argparse.ArgumentParser(description="Query similar SAP errors")
    parser.add_argument("--text", help="Free-text search query")
    parser.add_argument("--hash-key", help="Find errors similar to this one")
    parser.add_argument("--top-k", type=int, default=5, help="Number of results")
    parser.add_argument("--source", help="Filter by source (ST22, SLG1, SM21)")
    parser.add_argument("--severity", help="Filter by severity")
    parser.add_argument("--system-id", help="Filter by system ID")
    parser.add_argument("--json", action="store_true", help="JSON output")
    args = parser.parse_args()

    if not args.text and not args.hash_key:
        parser.error("Provide --text or --hash-key")

    if not VOYAGE_API_KEY or not PINECONE_API_KEY:
        sys.exit("Set VOYAGE_API_KEY and PINECONE_API_KEY env vars first")

    voyageai, Pinecone = _import_runtime()

    # ---- Build query text
    if args.hash_key:
        rec = fetch_record_by_hash(args.hash_key)
        if not rec:
            sys.exit(f"No record found for hash_key={args.hash_key}")
        query_text = build_query_text_from_record(rec)
        print(f"\n[Query record]")
        print(f"  hash_key: {rec['hash_key'][:16]}...")
        print(f"  source:   {rec['source']}")
        print(f"  short:    {(rec.get('short_text') or '')[:120]}\n")
    else:
        query_text = args.text

    # ---- Embed the query (input_type=query, NOT document)
    voyage_client = voyageai.Client(api_key=VOYAGE_API_KEY)
    result = voyage_client.embed(
        texts=[query_text],
        model=VOYAGE_MODEL,
        input_type="query",
    )
    query_vector = result.embeddings[0]

    # ---- Build Pinecone filter
    pc_filter: Dict[str, Any] = {}
    if args.source:
        pc_filter["source"] = {"$eq": args.source}
    if args.severity:
        pc_filter["severity"] = {"$eq": args.severity}
    if args.system_id:
        pc_filter["system_id"] = {"$eq": args.system_id}

    # ---- Query Pinecone
    pc = Pinecone(api_key=PINECONE_API_KEY)
    index = pc.Index(PINECONE_INDEX)
    resp = index.query(
        vector=query_vector,
        top_k=args.top_k,
        namespace=PINECONE_NAMESPACE,
        include_metadata=True,
        filter=pc_filter or None,
    )

    matches = resp.matches or []
    if args.hash_key:
        # Don't show the query record as its own top match
        matches = [m for m in matches if m.id != args.hash_key]

    # ---- Hydrate from Postgres for full short_text
    hash_keys = [m.id for m in matches]
    pg_details = fetch_postgres_details(hash_keys)

    # ---- Output
    if args.json:
        out = []
        for m in matches:
            md = dict(m.metadata or {})
            pg = pg_details.get(m.id, {})
            out.append({
                "score": float(m.score),
                "hash_key": m.id,
                "metadata": md,
                "postgres": {
                    k: (v.isoformat() if isinstance(v, datetime) else v)
                    for k, v in pg.items()
                },
            })
        print(json.dumps(out, indent=2, default=str))
    else:
        print(f"\n=== Top {len(matches)} similar errors ===\n")
        for i, m in enumerate(matches, 1):
            md = dict(m.metadata or {})
            pg = pg_details.get(m.id, {})
            print(f"[{i}] score={m.score:.4f}")
            print(f"    hash_key:   {m.id[:16]}...")
            print(f"    source:     {md.get('source', '?')}")
            print(f"    severity:   {md.get('severity', '?')}")
            print(f"    system:     {md.get('system_id', '?')}")
            print(f"    occurred:   {md.get('occurred_at', '?')}")
            short = pg.get("short_text") or md.get("short_text_preview") or ""
            print(f"    short:      {short[:140]}")
            if md.get("transaction"):
                print(f"    txn:        {md['transaction']}")
            if md.get("program"):
                print(f"    program:    {md['program']}")
            print()


if __name__ == "__main__":
    run()
