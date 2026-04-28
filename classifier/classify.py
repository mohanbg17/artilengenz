"""CLI: classify a single SAP error.

Usage:
    python classify.py --hash-key <hash>
    python classify.py --hash-key <hash> --json    # full JSON output
    python classify.py --latest                    # classify the latest unclassified error
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Optional

import psycopg2
import psycopg2.extras


PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


def find_latest_unclassified() -> Optional[str]:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT r.hash_key
                FROM raw.raw_errors r
                LEFT JOIN intel.v_latest_classifications c ON c.error_hash_key = r.hash_key
                WHERE c.classification_id IS NULL
                ORDER BY r.occurred_at DESC
                LIMIT 1
            """)
            row = cur.fetchone()
            return row[0] if row else None


def find_hash_by_prefix(prefix: str) -> Optional[str]:
    """Look up full hash from a prefix (convenience for short ids)."""
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT hash_key FROM raw.raw_errors WHERE hash_key LIKE %s LIMIT 2",
                (prefix + "%",),
            )
            rows = cur.fetchall()
            if len(rows) == 1:
                return rows[0][0]
            elif len(rows) > 1:
                sys.exit(f"Prefix {prefix!r} matched multiple hash_keys; use a longer prefix.")
            return None


def main() -> None:
    parser = argparse.ArgumentParser(description="Classify a single SAP error")
    parser.add_argument("--hash-key", help="Full hash_key (or unique prefix)")
    parser.add_argument("--latest", action="store_true",
                        help="Classify the most recent unclassified error")
    parser.add_argument("--json", action="store_true",
                        help="Emit full diagnosis as JSON (instead of pretty markdown)")
    args = parser.parse_args()

    if not args.hash_key and not args.latest:
        parser.error("Provide --hash-key or --latest")

    if args.latest:
        hash_key = find_latest_unclassified()
        if not hash_key:
            sys.exit("No unclassified errors found.")
        print(f"# Latest unclassified: {hash_key[:16]}...\n")
    else:
        hash_key = args.hash_key
        # If short, expand
        if len(hash_key) < 64:
            full = find_hash_by_prefix(hash_key)
            if not full:
                sys.exit(f"No raw error found matching prefix {hash_key!r}")
            hash_key = full

    from classifier_core import classify

    result = classify(hash_key)

    if args.json:
        print(json.dumps({
            "classification_id": result.classification_id,
            "final": {
                "root_cause": result.final_root_cause,
                "severity": result.final_severity,
                "remediation_steps": result.final_remediation,
                "confidence": result.final_confidence,
                "badge": result.badge,
                "summary_md": result.final_summary_md,
            },
            "sonnet_diagnosis": result.sonnet_diagnosis,
            "opus_critique": result.opus_critique,
            "adaptive_split": result.adaptive_split,
            "tokens": {
                "sonnet": result.sonnet_tokens,
                "opus": result.opus_tokens,
            },
        }, indent=2, ensure_ascii=False))
    else:
        # Pretty terminal output
        print("=" * 80)
        print(f"DIAGNOSIS  ({result.badge}, confidence={result.final_confidence:.2f})")
        print("=" * 80)
        print()
        print(f"Root cause: {result.final_root_cause}")
        print(f"Severity:   {result.final_severity}")
        print(f"Verdict:    {result.opus_critique.get('verdict', '?')}")
        print()
        print("Remediation steps:")
        for i, step in enumerate(result.final_remediation, 1):
            print(f"  {i}. {step}")
        print()
        print("-" * 80)
        print("SUMMARY")
        print("-" * 80)
        print(result.final_summary_md)
        print()
        print("-" * 80)
        print(f"Tokens used (Sonnet): {result.sonnet_tokens}")
        print(f"Tokens used (Opus):   {result.opus_tokens}")
        print(f"Retrieval split:      {result.adaptive_split}")
        print(f"Classification ID:    {result.classification_id}")


if __name__ == "__main__":
    main()
