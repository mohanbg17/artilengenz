"""Classification worker.

Polls Postgres for raw_errors that haven't been classified yet, runs each
through the Sonnet+Opus pipeline, and writes results to intel.classifications.

Run modes:
    python classify_worker.py                # continuous (NSSM service mode)
    python classify_worker.py --once         # process backlog once, then exit
    python classify_worker.py --limit 5      # process at most N records
"""
from __future__ import annotations

import argparse
import logging
import os
import signal
import sys
import time
from typing import List, Optional

import psycopg2
import structlog


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(format="%(message)s", stream=sys.stdout,
                    level=getattr(logging, LOG_LEVEL, logging.INFO))
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
)
log = structlog.get_logger("classify_worker")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
POLL_INTERVAL_SEC = int(os.getenv("CLASSIFY_POLL_INTERVAL_SEC", "60"))
BATCH_SIZE        = int(os.getenv("CLASSIFY_BATCH_SIZE", "5"))
INTER_CALL_DELAY  = float(os.getenv("CLASSIFY_INTER_CALL_DELAY", "1.0"))

PG_DSN = (
    f"host={os.getenv('PG_HOST', 'localhost')} "
    f"port={os.getenv('PG_PORT', '5432')} "
    f"user={os.getenv('PG_USER', 'artilegenz')} "
    f"password={os.getenv('PG_PASSWORD', 'artilegenz_local_dev')} "
    f"dbname={os.getenv('PG_DATABASE', 'sap_errors')}"
)


# ---------------------------------------------------------------------------
# Pending fetch
# ---------------------------------------------------------------------------
PENDING_SQL = """
SELECT hash_key, retry_count
FROM intel.v_pending_classifications
ORDER BY occurred_at ASC
LIMIT %(limit)s;
"""


def fetch_pending(limit: int) -> List[str]:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute(PENDING_SQL, {"limit": limit})
            return [row[0] for row in cur.fetchall()]


def fetch_status() -> dict:
    with psycopg2.connect(PG_DSN) as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM intel.v_classification_status;")
            row = cur.fetchone()
            if not row:
                return {"classified": 0, "pending": 0, "failed": 0, "total": 0}
            return {
                "classified": row[0],
                "pending":    row[1],
                "failed":     row[2],
                "total":      row[3],
            }


# ---------------------------------------------------------------------------
# Graceful shutdown
# ---------------------------------------------------------------------------
class GracefulExit:
    def __init__(self) -> None:
        self.stop = False
        signal.signal(signal.SIGINT, self._handle)
        signal.signal(signal.SIGTERM, self._handle)

    def _handle(self, signum, frame) -> None:
        log.info("worker_stop_requested", signal=signum)
        self.stop = True


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
def run(args: argparse.Namespace) -> None:
    if not os.getenv("ANTHROPIC_API_KEY"):
        sys.exit("ANTHROPIC_API_KEY env var not set")
    if not os.getenv("VOYAGE_API_KEY"):
        sys.exit("VOYAGE_API_KEY env var not set")
    if not os.getenv("PINECONE_API_KEY"):
        sys.exit("PINECONE_API_KEY env var not set")

    from classifier_core import classify_with_failure_logging

    log.info(
        "worker_starting",
        once=args.once,
        limit=args.limit,
        batch_size=BATCH_SIZE,
        poll_interval=POLL_INTERVAL_SEC,
    )

    status = fetch_status()
    log.info("startup_status", **status)

    exit_flag = GracefulExit()
    processed = 0
    limit_remaining: Optional[int] = args.limit

    while not exit_flag.stop:
        batch_limit = BATCH_SIZE
        if limit_remaining is not None:
            batch_limit = min(batch_limit, limit_remaining)
            if batch_limit <= 0:
                log.info("limit_reached_exiting", processed=processed)
                break

        hash_keys = fetch_pending(batch_limit)

        if not hash_keys:
            log.info("no_pending_records", sleep_sec=POLL_INTERVAL_SEC)
            if args.once:
                log.info("once_mode_complete_exiting", processed=processed)
                break
            for _ in range(POLL_INTERVAL_SEC):
                if exit_flag.stop:
                    break
                time.sleep(1)
            continue

        log.info("batch_starting", size=len(hash_keys))
        for hk in hash_keys:
            if exit_flag.stop:
                break
            log.info("classifying", hash_key=hk[:16] + "...")
            try:
                result = classify_with_failure_logging(hk)
                if result:
                    log.info(
                        "classified_ok",
                        hash_key=hk[:16] + "...",
                        confidence=round(result.final_confidence, 3),
                        badge=result.badge,
                        cache_savings_sonnet=result.sonnet_tokens.get("cache_read_input_tokens", 0),
                        cache_savings_opus=result.opus_tokens.get("cache_read_input_tokens", 0),
                    )
                processed += 1
                if limit_remaining is not None:
                    limit_remaining -= 1
            except Exception as exc:  # noqa: BLE001
                log.exception("unexpected_error", hash_key=hk[:16] + "...", err=str(exc))
            time.sleep(INTER_CALL_DELAY)

    final_status = fetch_status()
    log.info("worker_stopped", processed=processed, **final_status)


def main() -> None:
    parser = argparse.ArgumentParser(description="SAP error classification worker")
    parser.add_argument("--once", action="store_true",
                        help="Process pending records once, then exit")
    parser.add_argument("--limit", type=int, default=None,
                        help="Stop after N classifications")
    args = parser.parse_args()
    run(args)


if __name__ == "__main__":
    main()
