"""Raw error embedding worker.

Continuously polls Postgres for raw_errors that haven't been embedded
yet, batches them through Voyage AI, and upserts vectors to Pinecone.

Designed to run as a Windows service (via NSSM) or standalone.

Run:
    python raw_worker.py
    python raw_worker.py --once     # process pending then exit
    python raw_worker.py --dry-run  # log what would happen, no API calls

Env vars (set via NSSM service or .env):
    VOYAGE_API_KEY      -- required
    PINECONE_API_KEY    -- required
    PINECONE_INDEX      -- default: artilegenz-sap-errors
    PINECONE_NAMESPACE  -- default: raw_errors_v1
    PINECONE_CLOUD      -- default: aws
    PINECONE_REGION     -- default: us-east-1
    VOYAGE_MODEL        -- default: voyage-3-large
    VOYAGE_DIM          -- default: 1024
    PG_HOST/PORT/USER/PASSWORD/DATABASE
    POLL_INTERVAL_SEC   -- default: 30
    BATCH_SIZE          -- default: 32
    MAX_TEXT_TOKENS     -- default: 30000  (voyage-3-large supports 32k)
"""
from __future__ import annotations

import argparse
import logging
import os
import signal
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import structlog

# ---------------------------------------------------------------------------
# Structured logging setup
# ---------------------------------------------------------------------------
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    format="%(message)s",
    stream=sys.stdout,
    level=getattr(logging, LOG_LEVEL, logging.INFO),
)
structlog.configure(
    processors=[
        structlog.processors.add_log_level,
        structlog.processors.TimeStamper(fmt="iso"),
        structlog.processors.JSONRenderer(),
    ],
    wrapper_class=structlog.make_filtering_bound_logger(
        getattr(logging, LOG_LEVEL, logging.INFO)
    ),
)
log = structlog.get_logger("raw_worker")


# ---------------------------------------------------------------------------
# Imports that may need API keys -- delayed to allow --help without them
# ---------------------------------------------------------------------------

def _import_runtime():
    """Lazy import so --help works without API keys set."""
    from tenacity import retry, stop_after_attempt, wait_exponential
    import voyageai
    from pinecone import Pinecone, ServerlessSpec

    return voyageai, Pinecone, ServerlessSpec, retry, stop_after_attempt, wait_exponential


# ---------------------------------------------------------------------------
# Config (read from env)
# ---------------------------------------------------------------------------

class WorkerConfig:
    voyage_api_key: str
    voyage_model: str
    voyage_dim: int
    pinecone_api_key: str
    pinecone_index: str
    pinecone_namespace: str
    pinecone_cloud: str
    pinecone_region: str
    poll_interval_sec: int
    batch_size: int
    max_text_chars: int  # rough proxy for tokens (4 chars ~= 1 token)
    max_retries: int

    def __init__(self) -> None:
        self.voyage_api_key = os.getenv("VOYAGE_API_KEY", "")
        self.voyage_model = os.getenv("VOYAGE_MODEL", "voyage-3-large")
        self.voyage_dim = int(os.getenv("VOYAGE_DIM", "1024"))
        self.pinecone_api_key = os.getenv("PINECONE_API_KEY", "")
        self.pinecone_index = os.getenv("PINECONE_INDEX", "artilegenz-sap-errors")
        self.pinecone_namespace = os.getenv("PINECONE_NAMESPACE", "raw_errors_v1")
        self.pinecone_cloud = os.getenv("PINECONE_CLOUD", "aws")
        self.pinecone_region = os.getenv("PINECONE_REGION", "us-east-1")
        self.poll_interval_sec = int(os.getenv("POLL_INTERVAL_SEC", "30"))
        self.batch_size = int(os.getenv("BATCH_SIZE", "32"))
        # voyage-3-large = 32k token context. Use char-budget of 30k tokens * 4 = 120k chars
        # but cap practically at 100k chars to leave headroom for tokenizer expansion.
        self.max_text_chars = int(os.getenv("MAX_TEXT_CHARS", "100000"))
        self.max_retries = int(os.getenv("MAX_RETRIES", "3"))

    def validate(self) -> None:
        missing = []
        if not self.voyage_api_key:
            missing.append("VOYAGE_API_KEY")
        if not self.pinecone_api_key:
            missing.append("PINECONE_API_KEY")
        if missing:
            log.error("config_missing_env_vars", missing=missing)
            raise SystemExit(
                f"Missing required env vars: {', '.join(missing)}. "
                "Set them via NSSM service config or .env file."
            )


# ---------------------------------------------------------------------------
# Text builder -- combines short_text + long_text + metadata into single doc
# ---------------------------------------------------------------------------

def build_embed_text(record: Dict[str, Any], max_chars: int) -> tuple[str, bool]:
    """Compose the text we send to Voyage. Returns (text, was_truncated)."""
    parts: List[str] = []

    parts.append(f"Source: {record.get('source', 'UNKNOWN')}")
    parts.append(f"System: {record.get('system_id', '')}")

    if record.get("severity"):
        parts.append(f"Severity: {record['severity']}")
    if record.get("transaction"):
        parts.append(f"Transaction: {record['transaction']}")
    if record.get("program"):
        parts.append(f"Program: {record['program']}")
    if record.get("error_id"):
        parts.append(f"Error ID: {record['error_id']}")

    short_text = (record.get("short_text") or "").strip()
    if short_text:
        parts.append(f"Short text: {short_text}")

    long_text = (record.get("long_text") or "").strip()
    if long_text:
        parts.append(f"Long text:\n{long_text}")

    full = "\n".join(parts)

    if len(full) > max_chars:
        # Truncate at boundary -- keep header + start of long_text
        truncated = full[:max_chars] + "\n[...truncated]"
        return truncated, True
    return full, False


# ---------------------------------------------------------------------------
# Pinecone metadata builder
# ---------------------------------------------------------------------------

def build_metadata(record: Dict[str, Any]) -> Dict[str, Any]:
    """Pinecone metadata is for filtering and quick UI preview.
    Keep it small (<40KB per vector). Full text stays in Postgres."""
    occurred_at = record.get("occurred_at")
    if isinstance(occurred_at, datetime):
        occurred_iso = occurred_at.isoformat()
        occurred_unix = int(occurred_at.timestamp())
    else:
        occurred_iso = str(occurred_at) if occurred_at else ""
        occurred_unix = 0

    md: Dict[str, Any] = {
        "hash_key": record["hash_key"],
        "source": record.get("source") or "",
        "system_id": record.get("system_id") or "",
        "occurred_at": occurred_iso,
        "occurred_unix": occurred_unix,
        "severity": record.get("severity") or "",
    }
    if record.get("error_id"):
        md["error_id"] = str(record["error_id"])[:200]
    if record.get("transaction"):
        md["transaction"] = str(record["transaction"])[:64]
    if record.get("program"):
        md["program"] = str(record["program"])[:128]
    if record.get("user_name"):
        md["user_name"] = str(record["user_name"])[:64]
    if record.get("instance"):
        md["instance"] = str(record["instance"])[:64]
    if record.get("short_text"):
        md["short_text_preview"] = str(record["short_text"])[:500]

    return md


# ---------------------------------------------------------------------------
# Voyage embedder
# ---------------------------------------------------------------------------

class VoyageEmbedder:
    def __init__(self, cfg: WorkerConfig, voyageai_module, retry_decorator) -> None:
        self.client = voyageai_module.Client(api_key=cfg.voyage_api_key)
        self.model = cfg.voyage_model
        # Build the retry-wrapped embed call
        self._embed_with_retry = retry_decorator(self._embed_raw)

    def _embed_raw(self, texts: List[str]) -> List[List[float]]:
        result = self.client.embed(
            texts=texts,
            model=self.model,
            input_type="document",  # documents being indexed (not queries)
        )
        return result.embeddings

    def embed_batch(self, texts: List[str]) -> List[List[float]]:
        return self._embed_with_retry(texts)


# ---------------------------------------------------------------------------
# Pinecone wrapper
# ---------------------------------------------------------------------------

class PineconeWriter:
    def __init__(self, cfg: WorkerConfig, Pinecone, ServerlessSpec) -> None:
        self.pc = Pinecone(api_key=cfg.pinecone_api_key)
        self.index_name = cfg.pinecone_index
        self.namespace = cfg.pinecone_namespace
        self._ensure_index(cfg, ServerlessSpec)
        self.index = self.pc.Index(self.index_name)

    def _ensure_index(self, cfg: WorkerConfig, ServerlessSpec) -> None:
        existing = [i["name"] for i in self.pc.list_indexes()]
        if cfg.pinecone_index not in existing:
            log.info(
                "pinecone_create_index",
                name=cfg.pinecone_index,
                dim=cfg.voyage_dim,
                cloud=cfg.pinecone_cloud,
                region=cfg.pinecone_region,
            )
            self.pc.create_index(
                name=cfg.pinecone_index,
                dimension=cfg.voyage_dim,
                metric="cosine",
                spec=ServerlessSpec(
                    cloud=cfg.pinecone_cloud,
                    region=cfg.pinecone_region,
                ),
            )
            # Wait for index to be ready (serverless typically <30s)
            for _ in range(60):
                desc = self.pc.describe_index(cfg.pinecone_index)
                if desc.status.get("ready"):
                    break
                time.sleep(1)

    def upsert(
        self,
        ids: List[str],
        vectors: List[List[float]],
        metadatas: List[Dict[str, Any]],
    ) -> None:
        items = [
            {"id": i, "values": v, "metadata": m}
            for i, v, m in zip(ids, vectors, metadatas)
        ]
        self.index.upsert(vectors=items, namespace=self.namespace)


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

class GracefulExit:
    """Trap SIGINT/SIGTERM so the worker finishes the current batch
    before exiting (NSSM stop signal handling)."""

    def __init__(self) -> None:
        self.stop = False
        signal.signal(signal.SIGINT, self._handle)
        signal.signal(signal.SIGTERM, self._handle)

    def _handle(self, signum, frame) -> None:
        log.info("worker_stop_requested", signal=signum)
        self.stop = True


def process_batch(
    cfg: WorkerConfig,
    embedder: VoyageEmbedder,
    pinecone_writer: PineconeWriter,
    records: List[Dict[str, Any]],
    dsn: str,
    dry_run: bool,
) -> tuple[int, int]:
    """Embed + upsert one batch. Returns (success_count, failure_count)."""
    import raw_embeddings_db as db  # local import for service mode

    texts: List[str] = []
    ids: List[str] = []
    metadatas: List[Dict[str, Any]] = []
    truncated_flags: List[bool] = []
    text_lengths: List[int] = []

    for rec in records:
        text, truncated = build_embed_text(rec, cfg.max_text_chars)
        texts.append(text)
        ids.append(rec["hash_key"])
        metadatas.append(build_metadata(rec))
        truncated_flags.append(truncated)
        text_lengths.append(len(text))

    log.info("batch_prepared", size=len(records), sample_text_len=text_lengths[0] if text_lengths else 0)

    if dry_run:
        log.info("dry_run_skip_api_calls")
        return 0, 0

    # ---- Voyage call (whole batch, retries handled inside VoyageEmbedder)
    try:
        vectors = embedder.embed_batch(texts)
    except Exception as exc:
        log.exception("voyage_batch_failed", err=str(exc))
        # Mark every record in this batch as failed (so retry counter increments)
        for rec in records:
            db.log_failure(
                dsn=dsn,
                hash_key=rec["hash_key"],
                embedding_version=cfg.voyage_model,
                embed_error=f"voyage_error: {str(exc)[:1000]}",
            )
        return 0, len(records)

    if len(vectors) != len(records):
        log.error("voyage_returned_mismatched_count", expected=len(records), got=len(vectors))
        return 0, len(records)

    # ---- Pinecone upsert
    try:
        pinecone_writer.upsert(ids=ids, vectors=vectors, metadatas=metadatas)
    except Exception as exc:
        log.exception("pinecone_upsert_failed", err=str(exc))
        for rec in records:
            db.log_failure(
                dsn=dsn,
                hash_key=rec["hash_key"],
                embedding_version=cfg.voyage_model,
                embed_error=f"pinecone_error: {str(exc)[:1000]}",
            )
        return 0, len(records)

    # ---- Record success in Postgres
    success_rows = []
    for rec, truncated, tlen in zip(records, truncated_flags, text_lengths):
        success_rows.append({
            "hash_key": rec["hash_key"],
            "embedding_version": cfg.voyage_model,
            "pinecone_id": rec["hash_key"],
            "pinecone_namespace": cfg.pinecone_namespace,
            "text_length": tlen,
            "truncated": truncated,
        })

    db.log_success_batch(dsn, success_rows)
    log.info("batch_complete", embedded=len(records))
    return len(records), 0


def run_loop(args: argparse.Namespace) -> None:
    cfg = WorkerConfig()
    cfg.validate()

    voyageai, Pinecone, ServerlessSpec, retry, stop_after_attempt, wait_exponential = _import_runtime()

    # Build retry decorator with values from cfg
    retry_decorator = retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=2, min=2, max=60),
        reraise=True,
    )

    log.info(
        "worker_starting",
        voyage_model=cfg.voyage_model,
        pinecone_index=cfg.pinecone_index,
        namespace=cfg.pinecone_namespace,
        poll_interval=cfg.poll_interval_sec,
        batch_size=cfg.batch_size,
        dry_run=args.dry_run,
    )

    embedder = VoyageEmbedder(cfg, voyageai, retry_decorator)
    pinecone_writer = PineconeWriter(cfg, Pinecone, ServerlessSpec)

    import raw_embeddings_db as db
    dsn = db.dsn_from_env()

    # Initial status
    status = db.fetch_status(dsn)
    log.info("startup_status", **status)

    exit_flag = GracefulExit()

    while not exit_flag.stop:
        records = db.fetch_pending_records(
            dsn=dsn,
            target_version=cfg.voyage_model,
            limit=cfg.batch_size,
            max_retries=cfg.max_retries,
        )

        if not records:
            log.debug("no_pending_records", sleep_sec=cfg.poll_interval_sec)
            if args.once:
                log.info("once_mode_complete_exiting")
                break
            # Sleep in 1-sec chunks so SIGTERM is responsive
            for _ in range(cfg.poll_interval_sec):
                if exit_flag.stop:
                    break
                time.sleep(1)
            continue

        log.info("pending_batch_found", count=len(records))
        success, failure = process_batch(
            cfg=cfg,
            embedder=embedder,
            pinecone_writer=pinecone_writer,
            records=records,
            dsn=dsn,
            dry_run=args.dry_run,
        )
        log.info("batch_outcome", success=success, failure=failure)

        if args.once:
            # In once mode, drain everything in one go before exiting
            continue

    log.info("worker_stopped")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Raw error embedding worker")
    parser.add_argument("--once", action="store_true",
                        help="Process pending records and exit (no continuous loop)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Log what would happen, no Voyage/Pinecone calls")
    args = parser.parse_args()
    run_loop(args)


if __name__ == "__main__":
    main()
