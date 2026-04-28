"""CLI entrypoint for the IDoc extractor.

Examples::

    # One pass against a real S/4 OData service
    python -m extractors.idoc.cli --once

    # Continuous polling every 60s
    python -m extractors.idoc.cli --watch --interval 60

    # Use mock fixtures (no SAP, no DB-watermark) for smoke testing
    python -m extractors.idoc.cli --once --mock tests/extractors/idoc/fixtures \\
        --in-memory-watermark
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Optional

import structlog

from .mock_client import MockIDocClient
from .odata_client import ODataIDocClient
from .poller import IDocPoller
from .watermark import DBWatermarkStore, MemoryWatermarkStore, WatermarkStore

log = structlog.get_logger(__name__)


def build_poller(args: argparse.Namespace) -> IDocPoller:
    if args.mock:
        connector = MockIDocClient(fixtures_dir=Path(args.mock))
    else:
        connector = ODataIDocClient(
            base_url=args.odata_url or os.getenv("IDOC_ODATA_BASE_URL", ""),
            username=args.username or os.getenv("IDOC_SAP_USERNAME", ""),
            password=args.password or os.getenv("IDOC_SAP_PASSWORD", ""),
            verify_ssl=not args.insecure,
        )

    watermark: WatermarkStore
    if args.in_memory_watermark:
        watermark = MemoryWatermarkStore()
    else:
        watermark = DBWatermarkStore()

    return IDocPoller(
        connector=connector,
        sap_host=args.host,
        sap_sysnr=args.sysnr,
        sap_client=args.client,
        statuses=args.statuses,
        batch_size=args.batch_size,
        watermark_store=watermark,
    )


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m extractors.idoc.cli",
        description="Pull failed SAP IDocs and post them to the Artilegenz ingest endpoint.",
    )
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--once", action="store_true", help="Run one poll cycle and exit")
    mode.add_argument("--watch", action="store_true", help="Run forever on an interval")

    parser.add_argument(
        "--interval", type=int, default=60,
        help="Seconds between polls when --watch (default: 60)",
    )
    parser.add_argument(
        "--mock", default=None,
        help="Use MockIDocClient against the given fixtures directory",
    )
    parser.add_argument("--odata-url", default=None, help="OData base URL (overrides env)")
    parser.add_argument("--username", default=None, help="SAP user (overrides env)")
    parser.add_argument("--password", default=None, help="SAP password (overrides env)")
    parser.add_argument("--insecure", action="store_true", help="Disable TLS cert verify")
    parser.add_argument("--host", default=os.getenv("IDOC_SAP_HOST", ""), help="SAP host id")
    parser.add_argument("--sysnr", default=os.getenv("IDOC_SAP_SYSNR", "00"))
    parser.add_argument("--client", default=os.getenv("IDOC_SAP_CLIENT", "100"))
    parser.add_argument(
        "--statuses", nargs="*", default=None,
        help="Override default failed-status set (e.g. --statuses 51 56)",
    )
    parser.add_argument("--batch-size", type=int, default=50)
    parser.add_argument(
        "--in-memory-watermark", action="store_true",
        help="Use an ephemeral in-memory watermark store (no DB writes)",
    )

    args = parser.parse_args(argv)

    poller = build_poller(args)

    if args.once:
        result = poller.run_once()
        log.info(
            "idoc_cli_done",
            fetched=result.fetched,
            inserted=result.inserted,
            duplicates=result.duplicates,
            watermark_after=str(result.watermark_after),
        )
        print(
            f"fetched={result.fetched} inserted={result.inserted} "
            f"duplicates={result.duplicates} watermark={result.watermark_after}"
        )
        return 0

    poller.run_forever(interval_seconds=args.interval)
    return 0


if __name__ == "__main__":
    sys.exit(main())
