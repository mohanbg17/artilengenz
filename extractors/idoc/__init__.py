"""IDoc extractor — pulls failed SAP IDocs into the error intelligence platform.

Public surface:
  * :class:`IDoc`, :class:`IDocStatus`, :class:`IDocSegment` -- normalized models
  * :class:`IDocConnector` -- ABC implemented by ``ODataIDocClient`` and ``MockIDocClient``
  * :class:`IDocPoller` -- watermark-driven pull-and-post loop
  * :func:`idoc_to_ingest_record` -- one-shot transform for the ingest endpoint
"""
from .connector import IDocConnector
from .mock_client import MockIDocClient
from .models import INBOUND, OUTBOUND, IDoc, IDocSegment, IDocStatus
from .odata_client import ODataIDocClient
from .poller import IDocPoller, PollResult
from .severity import (
    DEFAULT_FAILED_STATUSES,
    DEFAULT_FAILED_STATUSES_INBOUND,
    DEFAULT_FAILED_STATUSES_OUTBOUND,
    is_failure,
    severity_for,
)
from .transformer import idoc_to_ingest_record
from .watermark import DBWatermarkStore, MemoryWatermarkStore, WatermarkStore

__all__ = [
    "DBWatermarkStore",
    "DEFAULT_FAILED_STATUSES",
    "DEFAULT_FAILED_STATUSES_INBOUND",
    "DEFAULT_FAILED_STATUSES_OUTBOUND",
    "INBOUND",
    "IDoc",
    "IDocConnector",
    "IDocPoller",
    "IDocSegment",
    "IDocStatus",
    "MemoryWatermarkStore",
    "MockIDocClient",
    "ODataIDocClient",
    "OUTBOUND",
    "PollResult",
    "WatermarkStore",
    "idoc_to_ingest_record",
    "is_failure",
    "severity_for",
]
