"""Abstract connector for fetching IDocs from SAP.

Implementations:
  * odata_client.ODataIDocClient -- reads from S/4 standard service (API_IDOC_SRV
    or equivalent) over HTTPS+OData v2. No code in SAP required.
  * mock_client.MockIDocClient   -- returns IDoc objects loaded from JSON
    fixtures, used for tests and local development without a live SAP system.

A future CDS-based connector would slot in here without changes downstream.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime
from typing import Iterable, List, Optional, Sequence

from .models import IDoc


class IDocConnector(ABC):
    """Read-side interface to a source of failed IDocs."""

    @abstractmethod
    def list_failed_idocs(
        self,
        *,
        since: Optional[datetime] = None,
        statuses: Optional[Sequence[str]] = None,
        max_results: int = 100,
    ) -> List[IDoc]:
        """Return failed IDocs created/changed strictly after `since`.

        - `since` is None on the first run (cold start).
        - `statuses` filters to specific status codes; None means caller-default.
        - `max_results` is a hard cap to prevent runaway pulls.
        Implementations must return results sorted by created_at ascending so
        the watermark advance is monotonic.
        """
        raise NotImplementedError
