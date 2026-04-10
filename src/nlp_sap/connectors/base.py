"""Abstract base class for all SAP connectors."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any


@dataclass
class SAPQueryRequest:
    """Unified query request fed to any connector."""

    # What to fetch
    table_or_function: str          # Table name, BAPI name, or OData entity set
    query_type: str                  # "table" | "bapi" | "odata"

    # Filters / parameters
    filters: dict[str, Any] = field(default_factory=dict)  # field → value(s)
    fields: list[str] = field(default_factory=list)         # columns to return
    max_rows: int = 500

    # Optional join
    join_table: str | None = None
    join_on: dict[str, str] | None = None   # {left_field: right_field}

    # OData extras
    odata_service: str | None = None
    odata_entity_set: str | None = None
    odata_expand: list[str] = field(default_factory=list)

    # RFC / BAPI extras
    bapi_import_params: dict[str, Any] = field(default_factory=dict)
    bapi_table_params: list[str] = field(default_factory=list)  # which output tables


@dataclass
class SAPQueryResult:
    """Unified query result returned by any connector."""

    data: list[dict[str, Any]]          # list of row dicts
    total_count: int = 0
    has_more: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)
    raw_response: Any = None
    error: str | None = None
    warnings: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return self.error is None

    def to_dict(self) -> dict:
        return {
            "data": self.data,
            "total_count": self.total_count,
            "has_more": self.has_more,
            "metadata": self.metadata,
            "error": self.error,
            "warnings": self.warnings,
        }


class BaseSAPConnector(ABC):
    """Abstract connector — subclasses implement execute()."""

    @abstractmethod
    async def execute(self, request: SAPQueryRequest) -> SAPQueryResult:
        """Execute a query against SAP and return a unified result."""

    @abstractmethod
    async def ping(self) -> bool:
        """Check connectivity to the SAP system."""

    @abstractmethod
    async def close(self) -> None:
        """Release any held resources."""
