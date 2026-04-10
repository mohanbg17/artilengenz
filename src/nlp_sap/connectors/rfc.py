"""SAP RFC Connector.

Uses pyrfc (SAP NetWeaver RFC SDK binding) to call BAPIs and RFC_READ_TABLE.
pyrfc is an optional dependency — import errors are caught gracefully.
"""

from __future__ import annotations

import asyncio
import logging
from typing import Any

from nlp_sap.config import Settings
from nlp_sap.connectors.base import BaseSAPConnector, SAPQueryRequest, SAPQueryResult

logger = logging.getLogger(__name__)

# pyrfc is optional; gracefully degrade if SAP NW RFC SDK not installed
try:
    import pyrfc  # type: ignore

    _PYRFC_AVAILABLE = True
except ImportError:
    _PYRFC_AVAILABLE = False
    logger.warning(
        "pyrfc not installed or SAP NW RFC SDK not found. RFC connector disabled. "
        "Install pyrfc and the SAP NW RFC SDK to enable RFC connectivity."
    )


class RFCConnector(BaseSAPConnector):
    """SAP RFC connector — calls BAPIs and RFC_READ_TABLE via pyrfc."""

    # Maximum chars per WHERE row for RFC_READ_TABLE
    _RFC_WHERE_MAX = 72

    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._conn: Any | None = None  # pyrfc.Connection

    def _get_connection(self) -> Any:
        if not _PYRFC_AVAILABLE:
            raise RuntimeError(
                "pyrfc is not available. Install it with: pip install pyrfc\n"
                "Also ensure the SAP NW RFC SDK is installed on this system."
            )
        if self._conn is None or not self._conn.alive:
            self._conn = pyrfc.Connection(**self._settings.rfc_params)
        return self._conn

    def _build_where_options(self, filters: dict[str, Any]) -> list[dict[str, str]]:
        """Convert filter dict to RFC_READ_TABLE OPTIONS rows (≤72 chars each)."""
        parts: list[str] = []
        for field_name, value in filters.items():
            if isinstance(value, list):
                in_vals = ", ".join(f"'{v}'" for v in value)
                parts.append(f"{field_name} IN ({in_vals})")
            elif isinstance(value, dict) and "gte" in value:
                parts.append(f"{field_name} >= '{value['gte']}'")
                if "lte" in value:
                    parts.append(f" AND {field_name} <= '{value['lte']}'")
            else:
                parts.append(f"{field_name} = '{value}'")

        where_str = " AND ".join(parts)

        # Split into ≤72-char chunks for RFC_READ_TABLE OPTIONS table
        options: list[dict[str, str]] = []
        while len(where_str) > self._RFC_WHERE_MAX:
            chunk = where_str[: self._RFC_WHERE_MAX]
            # Try to split at a space boundary
            split_pos = chunk.rfind(" ")
            if split_pos > 0:
                chunk = where_str[:split_pos]
                where_str = where_str[split_pos + 1 :]
            else:
                where_str = where_str[self._RFC_WHERE_MAX :]
            options.append({"TEXT": chunk})
        if where_str:
            options.append({"TEXT": where_str})
        return options

    def _parse_rfc_table_rows(
        self, raw_data: list[dict], delimiter: str, fields: list[dict]
    ) -> list[dict[str, str]]:
        """Parse RFC_READ_TABLE DATA rows into dicts using FIELDS metadata."""
        results: list[dict[str, str]] = []
        for row in raw_data:
            line = row.get("WA", "")
            record: dict[str, str] = {}
            pos = 0
            for f in fields:
                name = f["FIELDNAME"]
                length = int(f["LENGTH"])
                record[name] = line[pos : pos + length].strip()
                pos += length + len(delimiter)
            results.append(record)
        return results

    async def _execute_rfc_read_table(self, request: SAPQueryRequest) -> SAPQueryResult:
        loop = asyncio.get_running_loop()

        def _call() -> dict:
            conn = self._get_connection()
            delimiter = "|"
            fields_param = [{"FIELDNAME": f} for f in request.fields] if request.fields else []
            options = self._build_where_options(request.filters)

            result = conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE=request.table_or_function,
                DELIMITER=delimiter,
                ROWCOUNT=request.max_rows,
                FIELDS=fields_param,
                OPTIONS=options,
            )
            return result

        try:
            raw = await loop.run_in_executor(None, _call)
            fields_meta = raw.get("FIELDS", [])
            delimiter = "|"
            rows = self._parse_rfc_table_rows(raw.get("DATA", []), delimiter, fields_meta)
            return SAPQueryResult(
                data=rows,
                total_count=len(rows),
                has_more=len(rows) >= request.max_rows,
                raw_response=raw,
            )
        except Exception as exc:
            logger.error("RFC_READ_TABLE error for %s: %s", request.table_or_function, exc)
            return SAPQueryResult(data=[], error=str(exc))

    async def _execute_bapi(self, request: SAPQueryRequest) -> SAPQueryResult:
        loop = asyncio.get_running_loop()

        def _call() -> dict:
            conn = self._get_connection()
            return conn.call(request.table_or_function, **request.bapi_import_params)

        try:
            raw = await loop.run_in_executor(None, _call)
            # Gather requested output tables
            data: list[dict] = []
            for table_name in request.bapi_table_params:
                table_rows = raw.get(table_name, [])
                for row in table_rows:
                    data.append(dict(row))

            # Check RETURN messages for errors
            warnings: list[str] = []
            for ret in raw.get("RETURN", []):
                if ret.get("TYPE") in ("E", "A"):
                    return SAPQueryResult(
                        data=[],
                        error=f"BAPI error: {ret.get('MESSAGE')}",
                        raw_response=raw,
                    )
                if ret.get("TYPE") == "W":
                    warnings.append(ret.get("MESSAGE", ""))

            return SAPQueryResult(
                data=data,
                total_count=len(data),
                warnings=warnings,
                raw_response=raw,
            )
        except Exception as exc:
            logger.error("BAPI %s error: %s", request.table_or_function, exc)
            return SAPQueryResult(data=[], error=str(exc))

    async def execute(self, request: SAPQueryRequest) -> SAPQueryResult:
        if request.query_type == "table":
            return await self._execute_rfc_read_table(request)
        elif request.query_type == "bapi":
            return await self._execute_bapi(request)
        else:
            return SAPQueryResult(
                data=[],
                error=f"RFCConnector does not handle query_type='{request.query_type}'",
            )

    async def ping(self) -> bool:
        if not _PYRFC_AVAILABLE:
            return False
        loop = asyncio.get_running_loop()
        try:
            def _ping() -> bool:
                conn = self._get_connection()
                conn.call("RFC_PING")
                return True

            return await loop.run_in_executor(None, _ping)
        except Exception:
            return False

    async def close(self) -> None:
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None
