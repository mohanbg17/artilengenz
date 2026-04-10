"""SAP OData v2/v4 Connector.

Communicates with SAP S4/HANA via standard OData REST APIs.
Handles authentication (Basic, OAuth2), CSRF tokens, and result pagination.
"""

from __future__ import annotations

import asyncio
import base64
import logging
from datetime import date, datetime
from typing import Any
from urllib.parse import urlencode, urljoin

import httpx
from tenacity import retry, stop_after_attempt, wait_exponential

from nlp_sap.config import AuthType, Settings
from nlp_sap.connectors.base import BaseSAPConnector, SAPQueryRequest, SAPQueryResult

logger = logging.getLogger(__name__)


class ODataConnector(BaseSAPConnector):
    """SAP OData connector using httpx for async HTTP."""

    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._client: httpx.AsyncClient | None = None
        self._csrf_token: str | None = None
        self._oauth_token: str | None = None

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None:
            headers = {"Accept": "application/json", "sap-client": self._settings.sap_client}
            auth = None

            if self._settings.sap_auth_type == AuthType.BASIC:
                auth = (
                    self._settings.sap_username,
                    self._settings.sap_password.get_secret_value(),
                )
            elif self._settings.sap_auth_type == AuthType.OAUTH2:
                await self._refresh_oauth_token()
                headers["Authorization"] = f"Bearer {self._oauth_token}"

            self._client = httpx.AsyncClient(
                base_url=self._settings.sap_base_url,
                auth=auth,
                headers=headers,
                verify=True,
                timeout=httpx.Timeout(30.0, connect=10.0),
                follow_redirects=True,
            )
        return self._client

    async def _refresh_oauth_token(self) -> None:
        """Fetch a Bearer token from the SAP OAuth2 token endpoint."""
        async with httpx.AsyncClient() as client:
            cred = base64.b64encode(
                f"{self._settings.sap_client_id}:"
                f"{self._settings.sap_client_secret.get_secret_value()}".encode()
            ).decode()
            resp = await client.post(
                self._settings.sap_token_url,
                headers={"Authorization": f"Basic {cred}"},
                data={"grant_type": "client_credentials"},
            )
            resp.raise_for_status()
            self._oauth_token = resp.json()["access_token"]

    async def _get_csrf_token(self, service_path: str) -> str:
        """Fetch a CSRF token (required for write operations)."""
        if self._csrf_token:
            return self._csrf_token
        client = await self._get_client()
        resp = await client.get(service_path, headers={"x-csrf-token": "Fetch"})
        token = resp.headers.get("x-csrf-token", "")
        self._csrf_token = token
        return token

    def _build_odata_url(
        self,
        service: str,
        entity_set: str,
        filters: dict[str, Any],
        fields: list[str],
        top: int,
        expand: list[str] | None = None,
    ) -> str:
        """Construct an OData query URL with $filter, $select, $top."""
        base = f"{self._settings.sap_odata_base_path}/{service}/{entity_set}"

        params: dict[str, str] = {"$format": "json", "$top": str(top)}

        if fields:
            params["$select"] = ",".join(fields)

        if expand:
            params["$expand"] = ",".join(expand)

        filter_parts: list[str] = []
        for field_name, value in filters.items():
            if isinstance(value, list):
                # IN clause → expand to multiple eq with 'or'
                or_parts = [f"{field_name} eq '{v}'" for v in value]
                filter_parts.append("(" + " or ".join(or_parts) + ")")
            elif isinstance(value, dict) and "gte" in value:
                filter_parts.append(f"{field_name} ge '{value['gte']}'")
                if "lte" in value:
                    filter_parts.append(f"{field_name} le '{value['lte']}'")
            elif isinstance(value, (date, datetime)):
                filter_parts.append(f"{field_name} eq datetime'{value.isoformat()}'")
            else:
                filter_parts.append(f"{field_name} eq '{value}'")

        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)

        return f"{base}?{urlencode(params)}"

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=1, max=8))
    async def _fetch(self, url: str) -> dict:
        client = await self._get_client()
        resp = await client.get(url)
        resp.raise_for_status()
        return resp.json()

    async def execute(self, request: SAPQueryRequest) -> SAPQueryResult:
        if request.query_type != "odata":
            return SAPQueryResult(
                data=[],
                error=f"ODataConnector only handles 'odata' queries, got '{request.query_type}'",
            )

        service = request.odata_service or ""
        entity_set = request.odata_entity_set or request.table_or_function

        try:
            url = self._build_odata_url(
                service=service,
                entity_set=entity_set,
                filters=request.filters,
                fields=request.fields,
                top=request.max_rows,
                expand=request.odata_expand or None,
            )
            logger.debug("OData GET %s", url)
            raw = await self._fetch(url)
        except httpx.HTTPStatusError as exc:
            logger.error("OData HTTP error: %s", exc)
            return SAPQueryResult(data=[], error=str(exc))
        except Exception as exc:
            logger.error("OData unexpected error: %s", exc)
            return SAPQueryResult(data=[], error=str(exc))

        # SAP OData v2 returns {"d": {"results": [...]}}
        # SAP OData v4 returns {"value": [...]}
        rows: list[dict] = []
        if "d" in raw:
            rows = raw["d"].get("results", [])
        elif "value" in raw:
            rows = raw["value"]
        else:
            rows = [raw]

        total = raw.get("d", {}).get("__count") or len(rows)
        return SAPQueryResult(
            data=rows,
            total_count=int(total),
            has_more=len(rows) >= request.max_rows,
            raw_response=raw,
        )

    async def ping(self) -> bool:
        try:
            client = await self._get_client()
            resp = await client.get(self._settings.sap_odata_base_path)
            return resp.status_code < 400
        except Exception:
            return False

    async def close(self) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None
