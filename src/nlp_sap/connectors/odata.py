"""SAP OData v2/v4 Connector.

Communicates with SAP S4/HANA via standard OData REST APIs.
Handles authentication (Basic, OAuth2), CSRF tokens, result pagination,
corporate HTTP proxies, and self-signed SSL certificates.
"""

from __future__ import annotations

import asyncio
import base64
import logging
import ssl
import urllib.request
from datetime import date, datetime
from typing import Any
from urllib.parse import urlencode

import httpx
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from nlp_sap.config import AuthType, Settings
from nlp_sap.connectors.base import BaseSAPConnector, SAPQueryRequest, SAPQueryResult

logger = logging.getLogger(__name__)


def _detect_system_proxy(url: str) -> str | None:
    """Read proxy from environment variables or Windows system proxy settings.

    Priority:
      1. HTTPS_PROXY / HTTP_PROXY environment variables
      2. Windows registry / IE proxy (via urllib.request.getproxies())
      3. None (direct connection)
    """
    proxies = urllib.request.getproxies()
    scheme = "https" if url.startswith("https") else "http"
    proxy = proxies.get(scheme) or proxies.get("all")
    if proxy:
        logger.debug("Auto-detected system proxy: %s", proxy)
    return proxy


class ODataConnector(BaseSAPConnector):
    """SAP OData connector using httpx for async HTTP.

    Automatically handles:
    - Corporate HTTP/HTTPS proxies (reads system proxy settings)
    - Self-signed SAP dev certificates (SAP_VERIFY_SSL=false)
    - Basic Auth and OAuth2
    - CSRF token fetch for write operations
    - SAP OData v2 (d.results) and v4 (value) response formats
    """

    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._client: httpx.AsyncClient | None = None
        self._csrf_token: str | None = None
        self._oauth_token: str | None = None

    def _resolve_proxy(self) -> str | None:
        """Return proxy URL string or None (httpx 0.27+ uses proxy=url, not proxies=dict)."""
        # 1. Explicit config takes priority
        if self._settings.sap_proxy:
            proxy_url = self._settings.sap_proxy
            logger.info("Using configured proxy: %s", proxy_url)
            return proxy_url

        # 2. Auto-detect from system / environment
        detected = _detect_system_proxy(self._settings.sap_base_url)
        if detected:
            logger.info("Using auto-detected system proxy: %s", detected)
            return detected

        return None

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None:
            headers = {
                "Accept": "application/json",
                "sap-client": self._settings.sap_client,
                "sap-language": self._settings.sap_language,
            }
            auth = None

            if self._settings.sap_auth_type == AuthType.BASIC:
                auth = (
                    self._settings.sap_username,
                    self._settings.sap_password.get_secret_value(),
                )
            elif self._settings.sap_auth_type == AuthType.OAUTH2:
                await self._refresh_oauth_token()
                headers["Authorization"] = f"Bearer {self._oauth_token}"

            proxy = self._resolve_proxy()

            verify: ssl.SSLContext | bool = (
                self._build_ssl_context() if self._settings.sap_verify_ssl else False
            )
            client_kwargs: dict = dict(
                base_url=self._settings.sap_base_url,
                auth=auth,
                headers=headers,
                verify=verify,
                timeout=httpx.Timeout(30.0, connect=10.0),
                follow_redirects=True,
            )
            if proxy:
                client_kwargs["proxy"] = proxy
            self._client = httpx.AsyncClient(**client_kwargs)
            logger.info(
                "OData client ready → %s (proxy=%s, ssl_verify=%s)",
                self._settings.sap_base_url,
                proxy or "none",
                self._settings.sap_verify_ssl,
            )
        return self._client

    def _build_ssl_context(self) -> ssl.SSLContext:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        if not self._settings.sap_verify_ssl:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        return ctx

    async def _refresh_oauth_token(self) -> None:
        """Fetch a Bearer token from the SAP OAuth2 token endpoint."""
        proxy = self._resolve_proxy()
        verify: ssl.SSLContext | bool = (
            self._build_ssl_context() if self._settings.sap_verify_ssl else False
        )
        oauth_kwargs: dict = dict(verify=verify)
        if proxy:
            oauth_kwargs["proxy"] = proxy
        async with httpx.AsyncClient(**oauth_kwargs) as client:
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

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=8),
        retry=retry_if_exception_type((httpx.TimeoutException, httpx.ConnectError, httpx.NetworkError)),
    )
    async def _fetch(self, url: str) -> dict:
        client = await self._get_client()
        resp = await client.get(url)
        resp.raise_for_status()
        return resp.json()

    async def execute(self, request: SAPQueryRequest) -> SAPQueryResult:
        if request.query_type not in ("odata", "table", "bapi"):
            return SAPQueryResult(
                data=[],
                error=f"ODataConnector received unexpected query_type '{request.query_type}'",
            )

        # For table/bapi query types fall back to RFC_READ_TABLE via OData generic
        # service — only works if OData is the only connector available
        if request.query_type in ("table", "bapi"):
            logger.warning(
                "ODataConnector received query_type='%s' for %s — "
                "attempting OData fallback; consider enabling RFC connector.",
                request.query_type,
                request.table_or_function,
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
            logger.error("OData HTTP %s: %s", exc.response.status_code, exc)
            hint = ""
            if exc.response.status_code == 401:
                hint = " — check SAP_USERNAME / SAP_PASSWORD"
            elif exc.response.status_code == 403:
                hint = " — user lacks OData authorisation (ask Basis for /IWFND/RT_GW_USER)"
            elif exc.response.status_code == 404:
                hint = " — service not activated (run /IWFND/MAINT_SERVICE in SAP)"
            return SAPQueryResult(data=[], error=str(exc) + hint)
        except httpx.ProxyError as exc:
            logger.error("Proxy error: %s", exc)
            return SAPQueryResult(
                data=[],
                error=(
                    f"Proxy error: {exc}. "
                    "Set SAP_PROXY=http://proxy-host:port in .env or leave blank for auto-detect."
                ),
            )
        except httpx.ConnectError as exc:
            logger.error("Connection failed: %s", exc)
            return SAPQueryResult(
                data=[],
                error=(
                    f"Cannot connect to {self._settings.sap_base_url}: {exc}. "
                    "Check SAP_HOST, SAP_HTTP_PORT, and network/proxy settings."
                ),
            )
        except Exception as exc:
            logger.error("OData unexpected error: %s", exc)
            return SAPQueryResult(data=[], error=str(exc))

        # SAP OData v2: {"d": {"results": [...]}}
        # SAP OData v4: {"value": [...]}
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
            metadata={"source": "odata"},
            raw_response=raw,
        )

    async def ping(self) -> bool:
        try:
            client = await self._get_client()
            resp = await client.get(
                "/sap/opu/odata/IWFND/CATALOGSERVICE"
                ";v=2/ServiceCollection?$top=1&$format=json"
            )
            return resp.status_code in (200, 401)
        except Exception as exc:
            logger.warning("OData ping failed: %s", exc)
            return False

    async def close(self) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None
