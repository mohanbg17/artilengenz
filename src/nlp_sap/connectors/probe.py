"""SAP Connectivity Probe.

Lightweight async checks used by the /api/v1/connect endpoint and the
startup health-check.  Works without admin rights — OData over HTTP/HTTPS only.
"""

from __future__ import annotations

import asyncio
import socket
import time
from dataclasses import dataclass, field

import httpx

from nlp_sap.config import Settings


@dataclass
class ProbeResult:
    host: str
    port: int
    tcp_reachable: bool = False
    odata_reachable: bool = False
    authenticated: bool = False
    services_ok: list[str] = field(default_factory=list)
    services_missing: list[str] = field(default_factory=list)
    elapsed_ms: float = 0.0
    error: str | None = None
    suggested_env: dict[str, str] = field(default_factory=dict)

    @property
    def ready(self) -> bool:
        return self.authenticated


def _tcp_check(host: str, port: int, timeout: float = 4.0) -> bool:
    try:
        s = socket.create_connection((host, port), timeout=timeout)
        s.close()
        return True
    except OSError:
        return False


_COMMON_SAP_PORTS: list[tuple[int, bool]] = [
    (44300, True),
    (44301, True),
    (443,   True),
    (8443,  True),
    (8000,  False),
    (8001,  False),
    (8080,  False),
]

# NLP engine OData services that need to be activated in SAP
_NLP_SERVICES: list[tuple[str, str]] = [
    ("API_GLACCOUNTLINEITEM_SRV",   "A_GLAccountLineItem"),
    ("API_PURCHASEORDER_PROCESS_SRV", "A_PurchaseOrder"),
    ("API_SALES_ORDER_SRV",         "A_SalesOrder"),
]


async def probe(settings: Settings) -> ProbeResult:
    host = settings.sap_host
    port = settings.sap_http_port
    use_https = settings.sap_https
    t0 = time.monotonic()

    result = ProbeResult(host=host, port=port)

    # ── TCP ───────────────────────────────────────────────────────────────────
    loop = asyncio.get_running_loop()
    result.tcp_reachable = await loop.run_in_executor(None, _tcp_check, host, port)

    if not result.tcp_reachable:
        # Try auto-discovery on common ports
        for p, https in _COMMON_SAP_PORTS:
            ok = await loop.run_in_executor(None, _tcp_check, host, p, 2.0)
            if ok:
                port = p
                use_https = https
                result.tcp_reachable = True
                result.port = port
                break

    if not result.tcp_reachable:
        result.error = (
            f"Cannot reach {host} on any SAP ICM port. "
            "SAP may only be accessible from inside the Remote Desktop. "
            "See docs/CONNECT_RDP.md for tunnelling instructions."
        )
        result.elapsed_ms = round((time.monotonic() - t0) * 1000, 1)
        return result

    scheme   = "https" if use_https else "http"
    base_url = f"{scheme}://{host}:{port}"
    headers  = {
        "Accept":     "application/json",
        "sap-client": settings.sap_client,
    }
    auth = (settings.sap_username, settings.sap_password.get_secret_value())

    async with httpx.AsyncClient(
        verify=False,
        timeout=httpx.Timeout(15.0, connect=5.0),
        follow_redirects=True,
        headers=headers,
        auth=auth,
    ) as client:

        # ── OData anonymous reachability (401 = up, 200 = up without auth) ──
        try:
            r = await client.get(
                f"{base_url}/sap/opu/odata/IWFND/CATALOGSERVICE"
                f";v=2/ServiceCollection?$top=1&$format=json"
            )
            result.odata_reachable = r.status_code in (200, 401, 403)
        except Exception as exc:
            result.error = f"OData probe failed: {exc}"
            result.elapsed_ms = round((time.monotonic() - t0) * 1000, 1)
            return result

        # ── Authenticated catalogue fetch ─────────────────────────────────
        if result.odata_reachable:
            try:
                r = await client.get(
                    f"{base_url}/sap/opu/odata/IWFND/CATALOGSERVICE"
                    f";v=2/ServiceCollection?$top=1&$format=json"
                )
                result.authenticated = r.status_code == 200
                if r.status_code == 401:
                    result.error = "Authentication failed — check SAP_USERNAME / SAP_PASSWORD"
                elif r.status_code == 403:
                    result.error = "Authorisation denied — user lacks /IWFND/CATALOGSERVICE rights"
            except Exception as exc:
                result.error = str(exc)

        # ── Per-service checks ─────────────────────────────────────────────
        if result.authenticated:
            for svc, entity in _NLP_SERVICES:
                url = (
                    f"{base_url}/sap/opu/odata/sap/{svc}/{entity}"
                    f"?$top=1&$format=json"
                )
                try:
                    r = await client.get(url)
                    if r.status_code == 200:
                        result.services_ok.append(svc)
                    else:
                        result.services_missing.append(svc)
                except Exception:
                    result.services_missing.append(svc)

    result.elapsed_ms = round((time.monotonic() - t0) * 1000, 1)

    # Build a suggested .env block
    result.suggested_env = {
        "SAP_HOST":        host,
        "SAP_HTTP_PORT":   str(port),
        "SAP_HTTPS":       "true" if use_https else "false",
        "SAP_CLIENT":      settings.sap_client,
        "SAP_AUTH_TYPE":   "basic",
        "SAP_USERNAME":    settings.sap_username,
        "SAP_SYSTEM_TYPE": "S4HANA",
        "MOCK_SAP":        "false",
    }

    return result
