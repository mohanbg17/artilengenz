#!/usr/bin/env python3
"""SAP S/4HANA Connection Tester.

Run this script BEFORE starting the NLP-SAP engine to verify your
SAP connection settings are correct.

Usage:
    # With .env file:
    python scripts/test_sap_connection.py

    # With inline args (override .env):
    python scripts/test_sap_connection.py \\
        --host my-sap.example.com \\
        --port 44300 \\
        --client 100 \\
        --user MYUSER \\
        --password MyPass1 \\
        --https

Works WITHOUT admin rights — uses only OData over HTTP/HTTPS.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import socket
import sys
import time
from pathlib import Path

# Add src/ to path so we can import nlp_sap
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

try:
    import httpx
except ImportError:
    print("ERROR: httpx not installed. Run:  pip install httpx")
    sys.exit(1)

# ── ANSI colours ──────────────────────────────────────────────────────────────
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def ok(msg: str)   -> None: print(f"  {GREEN}✓{RESET} {msg}")
def fail(msg: str) -> None: print(f"  {RED}✗{RESET} {msg}")
def warn(msg: str) -> None: print(f"  {YELLOW}!{RESET} {msg}")
def info(msg: str) -> None: print(f"  {CYAN}→{RESET} {msg}")


# ── TCP port probe ─────────────────────────────────────────────────────────────
def tcp_reachable(host: str, port: int, timeout: float = 5.0) -> bool:
    try:
        sock = socket.create_connection((host, port), timeout=timeout)
        sock.close()
        return True
    except (socket.timeout, ConnectionRefusedError, OSError):
        return False


# ── Auto-discover ICM port ────────────────────────────────────────────────────
async def probe_icm_ports(host: str) -> list[tuple[int, bool]]:
    """Probe common SAP ICM ports and return [(port, is_https)] for open ones."""
    candidates = [
        (44300, True),   # S/4HANA HTTPS default (SYS 00)
        (44301, True),   # SYS 01
        (443,   True),   # Standard HTTPS (load balancer / reverse proxy)
        (8443,  True),   # Alternative HTTPS
        (8000,  False),  # S/4HANA HTTP default (SYS 00)
        (8001,  False),  # SYS 01
        (8080,  False),  # Alternative HTTP
        (50000, False),  # SAP HANA XS classic
        (50001, True),   # SAP HANA XS classic HTTPS
    ]
    open_ports = []
    for port, https in candidates:
        if tcp_reachable(host, port, timeout=2.0):
            open_ports.append((port, https))
    return open_ports


# ── OData endpoint probes ─────────────────────────────────────────────────────
ODATA_PROBE_PATHS = [
    # Service catalogue — always present on S/4HANA
    "/sap/opu/odata/IWFND/CATALOGSERVICE;v=2/ServiceCollection?$top=1&$format=json",
    # Simpler ping via OData metadata
    "/sap/opu/odata/IWFND/CATALOGSERVICE?$format=json",
    # SAP Web Dispatcher health check
    "/sap/public/ping",
    # ICM server info
    "/sap/bc/ping",
]


async def test_odata_anonymous(base_url: str) -> tuple[bool, str]:
    """Try OData endpoints without auth — even a 401 means OData is reachable."""
    async with httpx.AsyncClient(verify=False, timeout=10.0, follow_redirects=True) as client:
        for path in ODATA_PROBE_PATHS:
            try:
                resp = await client.get(f"{base_url}{path}")
                if resp.status_code in (200, 401, 403):
                    return True, f"HTTP {resp.status_code} on {path}"
                if resp.status_code == 404:
                    continue  # Try next path
            except Exception as exc:
                continue
    return False, "All probe paths returned 404 or failed"


async def test_odata_auth(base_url: str, client_no: str, user: str, password: str) -> dict:
    """Test OData with credentials — fetch service catalogue."""
    url = (
        f"{base_url}/sap/opu/odata/IWFND/CATALOGSERVICE;v=2"
        f"/ServiceCollection?$top=5&$format=json"
    )
    results: dict = {"url": url, "authenticated": False, "services": [], "error": None}

    async with httpx.AsyncClient(
        verify=False,
        timeout=15.0,
        follow_redirects=True,
        headers={"sap-client": client_no, "Accept": "application/json"},
        auth=(user, password),
    ) as client:
        try:
            t0 = time.monotonic()
            resp = await client.get(url)
            elapsed = round((time.monotonic() - t0) * 1000)

            results["http_status"] = resp.status_code
            results["elapsed_ms"] = elapsed

            if resp.status_code == 200:
                results["authenticated"] = True
                data = resp.json()
                services = data.get("d", {}).get("results", data.get("value", []))
                results["services"] = [s.get("TechnicalServiceName", s.get("ID", "?"))
                                       for s in services[:5]]
            elif resp.status_code == 401:
                results["error"] = "Authentication failed — check username/password"
            elif resp.status_code == 403:
                results["error"] = "Access denied — user lacks OData authorisation"
            elif resp.status_code == 404:
                results["error"] = "CATALOGSERVICE not found — check SAP_ODATA_BASE_PATH"
            else:
                results["error"] = f"Unexpected HTTP {resp.status_code}: {resp.text[:200]}"

        except httpx.ConnectError as exc:
            results["error"] = f"Connection refused: {exc}"
        except httpx.TimeoutException:
            results["error"] = "Connection timed out"
        except Exception as exc:
            results["error"] = str(exc)

    return results


async def test_specific_service(
    base_url: str, client_no: str, user: str, password: str, service: str, entity: str
) -> dict:
    """Test a specific OData service used by the NLP engine."""
    url = f"{base_url}/sap/opu/odata/sap/{service}/{entity}?$top=1&$format=json"
    async with httpx.AsyncClient(
        verify=False, timeout=10.0,
        headers={"sap-client": client_no, "Accept": "application/json"},
        auth=(user, password),
    ) as client:
        try:
            resp = await client.get(url)
            if resp.status_code == 200:
                return {"ok": True, "service": service, "entity": entity}
            return {
                "ok": False, "service": service, "entity": entity,
                "status": resp.status_code,
                "hint": "Service not activated" if resp.status_code == 404 else resp.text[:100],
            }
        except Exception as exc:
            return {"ok": False, "service": service, "entity": entity, "error": str(exc)}


# ── Main ──────────────────────────────────────────────────────────────────────
async def main(args: argparse.Namespace) -> int:
    print(f"\n{BOLD}{'='*60}{RESET}")
    print(f"{BOLD}  SAP S/4HANA Connection Tester{RESET}")
    print(f"{BOLD}{'='*60}{RESET}\n")

    host     = args.host
    port     = args.port
    use_https = args.https
    client_no = args.client
    user     = args.user
    password = args.password

    # ── Step 1: TCP reachability ────────────────────────────────────────────
    print(f"{BOLD}Step 1: TCP Reachability{RESET}")
    if port:
        reachable = tcp_reachable(host, port)
        if reachable:
            ok(f"{host}:{port} is reachable")
        else:
            fail(f"{host}:{port} is NOT reachable")
            warn("Trying to auto-discover open SAP ports...")
            open_ports = await probe_icm_ports(host)
            if open_ports:
                for p, h in open_ports:
                    ok(f"Port {p} is open ({'HTTPS' if h else 'HTTP'})")
                port, use_https = open_ports[0]
                warn(f"Using auto-discovered port {port}")
            else:
                fail("No SAP ICM ports reachable from this machine")
                print()
                print("  This means SAP is only accessible from inside the Remote Desktop.")
                print("  See docs/CONNECT_RDP.md for setup instructions.")
                return 1
    else:
        warn("No port specified — probing common SAP ICM ports...")
        open_ports = await probe_icm_ports(host)
        if open_ports:
            for p, h in open_ports:
                ok(f"Port {p} is open ({'HTTPS' if h else 'HTTP'})")
            port, use_https = open_ports[0]
            info(f"Will use port {port} ({'HTTPS' if use_https else 'HTTP'})")
        else:
            fail(f"No SAP ICM ports found open on {host}")
            return 1
    print()

    scheme   = "https" if use_https else "http"
    base_url = f"{scheme}://{host}:{port}"

    # ── Step 2: OData reachability (no auth) ───────────────────────────────
    print(f"{BOLD}Step 2: OData Endpoint Reachability{RESET}")
    reachable, detail = await test_odata_anonymous(base_url)
    if reachable:
        ok(f"OData endpoint is up — {detail}")
    else:
        fail(f"OData endpoint not reachable — {detail}")
        warn("The SAP ICM may have OData disabled. Check SICF in SAP.")
        return 1
    print()

    # ── Step 3: Authenticated OData ────────────────────────────────────────
    if user and password:
        print(f"{BOLD}Step 3: Authenticated OData (CATALOGSERVICE){RESET}")
        auth_result = await test_odata_auth(base_url, client_no, user, password)
        if auth_result["authenticated"]:
            ok(f"Authenticated as '{user}' on client {client_no} ({auth_result['elapsed_ms']}ms)")
            if auth_result["services"]:
                ok(f"Sample services: {', '.join(auth_result['services'])}")
        else:
            fail(f"Authentication failed: {auth_result.get('error')}")
            return 1
        print()

        # ── Step 4: Check services used by NLP engine ──────────────────────
        print(f"{BOLD}Step 4: NLP Engine OData Services{RESET}")
        services_to_test = [
            ("API_GLACCOUNTLINEITEM_SRV", "A_GLAccountLineItem"),
            ("API_PURCHASEORDER_PROCESS_SRV", "A_PurchaseOrder"),
            ("API_SALES_ORDER_SRV", "A_SalesOrder"),
        ]
        all_ok = True
        for svc, entity in services_to_test:
            result = await test_specific_service(base_url, client_no, user, password, svc, entity)
            if result["ok"]:
                ok(f"{svc} / {entity}")
            else:
                status = result.get("status", "ERR")
                hint   = result.get("hint", result.get("error", ""))
                warn(f"{svc} — HTTP {status}: {hint}")
                if status == 404:
                    info(f"  Activate in SAP: /IWFND/MAINT_SERVICE → add '{svc}'")
                all_ok = False

        if not all_ok:
            warn("Some services are unavailable — NLP engine will fall back to BAPI/Table reads")
        print()
    else:
        warn("No credentials provided — skipping auth and service tests")
        print()

    # ── Step 5: Print .env config ──────────────────────────────────────────
    print(f"{BOLD}Step 5: Recommended .env Configuration{RESET}")
    print()
    print("  Copy the following into your .env file:\n")
    print(f"  SAP_HOST={host}")
    print(f"  SAP_HTTP_PORT={port}")
    print(f"  SAP_HTTPS={'true' if use_https else 'false'}")
    print(f"  SAP_CLIENT={client_no}")
    print(f"  SAP_AUTH_TYPE=basic")
    print(f"  SAP_USERNAME={user or '<your-sap-user>'}")
    print(f"  SAP_PASSWORD={password or '<your-sap-password>'}")
    print(f"  SAP_SYSTEM_TYPE=S4HANA")
    print(f"  SAP_LANGUAGE=EN")
    print(f"  MOCK_SAP=false")
    print()
    ok("Connection test complete — ready to configure .env")
    print()
    return 0


def parse_args() -> argparse.Namespace:
    # Load .env defaults
    env_file = Path(__file__).parent.parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                os.environ.setdefault(k.strip(), v.strip())

    parser = argparse.ArgumentParser(description="Test SAP S/4HANA connectivity")
    parser.add_argument("--host",     default=os.getenv("SAP_HOST", ""),       help="SAP hostname or IP")
    parser.add_argument("--port",     default=int(os.getenv("SAP_HTTP_PORT") or 0) or None, type=int)
    parser.add_argument("--client",   default=os.getenv("SAP_CLIENT", "100"),  help="SAP client (3 digits)")
    parser.add_argument("--user",     default=os.getenv("SAP_USERNAME", ""),   help="SAP username")
    parser.add_argument("--password", default=os.getenv("SAP_PASSWORD", ""),   help="SAP password")
    parser.add_argument("--https",    default=os.getenv("SAP_HTTPS", "true").lower() == "true",
                        action=argparse.BooleanOptionalAction)
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    if not args.host:
        print(f"{RED}ERROR: No SAP host specified.{RESET}")
        print("  Either set SAP_HOST in your .env file, or run:")
        print("  python scripts/test_sap_connection.py --host <SAP-HOST>")
        sys.exit(1)
    sys.exit(asyncio.run(main(args)))
