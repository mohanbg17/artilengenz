"""Mock SAP Connector.

Returns realistic fixture data without needing a live SAP system.
Used in development, testing, and demo environments (MOCK_SAP=true).
Fixtures cover all major intents: GL, AR, AP, PO, Sales Orders, Inventory.
"""

from __future__ import annotations

import asyncio
import random
from datetime import date, timedelta
from typing import Any

from nlp_sap.connectors.base import BaseSAPConnector, SAPQueryRequest, SAPQueryResult


def _random_date(start: date, end: date) -> str:
    delta = (end - start).days
    return (start + timedelta(days=random.randint(0, delta))).isoformat()


def _mock_gl_line_items(filters: dict, max_rows: int) -> list[dict]:
    accounts = ["400000", "401000", "500000", "600100", "700000", "800000"]
    doc_types = ["SA", "RE", "KR", "ZP"]
    rows = []
    for i in range(min(max_rows, 25)):
        amount = round(random.uniform(-500_000, 500_000), 2)
        rows.append({
            "RBUKRS": filters.get("CompanyCode", "1000"),
            "GJAHR": filters.get("FiscalYear", "2024"),
            "BELNR": f"18000{i:05d}",
            "DOCLN": f"{i + 1:06d}",
            "BLART": random.choice(doc_types),
            "BUDAT": _random_date(date(2024, 1, 1), date(2024, 12, 31)),
            "RACCT": random.choice(accounts),
            "RCNTR": f"CC{random.randint(1000, 9999)}",
            "PRCTR": f"PC{random.randint(100, 999)}",
            "HSL": str(amount),
            "RHCUR": "USD",
        })
    return rows


def _mock_ar_open_items(filters: dict, max_rows: int) -> list[dict]:
    rows = []
    for i in range(min(max_rows, 15)):
        due_date = date.today() + timedelta(days=random.randint(-60, 90))
        posting_date = due_date - timedelta(days=30)
        rows.append({
            "BUKRS": filters.get("CompanyCode", "1000"),
            "KUNNR": filters.get("Customer", f"C{10000 + i}"),
            "BELNR": f"14000{i:05d}",
            "GJAHR": "2024",
            "DMBTR": str(round(random.uniform(1000, 200_000), 2)),
            "WAERS": "USD",
            "BUDAT": posting_date.isoformat(),
            "FAEDT": due_date.isoformat(),
            "SGTXT": f"Invoice for services {i + 1}",
        })
    return rows


def _mock_ap_open_items(filters: dict, max_rows: int) -> list[dict]:
    rows = []
    for i in range(min(max_rows, 12)):
        due_date = date.today() + timedelta(days=random.randint(-30, 60))
        posting_date = due_date - timedelta(days=14)
        rows.append({
            "BUKRS": filters.get("CompanyCode", "1000"),
            "LIFNR": filters.get("Vendor", f"V{20000 + i}"),
            "BELNR": f"19000{i:05d}",
            "GJAHR": "2024",
            "DMBTR": str(round(random.uniform(500, 150_000), 2)),
            "WAERS": "USD",
            "BUDAT": posting_date.isoformat(),
            "FAEDT": due_date.isoformat(),
        })
    return rows


def _mock_purchase_orders(filters: dict, max_rows: int) -> list[dict]:
    statuses = ["Open", "Partially Delivered", "Fully Delivered", "Cancelled"]
    rows = []
    for i in range(min(max_rows, 20)):
        rows.append({
            "EBELN": f"4500{100000 + i}",
            "LIFNR": filters.get("Vendor", f"V{20000 + i % 5}"),
            "BUKRS": filters.get("CompanyCode", "1000"),
            "EKORG": "1000",
            "EKGRP": "B01",
            "BSART": "NB",
            "BEDAT": _random_date(date(2024, 1, 1), date(2024, 12, 31)),
            "WAERS": "USD",
            "NETWR": str(round(random.uniform(5_000, 500_000), 2)),
            "STATUS": random.choice(statuses),
            "EINDT": _random_date(date(2024, 3, 1), date(2025, 3, 31)),
        })
    return rows


def _mock_sales_orders(filters: dict, max_rows: int) -> list[dict]:
    rows = []
    for i in range(min(max_rows, 20)):
        rows.append({
            "VBELN": f"1000{100000 + i}",
            "KUNNR": filters.get("Customer", f"C{10000 + i % 8}"),
            "AUART": "OR",
            "VKORG": filters.get("SalesOrganization", "1000"),
            "VTWEG": "10",
            "ERDAT": _random_date(date(2024, 1, 1), date(2024, 12, 31)),
            "WAERK": "USD",
            "NETWR": str(round(random.uniform(10_000, 1_000_000), 2)),
            "DELIVERY_STATUS": random.choice(["A", "B", "C"]),  # A=not delivered
        })
    return rows


def _mock_inventory(filters: dict, max_rows: int) -> list[dict]:
    materials = ["M-001", "M-002", "M-003", "M-100", "M-200", "FG-001", "RM-010"]
    rows = []
    for i, mat in enumerate(materials[:max_rows]):
        stock = round(random.uniform(0, 5000), 0)
        safety = round(random.uniform(100, 500), 0)
        rows.append({
            "MATNR": filters.get("Material", mat),
            "WERKS": filters.get("Plant", "1000"),
            "LGORT": "0001",
            "CLABS": str(stock),   # Unrestricted stock
            "CEINM": str(stock + round(random.uniform(0, 200), 0)),
            "EISBE": str(safety),  # Safety stock
            "BELOW_SAFETY": "X" if stock < safety else "",
        })
    return rows


def _mock_cost_centers(filters: dict, max_rows: int) -> list[dict]:
    rows = []
    centers = [("CC1001", "Finance"), ("CC1002", "HR"), ("CC2001", "Prod-A"), ("CC2002", "Prod-B")]
    for cc, name in centers[:max_rows]:
        actual = round(random.uniform(50_000, 200_000), 2)
        plan = round(actual * random.uniform(0.85, 1.15), 2)
        rows.append({
            "KOSTL": cc,
            "KTEXT": name,
            "KOKRS": "1000",
            "GJAHR": filters.get("FiscalYear", "2024"),
            "ACTUAL_COST": str(actual),
            "PLAN_COST": str(plan),
            "VARIANCE": str(round(actual - plan, 2)),
            "VARIANCE_PCT": str(round((actual - plan) / plan * 100, 1)),
        })
    return rows


# Registry maps intent → fixture generator
_FIXTURE_MAP: dict[str, Any] = {
    "gl_balance_inquiry": _mock_gl_line_items,
    "ar_open_items": _mock_ar_open_items,
    "ap_open_items": _mock_ap_open_items,
    "document_search": _mock_gl_line_items,
    "cost_center_report": _mock_cost_centers,
    "profit_center_report": _mock_gl_line_items,
    "purchase_order_status": _mock_purchase_orders,
    "goods_movement": _mock_purchase_orders,
    "inventory_stock": _mock_inventory,
    "sales_order_status": _mock_sales_orders,
    "revenue_report": _mock_sales_orders,
    "cash_flow": _mock_gl_line_items,
}

# Table-name fallback
_TABLE_MAP: dict[str, Any] = {
    "BKPF": _mock_gl_line_items,
    "BSEG": _mock_gl_line_items,
    "ACDOCA": _mock_gl_line_items,
    "BSID": _mock_ar_open_items,
    "BSIK": _mock_ap_open_items,
    "EKKO": _mock_purchase_orders,
    "VBAK": _mock_sales_orders,
    "MCHB": _mock_inventory,
    "CSKS": _mock_cost_centers,
}


class MockSAPConnector(BaseSAPConnector):
    """Returns deterministic fake SAP data — no live system required."""

    async def execute(self, request: SAPQueryRequest) -> SAPQueryResult:
        # Small artificial latency to simulate network
        await asyncio.sleep(0.05)

        intent = request.metadata.get("intent") if hasattr(request, "metadata") else None
        generator = (
            _FIXTURE_MAP.get(intent or "")
            or _TABLE_MAP.get(request.table_or_function, _mock_gl_line_items)
        )

        rows = generator(request.filters, request.max_rows)

        # Apply field projection if requested
        if request.fields:
            rows = [
                {k: v for k, v in row.items() if k in request.fields} for row in rows
            ]

        return SAPQueryResult(
            data=rows,
            total_count=len(rows),
            has_more=False,
            metadata={"source": "mock", "table": request.table_or_function},
        )

    async def ping(self) -> bool:
        return True

    async def close(self) -> None:
        pass
