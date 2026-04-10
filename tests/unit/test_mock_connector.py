"""Unit tests for the MockSAPConnector."""

import pytest

from nlp_sap.connectors.base import SAPQueryRequest
from nlp_sap.connectors.mock import MockSAPConnector


@pytest.fixture
def connector() -> MockSAPConnector:
    return MockSAPConnector()


def _req(table: str, intent: str = "", query_type: str = "table", **kwargs) -> SAPQueryRequest:
    req = SAPQueryRequest(table_or_function=table, query_type=query_type, **kwargs)
    req.__dict__["metadata"] = {"intent": intent}
    return req


class TestMockConnector:
    @pytest.mark.asyncio
    async def test_ping_returns_true(self, connector: MockSAPConnector) -> None:
        assert await connector.ping() is True

    @pytest.mark.asyncio
    async def test_gl_balance_returns_rows(self, connector: MockSAPConnector) -> None:
        req = _req("ACDOCA", intent="gl_balance_inquiry", filters={"CompanyCode": "1000"})
        result = await connector.execute(req)
        assert result.success
        assert len(result.data) > 0

    @pytest.mark.asyncio
    async def test_ar_open_items_returns_rows(self, connector: MockSAPConnector) -> None:
        req = _req("BSID", intent="ar_open_items", filters={"KUNNR": "C10001"})
        result = await connector.execute(req)
        assert result.success
        assert len(result.data) > 0
        assert all("KUNNR" in row or "BELNR" in row for row in result.data)

    @pytest.mark.asyncio
    async def test_purchase_orders_returns_rows(self, connector: MockSAPConnector) -> None:
        req = _req("EKKO", intent="purchase_order_status")
        result = await connector.execute(req)
        assert result.success
        assert len(result.data) > 0
        assert "EBELN" in result.data[0]

    @pytest.mark.asyncio
    async def test_inventory_returns_rows(self, connector: MockSAPConnector) -> None:
        req = _req("MCHB", intent="inventory_stock", filters={"WERKS": "1000"})
        result = await connector.execute(req)
        assert result.success
        assert len(result.data) > 0
        assert "MATNR" in result.data[0]

    @pytest.mark.asyncio
    async def test_max_rows_respected(self, connector: MockSAPConnector) -> None:
        req = _req("ACDOCA", intent="gl_balance_inquiry", max_rows=5)
        result = await connector.execute(req)
        assert len(result.data) <= 5

    @pytest.mark.asyncio
    async def test_field_projection(self, connector: MockSAPConnector) -> None:
        req = _req("EKKO", intent="purchase_order_status", fields=["EBELN", "LIFNR"])
        result = await connector.execute(req)
        if result.data:
            for row in result.data:
                for key in row:
                    assert key in ["EBELN", "LIFNR"]

    @pytest.mark.asyncio
    async def test_close_is_noop(self, connector: MockSAPConnector) -> None:
        await connector.close()  # Should not raise

    @pytest.mark.asyncio
    async def test_sales_order_returns_rows(self, connector: MockSAPConnector) -> None:
        req = _req("VBAK", intent="sales_order_status")
        result = await connector.execute(req)
        assert result.success
        assert "VBELN" in result.data[0]
