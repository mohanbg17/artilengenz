"""ODataIDocClient against a mocked httpx transport.

Validates the OData query construction and the JSON parsing path on a real
HTTP boundary (no live SAP). Uses fixture JSON directly as the canned response.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import httpx
import pytest

from extractors.idoc import INBOUND, ODataIDocClient
from extractors.idoc.odata_client import (
    _odata_v2_datetime,
    _parse_odata_datetime,
)


FIXTURES = Path(__file__).parent / "fixtures"


def _load_fixture_v2(filename: str) -> dict:
    return json.loads((FIXTURES / filename).read_text(encoding="utf-8"))


def test_v2_datetime_literal_format():
    ts = datetime(2025, 4, 15, 10, 23, 45)
    assert _odata_v2_datetime(ts) == "datetime'2025-04-15T10:23:45'"


def test_parse_odata_v2_ticks():
    parsed = _parse_odata_datetime("/Date(1700000000000)/")
    # 2023-11-14 22:13:20 UTC
    assert parsed.year == 2023
    assert parsed.month == 11


def test_parse_iso8601_with_z():
    parsed = _parse_odata_datetime("2025-04-15T10:23:45Z")
    assert parsed == datetime(2025, 4, 15, 10, 23, 45)


def test_parse_invalid_datetime_returns_now_safely():
    # Should not raise, should return *some* datetime
    parsed = _parse_odata_datetime("not-a-date")
    assert isinstance(parsed, datetime)


def test_query_params_include_filter_orderby_top_and_expand():
    client = ODataIDocClient(
        base_url="http://x/",
        http_client=httpx.Client(transport=httpx.MockTransport(lambda r: httpx.Response(200, json={"d": {"results": []}}))),
    )
    params = client._build_query_params(
        since=datetime(2025, 4, 1),
        statuses=["51", "56"],
        top=25,
    )
    assert params["$top"] == "25"
    assert params["$format"] == "json"
    assert params["$orderby"].endswith("asc")
    assert "(Status eq '51' or Status eq '56')" in params["$filter"]
    assert "CreationDateTime gt datetime'2025-04-01T00:00:00'" in params["$filter"]
    assert "IdocStatusRecord" in params["$expand"]


def test_full_request_response_with_mock_transport():
    captured = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["url"] = str(request.url)
        captured["params"] = dict(request.url.params)
        return httpx.Response(200, json=_load_fixture_v2("idoc_orders05_status51.json"))

    transport = httpx.MockTransport(handler)
    http = httpx.Client(transport=transport, headers={"Accept": "application/json"})

    client = ODataIDocClient(
        base_url="https://my-s4.example.com/sap/opu/odata/sap/API_IDOC_SRV",
        http_client=http,
    )
    idocs = client.list_failed_idocs(statuses=["51"], max_results=10)

    assert len(idocs) == 1
    idoc = idocs[0]
    assert idoc.docnum == "0000000000123451"
    assert idoc.idoctp == "ORDERS05"
    assert idoc.direction == INBOUND
    assert idoc.status == "51"
    # The request must have hit the right entity
    assert "Idoc" in captured["url"]
    assert captured["params"]["$top"] == "10"


def test_4xx_response_raises():
    transport = httpx.MockTransport(lambda r: httpx.Response(403, text="forbidden"))
    http = httpx.Client(transport=transport)
    client = ODataIDocClient(base_url="http://x/", http_client=http)
    with pytest.raises(httpx.HTTPStatusError):
        client.list_failed_idocs()


def test_v4_payload_with_value_array_also_parsed():
    payload = {
        "value": [
            {
                "IdocNumber": "999",
                "Status": "51",
                "CreationDateTime": "2025-04-15T10:00:00Z",
                "Direction": "2",
                "MessageType": "ORDERS",
                "BasicType": "ORDERS05",
                "SenderPartnerNumber": "P1",
                "ReceiverPartnerNumber": "P2",
                "Client": "100",
                "IdocStatusRecord": [{"Counter": 1, "Status": "51", "StatusText": "x", "LogDateTime": "2025-04-15T10:01:00Z"}],
                "IdocSegment": [],
            }
        ]
    }
    transport = httpx.MockTransport(lambda r: httpx.Response(200, json=payload))
    http = httpx.Client(transport=transport)
    client = ODataIDocClient(base_url="http://x/", http_client=http)
    idocs = client.list_failed_idocs()
    assert len(idocs) == 1
    assert idocs[0].docnum == "999"
