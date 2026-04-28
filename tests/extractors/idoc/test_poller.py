"""End-to-end IDoc poller test using a mock connector and mocked HTTP transport.

Does NOT touch the database — uses ``MemoryWatermarkStore``.
Does NOT call live SAP — uses ``MockIDocClient`` against the JSON fixtures.
Does NOT call real ingest — intercepts the POST via ``httpx.MockTransport``.
"""
from __future__ import annotations

import json
from datetime import datetime
from typing import List

import httpx
import pytest

from extractors.idoc import (
    IDocPoller,
    MemoryWatermarkStore,
    MockIDocClient,
)


def _ingest_handler_factory(captured: List[dict]):
    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        captured.append({
            "headers": dict(request.headers),
            "body": body,
        })
        return httpx.Response(
            status_code=200,
            json={
                "status": "ok",
                "received": len(body["records"]),
                "inserted": len(body["records"]),
                "duplicates": 0,
            },
        )
    return handler


def _make_poller(mock_client, *, captured, watermark=None, statuses=None):
    transport = httpx.MockTransport(_ingest_handler_factory(captured))
    http = httpx.Client(transport=transport)
    return IDocPoller(
        connector=mock_client,
        ingest_url="http://test.local/ingest/errors",
        ingest_token="test-token-abc123",
        sap_host="S4D",
        sap_sysnr="00",
        sap_client="100",
        statuses=statuses,
        watermark_store=watermark or MemoryWatermarkStore(),
        http_client=http,
    )


def test_run_once_posts_failure_records_and_advances_watermark(
    mock_client: MockIDocClient,
):
    """Default poll picks up only failure statuses (51/56/63/65/68 inbound,
    02/04/05/26/29 outbound). Fixture 5 is inbound 69 — an informational
    'IDoc was edited' marker, NOT a failure — so it must be skipped by
    default. Four fixtures should make it through."""
    captured: List[dict] = []
    watermark = MemoryWatermarkStore()
    poller = _make_poller(mock_client, captured=captured, watermark=watermark)

    result = poller.run_once()

    assert result.fetched == 4
    assert result.inserted == 4
    assert result.duplicates == 0
    assert result.watermark_before is None
    assert result.watermark_after is not None
    assert isinstance(result.watermark_after, datetime)

    # Exactly one ingest POST with all four failure records
    assert len(captured) == 1
    posted = captured[0]
    assert posted["headers"]["authorization"] == "Bearer test-token-abc123"
    assert posted["headers"]["x-sap-system-id"] == "00"
    assert posted["headers"]["x-sap-client"] == "100"
    assert posted["body"]["host"] == "S4D"
    assert len(posted["body"]["records"]) == 4
    assert all(r["source"] == "IDOC" for r in posted["body"]["records"])
    posted_statuses = {r["error_id"] for r in posted["body"]["records"]}
    # The 69 informational fixture (DOCNUM 123455) must NOT be posted
    assert "0000000000123455" not in posted_statuses

    # Watermark stored
    saved = watermark.get("IDOC", "S4D:00:100")
    assert saved is not None
    assert saved == result.watermark_after


def test_explicit_status_69_pulls_the_informational_fixture(
    mock_client: MockIDocClient,
):
    """Operators can explicitly poll for status 69 to surface edited IDocs."""
    captured: List[dict] = []
    poller = _make_poller(mock_client, captured=captured, statuses=["69"])
    result = poller.run_once()
    assert result.fetched == 1
    assert captured[0]["body"]["records"][0]["error_id"] == "0000000000123455"


def test_second_run_with_advanced_watermark_yields_no_new_records(
    mock_client: MockIDocClient,
):
    captured: List[dict] = []
    watermark = MemoryWatermarkStore()
    poller = _make_poller(mock_client, captured=captured, watermark=watermark)

    poller.run_once()
    captured.clear()

    second = poller.run_once()
    assert second.fetched == 0
    assert second.inserted == 0
    assert len(captured) == 0  # no POST when nothing to ingest


def test_status_filter_only_picks_matching_idocs(mock_client: MockIDocClient):
    captured: List[dict] = []
    poller = _make_poller(mock_client, captured=captured, statuses=["56"])

    result = poller.run_once()
    assert result.fetched == 1
    posted_records = captured[0]["body"]["records"]
    assert len(posted_records) == 1
    assert posted_records[0]["sub_object"] == "MATMAS05"


def test_missing_token_raises_before_any_http_call(mock_client: MockIDocClient):
    poller = IDocPoller(
        connector=mock_client,
        ingest_url="http://test.local/ingest/errors",
        ingest_token="",  # empty
        sap_host="S4D",
        watermark_store=MemoryWatermarkStore(),
    )
    with pytest.raises(RuntimeError, match="ARTILEGENZ_INGEST_TOKEN"):
        poller.run_once()


def test_ingest_4xx_propagates_and_does_not_advance_watermark(
    mock_client: MockIDocClient,
):
    def fail_handler(request):
        return httpx.Response(status_code=401, json={"detail": "invalid bearer token"})

    transport = httpx.MockTransport(fail_handler)
    http = httpx.Client(transport=transport)
    watermark = MemoryWatermarkStore()
    poller = IDocPoller(
        connector=mock_client,
        ingest_url="http://test.local/ingest/errors",
        ingest_token="bad-token",
        sap_host="S4D",
        watermark_store=watermark,
        http_client=http,
    )

    with pytest.raises(httpx.HTTPStatusError):
        poller.run_once()

    # Watermark must NOT advance on failed ingest
    assert watermark.get("IDOC", "S4D:00:100") is None
