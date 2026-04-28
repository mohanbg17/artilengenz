"""MockIDocClient parses fixture JSON exactly like the production OData parser."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from extractors.idoc import INBOUND, OUTBOUND, IDoc, MockIDocClient


def test_loads_all_five_fixtures(all_fixture_idocs: list[IDoc]):
    assert len(all_fixture_idocs) == 5
    docnums = {d.docnum for d in all_fixture_idocs}
    assert docnums == {
        "0000000000123451",
        "0000000000123452",
        "0000000000123453",
        "0000000000123454",
        "0000000000123455",
    }


def test_orders05_status51_inbound_parsed(all_fixture_idocs: list[IDoc]):
    idoc = next(d for d in all_fixture_idocs if d.docnum == "0000000000123451")
    assert idoc.idoctp == "ORDERS05"
    assert idoc.mestyp == "ORDERS"
    assert idoc.direction == INBOUND
    assert idoc.status == "51"
    assert idoc.client == "100"
    assert idoc.sender_partner == "PARTNER_CUST_A"
    assert len(idoc.statuses) == 3
    assert len(idoc.segments) == 3

    latest = idoc.latest_status
    assert latest is not None
    assert latest.status == "51"
    assert "Customer" in latest.status_text
    assert latest.msg_id == "V1"
    assert latest.msg_no == "319"
    assert latest.seg_fld == "PARTN"


def test_desadv01_status69_inbound_parsed(all_fixture_idocs: list[IDoc]):
    """Status 69 ('IDoc was edited') is inbound-only in standard SAP."""
    idoc = next(d for d in all_fixture_idocs if d.docnum == "0000000000123455")
    assert idoc.idoctp == "DESADV01"
    assert idoc.direction == INBOUND
    assert idoc.status == "69"
    # Latest status should be the most recent log_at — the 69 row, not the 26 row
    latest = idoc.latest_status
    assert latest is not None
    assert latest.status == "69"


def test_filtering_by_status(mock_client: MockIDocClient):
    only_51 = mock_client.list_failed_idocs(statuses=["51"], max_results=10)
    assert len(only_51) == 3
    assert all(d.status == "51" for d in only_51)


def test_filtering_by_since(mock_client: MockIDocClient):
    # Fixtures are at 2025-04-27 08:00, 09:00, 10:00, 11:00, 12:00 UTC.
    # A cutoff between #1 and #2 should drop #1 and keep the other four.
    cutoff = datetime(2025, 4, 27, 8, 30, 0)
    after = mock_client.list_failed_idocs(since=cutoff, max_results=10)
    assert all(d.created_at > cutoff for d in after)
    assert len(after) == 4

    # A cutoff after the latest fixture must drop them all.
    later_than_all = datetime(2025, 4, 28, 0, 0, 0)
    after_all = mock_client.list_failed_idocs(since=later_than_all, max_results=10)
    assert after_all == []


def test_max_results_caps_output(mock_client: MockIDocClient):
    capped = mock_client.list_failed_idocs(max_results=2)
    assert len(capped) == 2


def test_results_are_sorted_by_created_at(all_fixture_idocs: list[IDoc]):
    timestamps = [d.created_at for d in all_fixture_idocs]
    assert timestamps == sorted(timestamps)


def test_missing_fixtures_dir_raises(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        MockIDocClient(fixtures_dir=tmp_path / "does_not_exist")
