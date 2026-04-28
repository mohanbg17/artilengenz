"""IDoc → AbapErrorRecord transformer."""
from __future__ import annotations

import json

from extractors.idoc import IDoc, MockIDocClient
from extractors.idoc.transformer import idoc_to_ingest_record


def test_orders05_status51_record_shape(mock_client: MockIDocClient):
    idoc = next(
        d for d in mock_client.list_failed_idocs(max_results=999)
        if d.docnum == "0000000000123451"
    )
    rec = idoc_to_ingest_record(idoc)

    assert rec["source"] == "IDOC"
    assert rec["error_id"] == "0000000000123451"
    assert rec["severity"] == "HIGH"
    assert rec["transaction"] == "WE19"
    assert rec["object"] == "ORDERS"
    assert rec["sub_object"] == "ORDERS05"
    assert "Customer" in rec["short_text"]
    assert len(rec["short_text"]) <= 2000


def test_outbound_uses_bd87_transaction():
    """Constructed outbound IDoc — fixtures don't include outbound failures."""
    from datetime import datetime as dt

    from extractors.idoc import OUTBOUND, IDoc, IDocStatus

    outbound = IDoc(
        docnum="0000000000999999",
        mestyp="ORDRSP",
        idoctp="ORDERS05",
        direction=OUTBOUND,
        status="02",
        created_at=dt(2025, 4, 27, 9, 0),
        sender_partner="S4D_100",
        receiver_partner="PARTNER_X",
        client="100",
        statuses=[
            IDocStatus(
                counter=1,
                status="02",
                status_text="Error passing data to port",
                log_at=dt(2025, 4, 27, 9, 1),
            )
        ],
    )
    rec = idoc_to_ingest_record(outbound)
    assert rec["transaction"] == "BD87"
    assert rec["severity"] == "HIGH"


def test_long_text_includes_status_history_and_segments(mock_client: MockIDocClient):
    idoc = next(
        d for d in mock_client.list_failed_idocs(max_results=999)
        if d.docnum == "0000000000123451"
    )
    rec = idoc_to_ingest_record(idoc)
    long_text = rec["long_text"]
    # All three statuses should appear
    assert "status=50" in long_text
    assert "status=64" in long_text
    assert "status=51" in long_text
    # Status text should appear
    assert "Customer" in long_text
    # Segment names should appear
    assert "E1EDK01" in long_text
    assert "E1EDP01" in long_text
    # Message id/no rendered
    assert "V1/319" in long_text


def test_raw_payload_is_valid_json_with_full_idoc_metadata(mock_client: MockIDocClient):
    idoc = next(
        d for d in mock_client.list_failed_idocs(max_results=999)
        if d.docnum == "0000000000123452"
    )
    rec = idoc_to_ingest_record(idoc)
    raw = json.loads(rec["raw"])
    assert raw["docnum"] == "0000000000123452"
    assert raw["mestyp"] == "MATMAS"
    assert raw["idoctp"] == "MATMAS05"
    assert raw["direction"] == "INBOUND"
    assert len(raw["statuses"]) == 2
    assert len(raw["segments"]) == 3
    # Status timestamps must be ISO strings (json-serializable)
    for s in raw["statuses"]:
        assert isinstance(s["log_at"], str)


def test_occurred_at_is_abap_timestamp_format(mock_client: MockIDocClient):
    idoc = next(
        d for d in mock_client.list_failed_idocs(max_results=999)
        if d.docnum == "0000000000123451"
    )
    rec = idoc_to_ingest_record(idoc)
    # ABAP push format from api.ingest._parse_abap_timestamp: "%Y%m%dT%H%M%S"
    assert len(rec["occurred_at"]) == 15
    assert rec["occurred_at"][8] == "T"


def test_transformer_roundtrip_through_ingest_pydantic_model(mock_client: MockIDocClient):
    """Each transformed record must be accepted by api.ingest.AbapErrorRecord."""
    from api.ingest import AbapErrorRecord

    for idoc in mock_client.list_failed_idocs(max_results=999):
        rec = idoc_to_ingest_record(idoc)
        # Pydantic v2: model_validate raises on invalid shape
        validated = AbapErrorRecord.model_validate(rec)
        assert validated.source == "IDOC"
        assert validated.error_id == idoc.docnum
