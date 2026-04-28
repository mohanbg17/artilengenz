"""Sanity tests for the IDoc-specific prompt calibration."""
from __future__ import annotations

import json

from classifier.idoc_prompts import (
    IDOC_CALIBRATION_BLOCK,
    idoc_user_preamble,
    is_idoc_error,
    maybe_augment_user_text,
)


def test_is_idoc_error_recognizes_source_idoc():
    assert is_idoc_error({"source": "IDOC"})
    assert is_idoc_error({"source": "idoc"})
    assert is_idoc_error({"SOURCE": "IDOC"})  # uppercase column name from DB row
    assert not is_idoc_error({"source": "ST22"})
    assert not is_idoc_error({})


def test_calibration_block_mentions_key_idoc_concepts():
    blob = IDOC_CALIBRATION_BLOCK
    for term in [
        "EDIDC", "EDIDS", "EDID4",
        "WE19", "WE20", "WE21", "BD87", "SLG1",
        "MESTYP", "IDOCTP",
        "Inbound failures", "Outbound failures",
        "51", "56", "02", "26",
    ]:
        assert term in blob, f"expected calibration block to mention '{term}'"


def test_preamble_extracts_fields_from_raw_jsonb_dict():
    record = {
        "source": "IDOC",
        "error_id": "0000000000123451",
        "raw": {
            "docnum": "0000000000123451",
            "mestyp": "ORDERS",
            "idoctp": "ORDERS05",
            "direction": "INBOUND",
            "current_status": "51",
            "sender_partner": "PARTNER_A",
            "receiver_partner": "S4D_100",
            "client": "100",
            "statuses": [
                {
                    "log_at": "2025-04-27T07:01:00",
                    "status": "51",
                    "status_text": "Customer not assigned to sales area",
                    "msg_id": "V1",
                    "msg_no": "319",
                    "msg_v1": "0001000123",
                    "seg_num": "000003",
                    "seg_fld": "PARTN",
                },
            ],
        },
    }
    preamble = idoc_user_preamble(record)

    assert "IDoc-specific diagnostic guidance" in preamble
    assert "DOCNUM:    0000000000123451" in preamble
    assert "MESTYP:    ORDERS" in preamble
    assert "IDOCTP:    ORDERS05" in preamble
    assert "Direction: INBOUND" in preamble
    assert "Customer not assigned to sales area" in preamble
    assert "msg V1/319" in preamble
    assert "segment=000003  field=PARTN" in preamble


def test_preamble_handles_raw_as_json_string():
    raw_payload = {
        "docnum": "0000000000123452",
        "mestyp": "MATMAS",
        "idoctp": "MATMAS05",
        "direction": "INBOUND",
        "current_status": "56",
        "statuses": [],
    }
    record = {"source": "IDOC", "raw": json.dumps(raw_payload)}
    preamble = idoc_user_preamble(record)
    assert "MESTYP:    MATMAS" in preamble


def test_preamble_with_no_raw_falls_back_to_top_level_fields():
    record = {"source": "IDOC", "error_id": "DOC-X"}
    preamble = idoc_user_preamble(record)
    assert "DOCNUM:    DOC-X" in preamble


def test_maybe_augment_passes_through_for_non_idoc():
    record = {"source": "ST22"}
    out = maybe_augment_user_text(record, "original user text")
    assert out == "original user text"


def test_maybe_augment_prepends_for_idoc():
    record = {"source": "IDOC", "raw": {"docnum": "1", "mestyp": "ORDERS"}}
    out = maybe_augment_user_text(record, "original user text")
    assert out.endswith("original user text")
    assert "IDoc-specific diagnostic guidance" in out
