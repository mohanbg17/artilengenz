"""Severity mapping for IDoc statuses."""
from __future__ import annotations

import pytest

from extractors.idoc.severity import (
    DEFAULT_FAILED_STATUSES,
    DEFAULT_FAILED_STATUSES_INBOUND,
    DEFAULT_FAILED_STATUSES_OUTBOUND,
    HEALTHY_STATUSES,
    is_failure,
    severity_for,
)


@pytest.mark.parametrize(
    "status,direction,expected",
    [
        ("51", "INBOUND", "HIGH"),
        ("56", "INBOUND", "HIGH"),
        ("63", "INBOUND", "HIGH"),
        ("65", "INBOUND", "HIGH"),
        ("68", "INBOUND", "MEDIUM"),
        ("60", "INBOUND", "MEDIUM"),
        ("61", "INBOUND", "LOW"),
        ("69", "INBOUND", "LOW"),
        ("02", "OUTBOUND", "HIGH"),
        ("04", "OUTBOUND", "HIGH"),
        ("05", "OUTBOUND", "HIGH"),
        ("26", "OUTBOUND", "MEDIUM"),
        ("29", "OUTBOUND", "HIGH"),
        ("30", "OUTBOUND", "LOW"),
        ("32", "OUTBOUND", "LOW"),
    ],
)
def test_severity_known_failure_statuses(status, direction, expected):
    assert severity_for(status, direction) == expected


@pytest.mark.parametrize("status", sorted(HEALTHY_STATUSES))
def test_healthy_statuses_are_low_severity_in_either_direction(status):
    assert severity_for(status, "INBOUND") == "LOW"
    assert severity_for(status, "OUTBOUND") == "LOW"


def test_unknown_status_defaults_to_medium():
    assert severity_for("ZZ", "INBOUND") == "MEDIUM"
    assert severity_for("ZZ", "OUTBOUND") == "MEDIUM"
    assert severity_for("ZZ", "UNKNOWN_DIRECTION") == "MEDIUM"


def test_is_failure_excludes_healthy_statuses():
    for s in HEALTHY_STATUSES:
        assert not is_failure(s, "INBOUND")
        assert not is_failure(s, "OUTBOUND")


@pytest.mark.parametrize("status", DEFAULT_FAILED_STATUSES_INBOUND)
def test_default_inbound_statuses_are_failures(status):
    assert is_failure(status, "INBOUND")


@pytest.mark.parametrize("status", DEFAULT_FAILED_STATUSES_OUTBOUND)
def test_default_outbound_statuses_are_failures(status):
    assert is_failure(status, "OUTBOUND")


def test_default_failed_set_is_union_of_inbound_and_outbound():
    assert set(DEFAULT_FAILED_STATUSES) == (
        set(DEFAULT_FAILED_STATUSES_INBOUND) | set(DEFAULT_FAILED_STATUSES_OUTBOUND)
    )
