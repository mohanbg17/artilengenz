"""Map SAP IDoc status codes to platform severity buckets.

Reference: SAP IDoc status codes (transaction WE05 / table TEDS1):

Inbound (DIRECT='2'):
  50  IDoc added                   (transient — not failed)
  51  Application document not posted   (failure to apply)
  52  Application document posted        (success)
  56  IDoc with errors added             (failure during inbound)
  60  Error during syntax check          (structural error)
  61  Processing despite syntax error    (processed-with-warnings)
  62  IDoc passed to application         (transient)
  63  Error passing IDoc to application  (interface failure)
  64  IDoc ready to be transferred       (transient — pending)
  65  Error in ALE service               (config / partner profile)
  66  IDoc is waiting for predecessor    (transient)
  68  Error — no further processing      (terminal failure)
  69  IDoc was edited                    (manual repair marker)

Outbound (DIRECT='1'):
  01  IDoc created
  02  Error passing data to port         (failure)
  03  Data passed to port OK             (success)
  04  Error within control information   (failure)
  05  Error during translation           (failure)
  18  Triggering OK
  26  Error during syntax check (out)    (failure)
  29  Error in ALE service (out)         (failure)
  30  IDoc ready for dispatch (ALE)      (transient — pending)
  32  IDoc was edited (out)              (manual repair marker)
  37  IDoc added (out, partner profile)
  41  Application document created in receiver
"""
from __future__ import annotations

# Inbound failure-or-attention statuses → severity
_INBOUND: dict[str, str] = {
    "51": "HIGH",
    "56": "HIGH",
    "60": "MEDIUM",
    "61": "LOW",
    "63": "HIGH",
    "65": "HIGH",
    "68": "MEDIUM",
    "69": "LOW",
}

# Outbound failure statuses → severity
_OUTBOUND: dict[str, str] = {
    "02": "HIGH",
    "04": "HIGH",
    "05": "HIGH",
    "26": "MEDIUM",
    "29": "HIGH",
    "30": "LOW",
    "32": "LOW",
}

# Statuses we treat as "open / not failed" (used by callers to filter)
HEALTHY_STATUSES = frozenset({"01", "03", "18", "41", "50", "52", "53", "62", "64", "66"})

# Statuses callers should poll for as "needs attention"
DEFAULT_FAILED_STATUSES_INBOUND = ("51", "56", "63", "65", "68")
DEFAULT_FAILED_STATUSES_OUTBOUND = ("02", "04", "05", "26", "29")
DEFAULT_FAILED_STATUSES = (
    DEFAULT_FAILED_STATUSES_INBOUND + DEFAULT_FAILED_STATUSES_OUTBOUND
)


def severity_for(status: str, direction: str) -> str:
    """Return CRITICAL / HIGH / MEDIUM / LOW for the given status + direction.

    Unknown statuses default to MEDIUM. Non-failure statuses still resolve
    rather than raising, so callers can render any IDoc.
    """
    if status in HEALTHY_STATUSES:
        return "LOW"
    if direction == "INBOUND":
        return _INBOUND.get(status, "MEDIUM")
    if direction == "OUTBOUND":
        return _OUTBOUND.get(status, "MEDIUM")
    return "MEDIUM"


def is_failure(status: str, direction: str) -> bool:
    """True if this status represents a failure mode (vs. success/transient)."""
    if status in HEALTHY_STATUSES:
        return False
    if direction == "INBOUND":
        return status in _INBOUND
    if direction == "OUTBOUND":
        return status in _OUTBOUND
    return False
