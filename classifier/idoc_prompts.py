"""IDoc-specific calibration block for the classifier prompts.

The base ``SONNET_SYSTEM_PROMPT`` in ``classifier/prompts.py`` covers ABAP
runtime, BASIS, DB, FI/MM/SD modules, RFC, BTP, etc. It does not call out
IDoc-specific patterns. This module supplies a calibration block that the
caller can prepend to the user message (or append to the system prompt) when
the error originates from the IDoc extractor (``source='IDOC'``).

Why a separate file: keeps the original ``prompts.py`` system prompt intact
and re-cacheable, while still letting source='IDOC' rows get extra guidance
without bloating the cached portion for non-IDoc errors.

Wiring (suggested):

    from classifier.idoc_prompts import idoc_user_preamble, is_idoc_error
    if is_idoc_error(record):
        user_text = idoc_user_preamble(record) + "\\n\\n" + user_text

This puts IDoc context in the *user* turn (uncached, per-error) so the
~10K-token cached system prompt stays untouched and the cache-hit ratio
stays high.
"""
from __future__ import annotations

from typing import Any, Dict, Optional


IDOC_CALIBRATION_BLOCK = """## SAP IDoc-specific diagnostic guidance (this error is an IDoc)

You are diagnosing a SAP IDoc failure. IDocs are persistent message envelopes used for ALE/EDI integration. Diagnose with this in mind.

### IDoc structure
- **EDIDC** (control header): IDoc number (DOCNUM), basic type (IDOCTP, e.g. ORDERS05), message type (MESTYP, e.g. ORDERS), direction (1=outbound, 2=inbound), partner numbers (SNDPRN, RCVPRN), and the *current* status code.
- **EDIDS** (status records): one row per status transition. The **most recent EDIDS row carries the human-readable failure text and message class/number** — this is your primary signal.
- **EDID4** (data segments): payload structures (E1EDK01 header, E1EDP01 line, etc.). Reference these only when the status text points to a specific segment/field.

### Status-code semantics (memorize these)
**Inbound failures:**
- 51 = Application document not posted. The IDoc structure is fine; the receiving SAP application (e.g. order create, GR posting, IDoc-driven master-data update) rejected it. Root cause is almost always **data, config, or authorization** in the receiving module — not the IDoc itself.
- 56 = IDoc with errors added. Failure occurred during inbound processing (function module called by partner profile rejected the IDoc). Often partner-profile config (WE20), missing process code (WE41/WE42), or function-module-level data validation.
- 63 = Error passing IDoc to application — indicates the IDoc reached the inbound function but the function module failed to call the application correctly (often custom Z-code).
- 65 = Error in ALE service — partner profile (WE20) misconfiguration, missing inbound process code, or BD64 model view issue.
- 68 = Error — no further processing; manual reset to 51 with BD87 if the underlying issue is fixed.
- 69 = IDoc was edited (by WE19/WE02) — informational marker, not a fault per se.

**Outbound failures:**
- 02 = Error passing data to port. Port (WE21) misconfigured: file path / RFC dest / queue not reachable.
- 04 = Error within control information. Sender/receiver partner not in WE20, or partner profile missing outbound parameters for the message type.
- 05 = Error during translation (subsystem / EDI converter).
- 26 = Error during syntax check. The outbound IDoc violates the basic-type structure — usually a custom population error in the outbound function module.
- 29 = Error in ALE service (outbound) — model view (BD64), filter (BD59), or partner profile issue.
- 30 = IDoc ready for dispatch (waiting). Often stuck because RSEOUT00 isn't scheduled or queue is paused.
- 32 = IDoc was edited (manual repair marker).

### Diagnostic priorities for IDocs
1. **Read the latest EDIDS status text first.** It carries the application's own error message (often an MSGID/MSGNO pair like ``F5 263``). Resolve those via SE91 mentally — if you recognize the message class/number, name it.
2. **Inbound 51 root causes** are dominated by: missing customizing in the target module (e.g. material master not extended to plant; vendor not in company code; pricing condition record missing), authorization on the inbound user (often ALERMAS / WF-BATCH), data integrity in the segments (currency/UoM conversion, decimal places), or version mismatch (sender uses ORDERS05, receiver expects ORDERS04 fields).
3. **Inbound 56 root causes** are dominated by: WE20 partner profile inbound parameters missing/incomplete, WE41/WE42 process code wrong, posting logic exception in the inbound function module (often a Z* function for custom message types).
4. **Outbound 02 root causes**: WE21 port definition (file directory permission, RFC destination not reachable, queue paused), or an OS-level write failure on the application server filesystem.
5. **Outbound 26 root causes**: customer outbound function module populating the IDoc didn't respect the basic-type structure (wrong segment qualifier, mandatory field empty, wrong segment hierarchy).
6. **Don't recommend "delete and reprocess"** unless the root cause is genuinely transient. The IDoc carries business data — losing it has real cost. Prefer WE19 (edit) → WE02 (reprocess) workflows.

### Remediation step style for IDocs
Always reference real transactions in the remediation steps:
- ``WE02`` — IDoc list display
- ``WE05`` — IDoc display
- ``WE19`` — Test tool / IDoc editor (edit failed inbound IDoc and reprocess)
- ``WE20`` — Partner profile maintenance
- ``WE21`` — Port definition
- ``BD87`` — Status monitor for ALE messages (re-trigger inbound)
- ``BD64`` — Distribution model view
- ``SM58`` — tRFC monitor (look here for outbound 03/30 stuck IDocs)
- ``SLG1`` — Application log (often where the inbound application logs the real error)

### Severity calibration for IDocs
- **CRITICAL**: integration is fully broken (port down, all IDocs of a type failing) and customer-facing.
- **HIGH**: inbound 51/56 on master-data or order flows where the data is irrecoverable without manual rework.
- **MEDIUM**: outbound 26/30 (data isn't lost, just stuck), or 51 on a low-volume custom message type.
- **LOW**: 32/69 (manual edit markers — informational).

### What to put in `citations`
Cite both retrieved corpus items (Stack Overflow IDoc threads, SAP Notes) **and** any retrieved raw_errors that share the same MESTYP+IDOCTP+STATUS combination — these prove the issue is recurring vs. one-off.
"""


def is_idoc_error(record: Dict[str, Any]) -> bool:
    """Return True iff this raw_errors record came from the IDoc extractor."""
    src = record.get("source") or record.get("SOURCE") or ""
    return str(src).upper() == "IDOC"


def idoc_user_preamble(record: Dict[str, Any]) -> str:
    """Build a short preamble injecting the IDoc calibration block + a parsed
    summary of the IDoc-specific fields from the ``raw`` JSONB.

    This goes into the *user* turn (uncached, ~2-3K tokens) so the cached
    system prompt stays unchanged for non-IDoc traffic.
    """
    raw = record.get("raw") or record.get("RAW") or {}
    if isinstance(raw, str):
        try:
            import json
            raw = json.loads(raw)
        except (TypeError, ValueError):
            raw = {}

    parts = [IDOC_CALIBRATION_BLOCK, "", "## Parsed IDoc fields for this error"]
    parts.append(f"- DOCNUM:    {raw.get('docnum') or record.get('error_id') or '?'}")
    parts.append(f"- MESTYP:    {raw.get('mestyp') or '?'}")
    parts.append(f"- IDOCTP:    {raw.get('idoctp') or '?'}")
    parts.append(f"- Direction: {raw.get('direction') or '?'}")
    parts.append(f"- Status:    {raw.get('current_status') or record.get('error_id') or '?'}")
    parts.append(f"- Sender:    {raw.get('sender_partner') or '?'}")
    parts.append(f"- Receiver:  {raw.get('receiver_partner') or '?'}")
    parts.append(f"- Client:    {raw.get('client') or '?'}")

    statuses = raw.get("statuses") or []
    if statuses:
        parts.append("")
        parts.append("### Latest 3 EDIDS rows (most recent last)")
        latest = statuses[-3:] if len(statuses) > 3 else statuses
        for s in latest:
            parts.append(
                f"- [{s.get('log_at', '?')}] status={s.get('status', '?')} "
                f"text={(s.get('status_text') or '')[:200]}"
            )
            mid = s.get("msg_id")
            mno = s.get("msg_no")
            if mid and mno:
                vs = " ".join(filter(None, [s.get("msg_v1"), s.get("msg_v2"),
                                            s.get("msg_v3"), s.get("msg_v4")]))
                parts.append(f"  msg {mid}/{mno} {vs}".rstrip())
            if s.get("seg_num") or s.get("seg_fld"):
                parts.append(
                    f"  segment={s.get('seg_num') or '?'}  field={s.get('seg_fld') or '?'}"
                )
    return "\n".join(parts)


def maybe_augment_user_text(record: Dict[str, Any], user_text: str) -> str:
    """Convenience wrapper: if ``record`` is an IDoc, prepend the preamble.

    Otherwise return ``user_text`` unchanged. Use this in
    ``classifier/prompts.build_sonnet_messages`` to add IDoc guidance without
    affecting non-IDoc errors.
    """
    if not is_idoc_error(record):
        return user_text
    return idoc_user_preamble(record) + "\n\n" + user_text
