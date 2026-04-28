"""Transform a normalized ``IDoc`` into a record matching the platform's
``AbapErrorRecord`` ingest contract (defined in ``api/ingest.py``).

The platform's ``raw.raw_errors`` schema is source-agnostic: ``source`` is a
free-string discriminator and the ``raw`` JSONB column carries
source-specific detail. So an IDoc fits as ``source='IDOC'`` with the IDoc
header and status/segment data preserved verbatim under ``raw``.

Field mapping rationale:
  source       'IDOC'                          (source discriminator)
  error_id     IDoc DOCNUM                     (unique within a system)
  occurred_at  EDIDC creation timestamp        (ABAP timestamp string)
  severity     derived from STATUS+direction   (see severity.py)
  short_text   most-recent EDIDS status text   (capped at 2000)
  long_text    full status history + segments  (helps the classifier diagnose)
  transaction  WE19 (inbound) / BD87 (outbound)  (operator workflow t-code)
  object       MESTYP (message type)
  sub_object   IDOCTP (basic IDoc type)
  raw          JSON blob with full EDIDC + EDIDS list + EDID4 list
"""
from __future__ import annotations

import json
from dataclasses import asdict
from typing import Any, Dict

from .models import INBOUND, IDoc
from .severity import severity_for


_SHORT_TEXT_MAX = 2000


def idoc_to_ingest_record(idoc: IDoc) -> Dict[str, Any]:
    """Return a dict matching ``api.ingest.AbapErrorRecord`` for one IDoc."""
    latest = idoc.latest_status
    short_text = (
        latest.status_text if latest and latest.status_text
        else f"IDoc {idoc.docnum} {idoc.idoctp}/{idoc.mestyp} status {idoc.status}"
    )
    short_text = short_text[:_SHORT_TEXT_MAX]

    return {
        "source": "IDOC",
        "error_id": idoc.docnum,
        "occurred_at": idoc.created_at.strftime("%Y%m%dT%H%M%S"),
        "severity": severity_for(idoc.status, idoc.direction),
        "short_text": short_text,
        "long_text": _build_long_text(idoc),
        "transaction": "WE19" if idoc.direction == INBOUND else "BD87",
        "object": idoc.mestyp or None,
        "sub_object": idoc.idoctp or None,
        "raw": json.dumps(_raw_payload(idoc), default=str, ensure_ascii=False),
    }


def _build_long_text(idoc: IDoc) -> str:
    lines = [
        f"IDoc {idoc.docnum}  basic-type={idoc.idoctp}  message-type={idoc.mestyp}",
        f"Direction={idoc.direction}  current-status={idoc.status}  client={idoc.client or '?'}",
        f"Sender partner={idoc.sender_partner or '?'}    Receiver partner={idoc.receiver_partner or '?'}",
    ]

    if idoc.statuses:
        lines.append("")
        lines.append("Status history (chronological):")
        for s in sorted(idoc.statuses, key=lambda x: (x.log_at, x.counter)):
            ts = s.log_at.strftime("%Y-%m-%d %H:%M:%S")
            lines.append(f"  [{ts}] status={s.status}  {s.status_text}")
            if s.seg_num or s.seg_fld:
                lines.append(
                    f"      segment={s.seg_num or '?'}  field={s.seg_fld or '?'}"
                )
            mv = " ".join(filter(None, [s.msg_v1, s.msg_v2, s.msg_v3, s.msg_v4]))
            if s.msg_id and s.msg_no:
                line = f"      msg {s.msg_id}/{s.msg_no}"
                if mv:
                    line += f"  {mv}"
                lines.append(line)

    if idoc.segments:
        lines.append("")
        lines.append("Segment data (truncated to 200 chars per segment):")
        for seg in sorted(idoc.segments, key=lambda x: x.counter):
            data = (seg.data or "")[:200]
            lines.append(f"  [{seg.seg_name}] {data}")

    return "\n".join(lines)


def _raw_payload(idoc: IDoc) -> Dict[str, Any]:
    return {
        "docnum": idoc.docnum,
        "mestyp": idoc.mestyp,
        "idoctp": idoc.idoctp,
        "direction": idoc.direction,
        "current_status": idoc.status,
        "client": idoc.client,
        "sender_partner": idoc.sender_partner,
        "receiver_partner": idoc.receiver_partner,
        "created_at": idoc.created_at.isoformat(),
        "statuses": [_status_dict(s) for s in idoc.statuses],
        "segments": [asdict(seg) for seg in idoc.segments],
    }


def _status_dict(status) -> Dict[str, Any]:
    d = asdict(status)
    d["log_at"] = status.log_at.isoformat()
    return d
