"""Normalized IDoc data structures used between connector and transformer.

Decouples the OData/CDS/mock wire format from the transformer so adding new
extraction backends later (CDS view, Event Mesh) only requires a new
connector implementation, not changes downstream.

EDIDC = IDoc control header
EDIDS = IDoc status record (one IDoc has many)
EDID4 = IDoc data segment (one IDoc has many)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


INBOUND = "INBOUND"
OUTBOUND = "OUTBOUND"


@dataclass(frozen=True)
class IDocStatus:
    counter: int
    status: str
    status_text: str
    log_at: datetime
    msg_id: Optional[str] = None
    msg_no: Optional[str] = None
    msg_v1: Optional[str] = None
    msg_v2: Optional[str] = None
    msg_v3: Optional[str] = None
    msg_v4: Optional[str] = None
    seg_num: Optional[str] = None
    seg_fld: Optional[str] = None


@dataclass(frozen=True)
class IDocSegment:
    counter: int
    seg_name: str
    data: str


@dataclass
class IDoc:
    docnum: str
    mestyp: str
    idoctp: str
    direction: str
    status: str
    created_at: datetime
    sender_partner: str
    receiver_partner: str
    client: str
    statuses: List[IDocStatus] = field(default_factory=list)
    segments: List[IDocSegment] = field(default_factory=list)

    @property
    def latest_status(self) -> Optional[IDocStatus]:
        if not self.statuses:
            return None
        return max(self.statuses, key=lambda s: (s.log_at, s.counter))
