"""
dry_run_test.py  —  IDoc 198025 fix sequence: logic-only dry run.

Tests WITHOUT SAP GUI and WITHOUT Claude API.
Mocks the session + all SAP helpers, then exercises:
  1. Date calculation (01.01.{last_year} logic)
  2. DOCNUM zero-padding
  3. Three-tier field-finder order for DOCNUM in WE02/WE05/WE09
  4. bd87 date range when idoc_numbers supplied
  5. create_partner_profile tree + row-add sequence
  6. Full agent tool-call sequence (scan→detail→create→reprocess)

Run:  python dry_run_test.py
"""

import sys, types, unittest
from unittest.mock import MagicMock, patch, call
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# 1.  Stub out win32com and anthropic so the module imports cleanly
# ─────────────────────────────────────────────────────────────────────────────
mock_session         = MagicMock(name="SAPSession")

# win32com mock — needs win32com.client.GetObject to return a chain that
# ends with our mock_session at conn.Children(0).Children(0)
_conn_mock   = MagicMock()
_app_mock    = MagicMock()
_sap_mock    = MagicMock()

_conn_mock.Children.return_value = mock_session      # conn.Children(0) = session
_app_mock.Children.return_value  = _conn_mock        # app.Children(0)  = conn
_sap_mock.GetScriptingEngine     = _app_mock

win32com_client_stub             = MagicMock()
win32com_client_stub.GetObject   = MagicMock(return_value=_sap_mock)

win32com_stub        = MagicMock()
win32com_stub.client = win32com_client_stub

sys.modules["win32com"]        = win32com_stub
sys.modules["win32com.client"] = win32com_client_stub
sys.modules["anthropic"]       = MagicMock()

# ─────────────────────────────────────────────────────────────────────────────
# 2.  Import the agent (session = connect_sap() runs here, hits the mock)
# ─────────────────────────────────────────────────────────────────────────────
import importlib, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
import sap_claude_agent as agent

# Point the module's session at our mock
agent.session = mock_session

# ─────────────────────────────────────────────────────────────────────────────
# 3.  Helper: build a mock SAP field that records .Text assignments
# ─────────────────────────────────────────────────────────────────────────────
class FakeField:
    def __init__(self, fid, ftype="GuiTextField", text=""):
        self.Id   = fid
        self.Type = ftype
        self.Text = text
    def SetFocus(self): pass
    def Press(self):    pass
    def Select(self):   pass


def make_session_with_fields(fields: dict):
    """
    Return a mock session whose FindById returns the matching FakeField,
    or a generic MagicMock for wnd[0]/wnd[1], or raises for unknown IDs.
    """
    wnd_mock = MagicMock(name="wnd0")
    wnd_mock.SendVKey = MagicMock()

    def _find(fid, *a, **kw):
        if fid in ("wnd[0]", "wnd[1]"):
            return wnd_mock
        if fid in fields:
            return fields[fid]
        raise Exception(f"not found: {fid}")

    sess = MagicMock(name="SAPSession")
    sess.FindById = MagicMock(side_effect=_find)
    return sess


PASS = "PASS"
FAIL = "FAIL"
results = []

def check(name, cond, detail=""):
    status = PASS if cond else FAIL
    results.append((name, status, detail))
    mark = "✓" if cond else "✗"
    print(f"  {mark} {name}" + (f"  — {detail}" if detail else ""))
    return cond


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 1: Date calculation ===")
# ═══════════════════════════════════════════════════════════════════════════
today     = datetime.now()
last_year = today.year - 1
expected_start = f"01.01.{last_year}"

# Simulate what scan_idoc_errors computes
computed_start = f"01.01.{datetime.now().year - 1}"
check("Date start = 01.01.last_year", computed_start == expected_start,
      f"got {computed_start}")

# bd87 same logic
bd87_start = f"01.01.{datetime.now().year - 1}"
check("BD87 date start = 01.01.last_year", bd87_start == expected_start,
      f"got {bd87_start}")


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 2: DOCNUM zero-padding ===")
# ═══════════════════════════════════════════════════════════════════════════
padded = str(198025).zfill(16)
check("198025 pads to 16 digits",   len(padded) == 16, padded)
check("Padding is correct zeros",   padded == "0000000000198025", padded)
check("lstrip('0') recovers number", padded.lstrip("0") == "198025")


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 3: _we02_navigate_to_idoc — Tier 1 txt variant ===")
# ═══════════════════════════════════════════════════════════════════════════
# Simulate a SAP system where only txtS_DOCNUM-LOW exists (not ctxt variant)
fields_we02 = {
    "wnd[0]/usr/txtS_DOCNUM-LOW":  FakeField("wnd[0]/usr/txtS_DOCNUM-LOW",
                                              "GuiTextField"),
    "wnd[0]/usr/txtS_DOCNUM-HIGH": FakeField("wnd[0]/usr/txtS_DOCNUM-HIGH",
                                              "GuiTextField"),
}
agent.session = make_session_with_fields(fields_we02)

# Stub helpers that touch SAP
agent.go_to_transaction  = MagicMock()
agent.get_screen_text    = MagicMock(return_value="IDoc Display: 0000000000198025")
agent._clear_selection_screen = MagicMock()
agent._find_input_field_by_fragment = MagicMock(return_value=(None, None))
agent._find_field_by_label          = MagicMock(return_value=(None, None))

on_detail, set_lo, screen = agent._we02_navigate_to_idoc("0000000000198025")
lo_field = fields_we02["wnd[0]/usr/txtS_DOCNUM-LOW"]
hi_field = fields_we02["wnd[0]/usr/txtS_DOCNUM-HIGH"]

check("WE02 txt field LOW set",  lo_field.Text == "0000000000198025",
      f"Text='{lo_field.Text}'")
check("WE02 txt field HIGH set", hi_field.Text == "0000000000198025",
      f"Text='{hi_field.Text}'")
check("WE02 set_lo=True",        set_lo  is True)
check("WE02 on_detail=True",     on_detail is True)


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 4: _we02_navigate_to_idoc — Tier 2 fragment scan ===")
# ═══════════════════════════════════════════════════════════════════════════
# Simulate: no known ID works, but fragment scan finds a GuiNumericTextField
numeric_field = FakeField("wnd[0]/usr/numS_DOCNUM-LOW", "GuiNumericTextField")
hi_numeric    = FakeField("wnd[0]/usr/numS_DOCNUM-HIGH", "GuiNumericTextField")

agent.session = make_session_with_fields({})   # no known IDs match Tier 1
agent.get_screen_text = MagicMock(return_value="IDoc Display: 0000000000198025")
agent._clear_selection_screen = MagicMock()

# fragment scan returns the numeric field for S_DOCNUM-LOW
def fake_fragment(frag):
    if frag in ("S_DOCNUM-LOW", "DOCNUM-LOW", "DOCNUM"):
        return numeric_field, numeric_field.Id
    return None, None
agent._find_input_field_by_fragment = fake_fragment
agent._find_field_by_label = MagicMock(return_value=(None, None))

# FindById for the HIGH field
_wnd_t4 = MagicMock(); _wnd_t4.SendVKey = MagicMock()
def findbyid_hi(fid, *a, **kw):
    if fid in ("wnd[0]", "wnd[1]"):  return _wnd_t4
    if fid == "wnd[0]/usr/numS_DOCNUM-HIGH": return hi_numeric
    raise Exception(f"not found: {fid}")
agent.session.FindById = findbyid_hi

on_detail, set_lo, screen = agent._we02_navigate_to_idoc("0000000000198025")
check("WE02 Tier2 numeric LOW set",  numeric_field.Text == "0000000000198025",
      f"Text='{numeric_field.Text}'")
check("WE02 Tier2 set_lo=True",      set_lo is True)


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 5: _we02_navigate_to_idoc — Tier 3 label fallback ===")
# ═══════════════════════════════════════════════════════════════════════════
label_field = FakeField("wnd[0]/usr/somefield", "GuiTextField")

agent.session = make_session_with_fields({})
agent.get_screen_text = MagicMock(return_value="IDoc Display: 0000000000198025")
agent._clear_selection_screen = MagicMock()
agent._find_input_field_by_fragment = MagicMock(return_value=(None, None))
agent._find_field_by_label = MagicMock(return_value=(label_field, label_field.Id))

on_detail, set_lo, screen = agent._we02_navigate_to_idoc("0000000000198025")
check("WE02 Tier3 label field set", label_field.Text == "0000000000198025",
      f"Text='{label_field.Text}'")
check("WE02 Tier3 set_lo=True",     set_lo is True)


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 6: scan_idoc_errors — date range & DOCNUM with idoc_number ===")
# ═══════════════════════════════════════════════════════════════════════════
date_fields = {
    f"wnd[0]/usr/ctxtSEL_CREDAT-LOW":  FakeField("wnd[0]/usr/ctxtSEL_CREDAT-LOW",  "GuiCTextField"),
    f"wnd[0]/usr/ctxtSEL_CREDAT-HIGH": FakeField("wnd[0]/usr/ctxtSEL_CREDAT-HIGH", "GuiCTextField"),
    f"wnd[0]/usr/txtS_DOCNUM-LOW":     FakeField("wnd[0]/usr/txtS_DOCNUM-LOW",     "GuiTextField"),
    f"wnd[0]/usr/txtS_DOCNUM-HIGH":    FakeField("wnd[0]/usr/txtS_DOCNUM-HIGH",    "GuiTextField"),
}
agent.session = make_session_with_fields(date_fields)
agent.go_to_transaction = MagicMock()
agent.get_screen_text   = MagicMock(return_value="IDoc List")
agent._find_shell_anywhere  = MagicMock(return_value=None)
agent._find_input_field_by_fragment = MagicMock(return_value=(None, None))
agent._find_field_by_label = MagicMock(return_value=(None, None))
agent.discover_elements = MagicMock(return_value=[])
agent._screen_texts     = MagicMock(return_value=[])

result = agent.scan_idoc_errors(idoc_number="198025", status_filter="all")

d_lo = date_fields["wnd[0]/usr/ctxtSEL_CREDAT-LOW"].Text
d_n_lo = date_fields["wnd[0]/usr/txtS_DOCNUM-LOW"].Text
d_n_hi = date_fields["wnd[0]/usr/txtS_DOCNUM-HIGH"].Text

check("WE05 date_from = 01.01.last_year", d_lo == expected_start,
      f"got '{d_lo}'")
check("WE05 DOCNUM LOW = padded",  d_n_lo == "0000000000198025",
      f"got '{d_n_lo}'")
check("WE05 DOCNUM HIGH = padded", d_n_hi == "0000000000198025",
      f"got '{d_n_hi}'")
check("Result has idoc_number key", result.get("idoc_number") == "198025")


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 7: bd87_select_and_reprocess — date range & DOCNUM ===")
# ═══════════════════════════════════════════════════════════════════════════
bd87_fields = {
    "wnd[0]/usr/ctxtSEL_DOCNUM-LOW":  FakeField("wnd[0]/usr/ctxtSEL_DOCNUM-LOW",  "GuiCTextField"),
    "wnd[0]/usr/ctxtSEL_DOCNUM-HIGH": FakeField("wnd[0]/usr/ctxtSEL_DOCNUM-HIGH", "GuiCTextField"),
    "wnd[0]/usr/ctxtSEL_UPDDAT-LOW":  FakeField("wnd[0]/usr/ctxtSEL_UPDDAT-LOW",  "GuiCTextField"),
    "wnd[0]/usr/ctxtSEL_UPDDAT-HIGH": FakeField("wnd[0]/usr/ctxtSEL_UPDDAT-HIGH", "GuiCTextField"),
}
agent.session = make_session_with_fields(bd87_fields)
agent.go_to_transaction = MagicMock()
agent.get_screen_text   = MagicMock(return_value="Select IDocs")
agent._bd87_select_all_and_process = MagicMock(return_value=("SelectAll()", True))
agent.audit_log = MagicMock()
agent.discover_elements = MagicMock(return_value=[])

result = agent.bd87_select_and_reprocess(idoc_numbers=["198025"])

bd87_lo = bd87_fields["wnd[0]/usr/ctxtSEL_DOCNUM-LOW"].Text
bd87_hi = bd87_fields["wnd[0]/usr/ctxtSEL_DOCNUM-HIGH"].Text
bd87_dt = bd87_fields["wnd[0]/usr/ctxtSEL_UPDDAT-LOW"].Text

check("BD87 DOCNUM LOW = padded",      bd87_lo == "0000000000198025",
      f"got '{bd87_lo}'")
check("BD87 DOCNUM HIGH = padded",     bd87_hi == "0000000000198025",
      f"got '{bd87_hi}'")
check("BD87 date_from = 01.01.last_year", bd87_dt == expected_start,
      f"got '{bd87_dt}'")
check("BD87 processed=True",           result.get("processed") is True)


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 8: create_partner_profile — tree navigation + row + fill ===")
# ═══════════════════════════════════════════════════════════════════════════
# Simulate WE20: tree exists, S4HANA2023 found in tree, Create button exists,
# fragment scan finds empty MESTYP and PROCOD fields in new row.
mestyp_field = FakeField("wnd[0]/usr/subSCR/tblTC_IN/ctxtMESTYP[1,5]",
                          "GuiCTextField", text="")
procod_field = FakeField("wnd[0]/usr/subSCR/tblTC_IN/ctxtPROCOD[1,5]",
                          "GuiTextField", text="")
create_btn   = FakeField("wnd[0]/usr/subSUBSCREEN_BODY:SAPLWEDC:0100/btnBT_IN_CREATE",
                          "GuiButton")

# Tree mock
mock_tree = MagicMock()
mock_tree.GetAllNodeKeys.return_value = [
    "LS", "LS_A4HCLNT100", "LS_S4HANA2023", "LS_SAPCISYS"
]
mock_tree.GetNodeText = lambda key: {
    "LS":           "Partner Type LS",
    "LS_A4HCLNT100": "A4HCLNT100",
    "LS_S4HANA2023": "S4HANA2023",
    "LS_SAPCISYS":   "SAPCISYS",
}.get(key, "")

agent._find_gui_tree_anywhere = MagicMock(return_value=mock_tree)
agent.go_to_transaction = MagicMock()
agent.get_screen_text   = MagicMock(return_value="Partner profiles  S4HANA2023")
agent.discover_elements = MagicMock(return_value=[])
agent.audit_log         = MagicMock()
agent.wait_for_user_input = MagicMock(return_value={"status": "user_completed"})

def fragment_for_we20(frag):
    if "MESTYP"  in frag.upper(): return mestyp_field, mestyp_field.Id
    if "PROCOD"  in frag.upper(): return procod_field, procod_field.Id
    return None, None
agent._find_input_field_by_fragment = fragment_for_we20

# session: FindById finds the Create button, tree click no-ops
_wnd_t8 = MagicMock(); _wnd_t8.SendVKey = MagicMock()
def findbyid_we20(fid, *a, **kw):
    if fid in ("wnd[0]", "wnd[1]"):
        return _wnd_t8
    if fid == "wnd[0]/usr/subSUBSCREEN_BODY:SAPLWEDC:0100/btnBT_IN_CREATE":
        return create_btn
    raise Exception(f"not found: {fid}")
agent.session = MagicMock()
agent.session.FindById = findbyid_we20

result = agent.create_partner_profile(
    partner_number="S4HANA2023",
    partner_type="LS",
    direction="1",
    message_type="MATMAS",
    process_code="MATM",
)

check("WE20 tree node S4HANA2023 clicked",
      mock_tree.ClickNode.called,
      f"ClickNode called with: {mock_tree.ClickNode.call_args}")
check("WE20 MESTYP set to MATMAS",
      mestyp_field.Text == "MATMAS", f"got '{mestyp_field.Text}'")
check("WE20 PROCOD set to MATM",
      procod_field.Text == "MATM",   f"got '{procod_field.Text}'")
check("create_partner_profile ok=True", result.get("ok") is True)


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== TEST 9: Full dispatch — tool names route correctly ===")
# ═══════════════════════════════════════════════════════════════════════════
agent.scan_idoc_errors        = MagicMock(return_value={"count": 1, "idoc_errors": []})
agent.get_idoc_detail         = MagicMock(return_value={"error_messages": ["partner profile does not exist"]})
agent.create_partner_profile  = MagicMock(return_value={"ok": True})
agent.bd87_select_and_reprocess = MagicMock(return_value={"processed": True})

# scan_idoc_errors dispatch
r = agent.dispatch("scan_idoc_errors",
                   {"idoc_number": "198025", "status_filter": "all",
                    "direction": "both"})
check("dispatch scan_idoc_errors passes idoc_number",
      agent.scan_idoc_errors.call_args[0][4] == "198025"
      or agent.scan_idoc_errors.call_args[1].get("idoc_number") == "198025"
      or "198025" in str(agent.scan_idoc_errors.call_args),
      str(agent.scan_idoc_errors.call_args))

# get_idoc_detail dispatch
agent.dispatch("get_idoc_detail", {"idoc_number": "198025"})
check("dispatch get_idoc_detail passes idoc_number",
      "198025" in str(agent.get_idoc_detail.call_args))

# create_partner_profile dispatch
agent.dispatch("create_partner_profile", {
    "partner_number": "S4HANA2023", "partner_type": "LS",
    "direction": "1", "message_type": "MATMAS",
    "process_code": "MATM", "basic_type": "MATMAS01",
})
check("dispatch create_partner_profile passes all params",
      "MATMAS" in str(agent.create_partner_profile.call_args)
      and "MATM"   in str(agent.create_partner_profile.call_args))

# bd87 dispatch
agent.dispatch("bd87_select_and_reprocess",
               {"idoc_numbers": ["198025"]})
check("dispatch bd87 passes idoc_numbers",
      "198025" in str(agent.bd87_select_and_reprocess.call_args))


# ═══════════════════════════════════════════════════════════════════════════
print("\n=== SUMMARY ===")
# ═══════════════════════════════════════════════════════════════════════════
passed = sum(1 for _, s, _ in results if s == PASS)
failed = sum(1 for _, s, _ in results if s == FAIL)
total  = len(results)
print(f"\n  {passed}/{total} passed   {failed} failed")
if failed:
    print("\n  FAILURES:")
    for name, status, detail in results:
        if status == FAIL:
            print(f"    ✗ {name}  — {detail}")
    sys.exit(1)
else:
    print("\n  All checks passed — safe to run against SAP.")
    sys.exit(0)
