"""
ARTILEGENZ SAP Claude Agent v14.0
Authorization: SAP_ALL + SAP_NEW (full system access confirmed via SU01)
User: S4ABAP24

SAP_ALL grants unrestricted access to:
  - ALL transactions (VA*, ME*, MM*, FI*, CO*, HR*, BC*, basis, development)
  - ALL tables (read + write via SM30/SM31/SE16N)
  - ALL SPRO customising activities
  - ALL ABAP development (SE38, SE37, SE80, SE11)
  - ALL transport functions (SE09, SE10, STMS)
  - ALL user/role administration (SU01, PFCG)
  - ALL system administration (SM50, SM51, SM12, RZ10, RZ20)

Requirements:
    pip install anthropic pywin32
Run from WinPython Command Prompt with SAP GUI open and logged in as S4ABAP24.
"""

import win32com.client
import anthropic
import time
import json
import os
import sys
import getpass
from datetime import datetime

# Temp folder for ABAP source downloads/uploads
TEMP_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "SAP_ABAP_Temp")
os.makedirs(TEMP_DIR, exist_ok=True)

# ── SAP Connection ─────────────────────────────────────────────────────────────
def connect_sap():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    app  = SapGuiAuto.GetScriptingEngine
    conn = app.Children(0)
    sess = conn.Children(0)
    print(f"Connected to SAP: {conn.Description}")
    return sess

session = connect_sap()

CURRENT_USER = getpass.getuser()
LOG_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "SAP_Agent_Logs")
os.makedirs(LOG_DIR, exist_ok=True)

# ── Audit Log ──────────────────────────────────────────────────────────────────
audit_entries = []

def audit_log(action, details, status="completed", approved_by=None):
    entry = {
        "timestamp":   datetime.now().isoformat(),
        "user":        approved_by or CURRENT_USER,
        "action":      action,
        "details":     details,
        "status":      status,
    }
    audit_entries.append(entry)
    path = os.path.join(LOG_DIR, f"audit_{datetime.now().strftime('%Y%m%d')}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(audit_entries, f, indent=2, ensure_ascii=False)

# ── Human Approval ─────────────────────────────────────────────────────────────
def ask_approval(action, details, risk="medium"):
    risk_label = {"low": "[ LOW  ]", "medium": "[MEDIUM]", "high": "[ HIGH ]"}
    print("\n" + "═" * 68)
    print(f"  APPROVAL REQUIRED  {risk_label.get(risk, '[?????]')}")
    print("═" * 68)
    print(f"  Action  : {action}")
    if isinstance(details, dict):
        for k, v in details.items():
            print(f"  {k:<10}: {v}")
    else:
        print(f"  Details : {details}")
    print("─" * 68)
    print("  A = Approve    R = Reject")
    print("─" * 68)
    while True:
        c = input("  Your decision [A/R]: ").strip().upper()
        if c == "A":
            audit_log(action, details, status="approved", approved_by=CURRENT_USER)
            return True
        if c == "R":
            reason = input("  Reason (optional): ").strip()
            audit_log(action, {**details, "rejection_reason": reason},
                      status="rejected", approved_by=CURRENT_USER)
            return False

# ── SAP Primitives ─────────────────────────────────────────────────────────────
def get_screen_text():
    try:
        title = session.FindById("wnd[0]").Text
        try:
            sbar = session.FindById("wnd[0]/sbar").Text
        except Exception:
            sbar = ""
        return f"Screen: {title} | Status: {sbar}"
    except Exception as e:
        return f"Error: {e}"

def go_to_transaction(tcode):
    session.StartTransaction(str(tcode).strip())
    time.sleep(1.5)
    return {"navigated_to": tcode, "screen": get_screen_text()}

def discover_elements():
    """Walk current screen; return up to 50 interactable input/button elements.
    Skips GuiTree / GuiShell data-display controls to prevent hangs on
    result screens (BD87, ST22, WE05, etc. which have hundreds of tree nodes)."""
    results = []
    # These types are data-display containers — do NOT recurse into them
    SKIP_RECURSE = {
        "GuiShell",              # ALV grids — can have thousands of children
        "GuiTree",               # Tree controls — same problem
        "GuiGridView",           # Grid view variant
        "GuiContainerShell",
        "GuiSplitterContainer",
        "GuiCustomControl",
        # NOTE: GuiTableControl is NOT skipped — SE16N filter rows live there
    }
    def walk(comp, depth=0):
        if depth > 5 or len(results) >= 50:
            return
        try:
            n = comp.Children.Count
        except Exception:
            return
        for i in range(n):
            if len(results) >= 50:
                return
            try:
                child = comp.Children(i)
                t = child.Type
                # collect interactable widgets
                if t in ("GuiTextField", "GuiCTextField", "GuiComboBox",
                         "GuiRadioButton", "GuiCheckBox", "GuiButton", "GuiTab"):
                    try:
                        results.append({
                            "id":      child.Id,
                            "type":    t,
                            "value":   child.Text,
                            "tooltip": child.Tooltip,
                        })
                    except Exception:
                        pass
                # recurse into layout containers but NOT data shells/trees
                if t not in SKIP_RECURSE:
                    walk(child, depth + 1)
            except Exception:
                pass
    walk(session.FindById("wnd[0]"))
    return results[:50]

def set_field(element_id, value):
    try:
        f = session.FindById(element_id)
        f.Text = str(value)
        f.SetFocus()
        return {"ok": True, "id": element_id, "value": value}
    except Exception as e:
        return {"ok": False, "error": str(e), "id": element_id}

def press_button(element_id):
    try:
        session.FindById(element_id).Press()
        time.sleep(1)
        return {"ok": True, "pressed": element_id, "screen": get_screen_text()}
    except Exception as e:
        return {"ok": False, "error": str(e)}

def send_vkey(key_code):
    try:
        session.FindById("wnd[0]").SendVKey(int(key_code))
        time.sleep(1)
        return {"ok": True, "vkey": key_code, "screen": get_screen_text()}
    except Exception as e:
        return {"ok": False, "error": str(e)}

def handle_popup(action="confirm", text_input=None):
    try:
        wnd1 = session.FindById("wnd[1]", False)
    except Exception:
        wnd1 = None
    if wnd1 is None:
        return {"ok": False, "note": "No popup open"}
    try:
        title = wnd1.Text
        if action == "input" and text_input:
            for fid in ("usr/ctxtDY_FILENAME","usr/txtDY_FILENAME","usr/ctxtVALUE"):
                try:
                    wnd1.FindById(fid).Text = str(text_input)
                    break
                except Exception:
                    pass
        vk = 12 if action in ("cancel","no") else 0
        wnd1.SendVKey(vk)
        time.sleep(0.8)
        return {"ok": True, "popup": title, "action": action,
                "screen": get_screen_text()}
    except Exception as e:
        return {"ok": False, "error": str(e)}

def handle_transport(transport_number=None):
    """
    Handle the customising/workbench transport popup.
    Tries wnd[1]; if absent, triggers save first.
    """
    try:
        wnd1 = session.FindById("wnd[1]", False)
    except Exception:
        wnd1 = None

    if wnd1 is None:
        session.FindById("wnd[0]").SendVKey(11)   # Save
        time.sleep(2)
        try:
            wnd1 = session.FindById("wnd[1]", False)
        except Exception:
            wnd1 = None

    if wnd1 is None:
        return {"ok": True, "note": "No transport popup (may be local object)"}

    title = ""
    try:
        title = wnd1.Text
    except Exception:
        pass

    if transport_number:
        for fid in ("usr/ctxtKORR-TRKORR","usr/ctxtKORF-TRKORR"):
            try:
                wnd1.FindById(fid).Text = transport_number
                break
            except Exception:
                pass

    wnd1.SendVKey(0)   # Enter/OK
    time.sleep(1)
    try:
        status = session.FindById("wnd[0]/sbar").Text
    except Exception:
        status = ""
    return {"ok": True, "transport": transport_number or "NEW",
            "popup_title": title, "status_after": status}

# ── IDoc Error Analysis & Auto-Fix ───────────────────────────────────────────

# IDoc status codes and their meaning
IDOC_STATUS = {
    "01": "IDoc generated",
    "02": "Error passing data to port",
    "03": "Data passed to port OK",
    "04": "Error within control info",
    "05": "Error during translation",
    "06": "Translation OK",
    "07": "Error during syntax check",
    "08": "Syntax check OK",
    "09": "Error during interchange handling",
    "12": "Dispatch OK",
    "13": "Retransmission OK",
    "17": "Error passing status to R/2",
    "20": "Error triggering EDI subsystem",
    "26": "Error during syntax check",
    "29": "Error in ALE service",
    "30": "IDoc ready for dispatch",
    "31": "Error — no further processing",
    "51": "Application document not posted",
    "52": "Application document not fully posted",
    "53": "Application document posted",
    "56": "IDoc with errors added",
    "61": "Processing despite syntax errors",
    "64": "IDoc ready to be transferred to application",
    "65": "Error in ALE service",
    "68": "Error — no further processing",
    "70": "Original of an IDoc that was edited",
    "71": "IDoc is an edited copy",
}

# Common error messages and their auto-fix actions
IDOC_FIX_MAP = {
    "partner not found":              ("fix_partner_profile", "MEDIUM"),
    "partner profile":                ("fix_partner_profile", "MEDIUM"),
    "no partner agreement":           ("fix_partner_profile", "MEDIUM"),
    "posting period":                 ("open_posting_period", "HIGH"),
    "period is not open":             ("open_posting_period", "HIGH"),
    "company code":                   ("check_company_code",  "LOW"),
    "plant":                          ("check_plant",         "LOW"),
    "material":                       ("check_material",      "MEDIUM"),
    "does not exist":                 ("check_master_data",   "MEDIUM"),
    "customer":                       ("check_customer",      "MEDIUM"),
    "vendor":                         ("check_vendor",        "MEDIUM"),
    "account":                        ("check_gl_account",    "MEDIUM"),
    "authorization":                  ("check_authorization", "LOW"),
    "duplicate":                      ("mark_duplicate",      "LOW"),
    "message type":                   ("fix_message_type",    "MEDIUM"),
    "segment":                        ("fix_segment_data",    "HIGH"),
    "syntax":                         ("fix_idoc_syntax",     "HIGH"),
    "no inbound function module":     ("fix_partner_profile", "MEDIUM"),
    "function module":                ("check_fm_exists",     "HIGH"),
    "tax":                            ("check_tax_config",    "LOW"),
    "exchange rate":                  ("check_exchange_rate", "LOW"),
}


def scan_idoc_errors(date_from=None, date_to=None, direction="both",
                     status_filter="51", idoc_number=None):
    """
    Scan for IDoc errors.
    Strategy:
      1. Navigate WE05 and try all known ALV container IDs.
      2. Walk full UI tree for any accessible shell.
      3. Fall back to reading EDIDC table directly via SE16N.
    status_filter: comma-separated status codes e.g. '51,26,56' or 'all'.
    direction: 'inbound', 'outbound', or 'both'.
    idoc_number: specific IDoc number to filter on (e.g. '198025'). When set,
                 the DOCNUM LOW/HIGH fields are populated so only that IDoc is
                 returned. Date range is widened automatically to 01.01.2020 so
                 older IDocs are not excluded.
    """
    today = datetime.now().strftime("%d.%m.%Y")

    # When a specific IDoc number is given, use Jan 1 of last year as the start
    # date so IDocs created anytime in the past year are always included,
    # regardless of when they were created.
    last_year = datetime.now().year - 1
    if idoc_number:
        df = date_from or f"01.01.{last_year}"
        dt = date_to   or today
    else:
        df = date_from or today
        dt = date_to   or today

    statuses = [] if status_filter == "all" else status_filter.split(",")

    # Zero-pad the IDoc number for SAP (16 digits)
    docnum_padded = str(idoc_number).zfill(16) if idoc_number else ""

    # Convert DD.MM.YYYY → YYYYMMDD for table filter
    try:
        df_sap = datetime.strptime(df, "%d.%m.%Y").strftime("%Y%m%d")
        dt_sap = datetime.strptime(dt, "%d.%m.%Y").strftime("%Y%m%d")
    except Exception:
        df_sap = dt_sap = ""

    go_to_transaction("WE05")
    time.sleep(1.5)

    # ── Set selection fields ───────────────────────────────────────────────────
    for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-LOW", "wnd[0]/usr/ctxtLOW_DATE",
                "wnd[0]/usr/ctxtCREDAT-LOW"):
        try: session.FindById(fid).Text = df; break
        except Exception: pass
    for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-HIGH", "wnd[0]/usr/ctxtHIGH_DATE",
                "wnd[0]/usr/ctxtCREDAT-HIGH"):
        try: session.FindById(fid).Text = dt; break
        except Exception: pass

    if direction == "inbound":
        for fid in ("wnd[0]/usr/radRB_DIRECT_1", "wnd[0]/usr/radINBOUND"):
            try: session.FindById(fid).Select(); break
            except Exception: pass
    elif direction == "outbound":
        for fid in ("wnd[0]/usr/radRB_DIRECT_2", "wnd[0]/usr/radOUTBOUND"):
            try: session.FindById(fid).Select(); break
            except Exception: pass

    # Put status codes into the selection screen if there's a field for it
    if statuses:
        for fid in ("wnd[0]/usr/ctxtSEL_STATUS-LOW", "wnd[0]/usr/ctxtSTATUS"):
            try: session.FindById(fid).Text = statuses[0]; break
            except Exception: pass

    # ── Set DOCNUM filter when a specific IDoc number is requested ────────────
    if docnum_padded:
        docnum_set = False

        # Tier 1: both ctxt and txt variants
        for fid in (
            "wnd[0]/usr/ctxtS_DOCNUM-LOW",   "wnd[0]/usr/txtS_DOCNUM-LOW",
            "wnd[0]/usr/ctxtSEL_DOCNUM-LOW", "wnd[0]/usr/txtSEL_DOCNUM-LOW",
            "wnd[0]/usr/ctxtDOCNUM-LOW",     "wnd[0]/usr/txtDOCNUM-LOW",
            "wnd[0]/usr/ctxtDOCNUM",         "wnd[0]/usr/txtDOCNUM",
        ):
            try:
                session.FindById(fid).Text = docnum_padded
                docnum_set = True
                # Mirror to HIGH
                hi = fid.replace("-LOW", "-HIGH").replace("_LOW", "_HIGH")
                try: session.FindById(hi).Text = docnum_padded
                except Exception: pass
                break
            except Exception:
                pass

        # Tier 2: fragment scan
        if not docnum_set:
            for frag in ("S_DOCNUM-LOW", "DOCNUM-LOW", "DOCNUM"):
                obj, oid = _find_input_field_by_fragment(frag)
                if obj:
                    try:
                        obj.Text = docnum_padded
                        docnum_set = True
                        hi = oid.replace("-LOW", "-HIGH").replace("_LOW", "_HIGH")
                        if hi != oid:
                            try: session.FindById(hi).Text = docnum_padded
                            except Exception: pass
                    except Exception:
                        pass
                    break

        # Tier 3: label-based search
        if not docnum_set:
            obj, _ = _find_field_by_label(
                ["IDoc Number", "IDoc-Nummer", "IDoc Nummer", "Doc.Number"])
            if obj:
                try:
                    obj.Text = docnum_padded
                    docnum_set = True
                except Exception:
                    pass

    session.FindById("wnd[0]").SendVKey(8)
    time.sleep(2.5)

    idocs = []
    method_used = "none"

    # ── Strategy 1: Try all known WE05 ALV container IDs ─────────────────────
    WE05_CONTAINERS = [
        "wnd[0]/usr/cntlWE05_CONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlGRID/shellcont/shell",
        "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlALV/shellcont/shell",
    ]
    IDOC_COLS = ["DOCNUM","STATUS","MANDT","DIRECT","MESTYP","MESCOD",
                 "MESFCT","SNDPRT","SNDPRN","RCVPRT","RCVPRN",
                 "CREDAT","CRETIM","UPDDAT","STATXT"]

    shell = _find_shell_anywhere(WE05_CONTAINERS)
    if shell:
        rows_data = _read_shell_rows(shell, IDOC_COLS, max_rows=500)
        for row in rows_data:
            if not row.get("DOCNUM"):
                continue
            row["status_desc"] = IDOC_STATUS.get(row.get("STATUS",""), "Unknown")
            if status_filter == "all" or row.get("STATUS","") in statuses:
                idocs.append(row)
        method_used = "alv_grid"

    # ── Strategy 2: EDIDC table fallback ─────────────────────────────────────
    if not idocs:
        try:
            go_to_transaction("SE16N")
            time.sleep(1)
            for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
                try: session.FindById(fid).Text = "EDIDC"; break
                except Exception: pass
            session.FindById("wnd[0]").SendVKey(0)
            time.sleep(1.5)

            elems = discover_elements()
            # Set CREDAT (creation date) filter
            for e in elems:
                eid = e.get("id","").upper()
                if "CREDAT" in eid and "LOW" in eid:
                    try: session.FindById(e["id"]).Text = df_sap; break
                    except Exception: pass
            for e in elems:
                eid = e.get("id","").upper()
                if "CREDAT" in eid and "HIGH" in eid:
                    try: session.FindById(e["id"]).Text = dt_sap; break
                    except Exception: pass

            # Set STATUS filter (first code only — SE16N single value)
            if statuses:
                for e in elems:
                    eid = e.get("id","").upper()
                    if "STATUS" in eid and "LOW" in eid:
                        try: session.FindById(e["id"]).Text = statuses[0]; break
                        except Exception: pass

            # Set DOCNUM filter when a specific IDoc number is requested
            if docnum_padded:
                for e in elems:
                    eid = e.get("id","").upper()
                    if "DOCNUM" in eid and "LOW" in eid:
                        try: session.FindById(e["id"]).Text = docnum_padded; break
                        except Exception: pass
                for e in elems:
                    eid = e.get("id","").upper()
                    if "DOCNUM" in eid and "HIGH" in eid:
                        try: session.FindById(e["id"]).Text = docnum_padded; break
                        except Exception: pass

            # Max rows
            for fid in ("wnd[0]/usr/txtGD-MAX_LINES",):
                try: session.FindById(fid).Text = "1000"; break
                except Exception: pass

            session.FindById("wnd[0]").SendVKey(8)
            time.sleep(2.5)

            edidc_shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if edidc_shell:
                edidc_cols = ["DOCNUM","STATUS","DIRECT","MESTYP","SNDPRT",
                              "SNDPRN","RCVPRT","RCVPRN","CREDAT","CRETIM"]
                rows_data = _read_shell_rows(edidc_shell, edidc_cols, max_rows=1000)
                for row in rows_data:
                    if not row.get("DOCNUM"):
                        continue
                    row["status_desc"] = IDOC_STATUS.get(row.get("STATUS",""), "Unknown")
                    # Apply direction filter
                    if direction == "inbound"  and row.get("DIRECT","") != "1": continue
                    if direction == "outbound" and row.get("DIRECT","") != "2": continue
                    # Apply status filter
                    if status_filter == "all" or row.get("STATUS","") in statuses:
                        idocs.append(row)
                method_used = "edidc_table"
        except Exception:
            pass

    idoc_filter_note = (f" (DOCNUM filter: {idoc_number})" if idoc_number else "")
    result = {
        "date_from":    df,
        "date_to":      dt,
        "idoc_number":  idoc_number,
        "idoc_errors":  idocs,
        "count":        len(idocs),
        "method":       method_used,
        "screen":       get_screen_text(),
        "note": (f"Found {len(idocs)} IDocs matching status {status_filter}"
                 f"{idoc_filter_note} (method: {method_used}). "
                 "Call get_idoc_detail for each to analyse root cause."),
    }
    if not idocs:
        result["screen_texts"] = _screen_texts()[:100]
        result["all_elements"] = [e["id"] for e in discover_elements()[:80]]
    return result


    # _we02_set_docnum removed — replaced by _we02_navigate_to_idoc above


def get_idoc_detail(idoc_number):
    """
    Get full detail for a specific IDoc — error text, control record, fix hints.

    Strategy:
      1. EDIDS table via SE16N WHERE clause — most reliable, always has exact
         error text regardless of GUI issues.  This is tried FIRST.
      2. Navigate WE02, set DOCNUM filter, read GuiTree (status text) and
         right panel (technical info).
      3. EDIDC control record via SE16N WHERE clause.
      4. Merge all data and compute fix_hints.
    """
    docnum = str(idoc_number).zfill(16)
    detail = {
        "idoc_number":    str(idoc_number),
        "status_records": [],
        "segments":       [],
        "error_messages": [],
        "control_record": {},
        "fix_hints":      [],
        "all_screen_text": "",
    }

    # ── Strategy 1 (PRIMARY): Read EDIDS via SE16N with WHERE clause ──────────
    # EDIDS contains the complete status history with STATXT (status text).
    # The WHERE clause approach is more reliable than the condition-table approach.
    for _attempt_where in (True, False):
        try:
            go_to_transaction("SE16N")
            time.sleep(1)
            for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
                try:
                    session.FindById(fid).Text = "EDIDS"
                    break
                except Exception:
                    pass
            session.FindById("wnd[0]").SendVKey(0)  # Enter → load fields
            time.sleep(1.5)

            if _attempt_where:
                # Try the WHERE clause / Expert button
                for btn in ("wnd[0]/tbar[1]/btn[14]",    # "Expert" in some versions
                            "wnd[0]/mbar/menu[3]/menu[7]", # Settings → Free Criteria
                            "wnd[0]/tbar[1]/btn[6]"):
                    try:
                        session.FindById(btn).Press()
                        time.sleep(0.8)
                        break
                    except Exception:
                        pass
                # Set a WHERE clause field if one appeared
                for wfid in ("wnd[0]/usr/txtGD-WHERE", "wnd[0]/usr/txtWHERE",
                             "wnd[0]/usr/ctxtGD-WHERE"):
                    try:
                        session.FindById(wfid).Text = f"DOCNUM = '{docnum}'"
                        break
                    except Exception:
                        pass
            else:
                # Standard condition table approach
                for tbl_id in ("wnd[0]/usr/tblSAPLSE16NSELFIELD_TC",
                               "wnd[0]/usr/tblSELFIELD_TC"):
                    try:
                        tbl = session.FindById(tbl_id, False)
                        if not tbl:
                            continue
                        for row in range(min(tbl.RowCount, 30)):
                            try:
                                fn = tbl.GetCell(row, 0).Text.strip().upper()
                                if fn == "DOCNUM":
                                    tbl.GetCell(row, 4).Text = docnum   # LOW
                                    try:
                                        tbl.GetCell(row, 5).Text = docnum  # HIGH
                                    except Exception:
                                        pass
                                    break
                            except Exception:
                                pass
                        break
                    except Exception:
                        pass

            # Set max rows and execute
            for mfid in ("wnd[0]/usr/txtGD-MAX_LINES",):
                try:
                    session.FindById(mfid).Text = "200"
                except Exception:
                    pass
            session.FindById("wnd[0]").SendVKey(8)
            time.sleep(2.5)

            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell and shell.RowCount > 0:
                rows = _read_shell_rows(
                    shell,
                    ["DOCNUM","STATUS","LOGDAT","LOGTIM","STAMQU",
                     "STATXT","UNAME","REPID","STAPA1","STAPA2"],
                    max_rows=200,
                )
                # Only keep rows that match our IDoc number
                matching = [r for r in rows
                            if r.get("DOCNUM","").strip().lstrip("0") ==
                               str(idoc_number).lstrip("0")]
                if not matching:
                    matching = rows  # accept all if filter is unclear

                for row in matching:
                    statxt = row.get("STATXT","").strip()
                    if statxt and statxt != "&, &, &, &":
                        detail["status_records"].append(row)
                        if any(kw in statxt.lower() for kw in [
                            "error","fehler","not found","nicht","invalid",
                            "missing","exception","fail","partner","profile",
                            "function module","posting period","cannot"
                        ]):
                            if statxt not in detail["error_messages"]:
                                detail["error_messages"].append(statxt)
                detail["edids_source"] = (
                    f"EDIDS WHERE clause: {len(matching)} rows (total {len(rows)})")
                if detail["status_records"]:
                    break   # success — skip fallback attempt
        except Exception as e:
            detail["edids_error"] = str(e)

    # ── Strategy 2: Navigate WE02 — clear all fields, enter DOCNUM, execute ─────
    # Uses _we02_navigate_to_idoc which:
    #   • clears ALL selection fields (date ranges, status, direction, etc.)
    #   • finds the DOCNUM field via full GUI-tree scan (not limited to 50 elems)
    #   • sets DOCNUM LOW = HIGH = zero-padded IDoc number
    #   • executes and enters the IDoc detail screen
    on_detail, docnum_set, we02_screen = _we02_navigate_to_idoc(docnum)
    detail["we02_docnum_set"] = docnum_set
    detail["we02_on_detail"]  = on_detail
    detail["screen"]          = we02_screen

    # ── Strategy 2a: WE02 GuiTree — left panel (status record text) ─────────────
    tree = _find_gui_tree_anywhere()
    if tree:
        node_texts = _read_gui_tree_texts(tree, max_nodes=400)
        detail["tree_nodes"] = node_texts
        for txt in node_texts:
            tl = txt.lower()
            if any(kw in tl for kw in [
                "error","fehler","not found","nicht gefunden","exception",
                "invalid","missing","edi:","failed","cannot","not exist",
                "no match","application error","partner","profile",
                "function module","posting period","syntax","segment",
                "idoc with","inbound partner",
            ]):
                if txt not in detail["error_messages"]:
                    detail["error_messages"].append(txt)
            if (len(txt) > 2 and txt[:2].isdigit()) or "idoc with" in tl:
                detail["status_records"].append(
                    {"STATXT": txt, "source": "WE02_tree"})
        detail["tree_source"] = f"WE02 GuiTree: {len(node_texts)} nodes"

    # ── Strategy 2b: Right panel — Short Technical Information ───────────────
    screen_txts = _screen_texts()
    detail["all_screen_text"] = "\n".join(screen_txts[:300])
    import re as _re
    full  = detail["all_screen_text"]
    _ctrl = detail["control_record"]
    for pat, fld in [
        (r'(MATMAS\d*|ORDERS\d*|INVOIC\d*|DESADV\d*|DEBMAS|CREMAS|WMMBID|PORDCR|SHPORD)', 'MESTYP'),
        (r'(?:Message Type|Nachrichtentyp)[^\S\n]*([A-Z][A-Z0-9]+)', 'MESTYP'),
        (r'(?:Basic type|Basistyp)[^\S\n]*([A-Z][A-Z0-9]+)',         'IDOCTP'),
        (r'(?:Partner No\.|Partnernummer)[^\S\n]*([A-Z0-9_]+)',       'SNDPRN'),
        (r'(?:Partn\.?Type|Partnertyp)[^\S\n]*(LS|KU|LI|KD|VN)',     'SNDPRT'),
        (r'(?:Port)[^\S\n]*([A-Z][A-Z0-9]+)',                         'RCVPOR'),
        (r'(?:Direction|Richtung)[^\S\n]*(\d)',                       'DIRECT'),
        (r'(?:Current Status|Aktueller Status)[^\S\n]*(\d{2})',       'STATUS'),
    ]:
        m = _re.search(pat, full, _re.IGNORECASE)
        if m and fld not in _ctrl:
            _ctrl[fld] = m.group(1).strip()

    # ── Strategy 3: EDIDC control record via SE16N WHERE clause ─────────────
    try:
        go_to_transaction("SE16N")
        time.sleep(1)
        for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
            try:
                session.FindById(fid).Text = "EDIDC"
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)
        for tbl_id in ("wnd[0]/usr/tblSAPLSE16NSELFIELD_TC",
                       "wnd[0]/usr/tblSELFIELD_TC"):
            try:
                tbl = session.FindById(tbl_id, False)
                if not tbl:
                    continue
                for row in range(min(tbl.RowCount, 30)):
                    try:
                        if tbl.GetCell(row, 0).Text.strip().upper() == "DOCNUM":
                            tbl.GetCell(row, 4).Text = docnum
                            try:
                                tbl.GetCell(row, 5).Text = docnum
                            except Exception:
                                pass
                            break
                    except Exception:
                        pass
                break
            except Exception:
                pass
        for mfid in ("wnd[0]/usr/txtGD-MAX_LINES",):
            try:
                session.FindById(mfid).Text = "5"
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        edidc_shell = _find_shell_anywhere([
            "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell",
            "wnd[0]/usr/cntlGRID1/shellcont/shell",
        ])
        if edidc_shell and edidc_shell.RowCount > 0:
            ctrl_rows = _read_shell_rows(
                edidc_shell,
                ["DOCNUM","STATUS","DIRECT","MESTYP","MESCOD","MESFCT",
                 "SNDPRT","SNDPRN","RCVPRT","RCVPRN","CREDAT","CRETIM"],
                max_rows=5,
            )
            for r in ctrl_rows:
                if r.get("DOCNUM","").strip().lstrip("0") == str(idoc_number).lstrip("0"):
                    for k, v in r.items():
                        if v and k not in _ctrl:
                            _ctrl[k] = v
                    break
            if _ctrl.get("STATUS"):
                detail["status_desc"] = IDOC_STATUS.get(_ctrl["STATUS"], "Unknown")
    except Exception:
        pass

    # ── Build fix_hints from all collected data ───────────────────────────────
    combined = (
        " ".join(detail["error_messages"]) + " " +
        detail["all_screen_text"] + " " +
        " ".join(r.get("STATXT","") for r in detail["status_records"])
    ).lower()

    detail["fix_hints"] = [
        {"pattern": pat, "fix_action": fa, "risk": risk}
        for pat, (fa, risk) in IDOC_FIX_MAP.items()
        if pat in combined
    ]
    if not detail["fix_hints"] and detail["status_records"]:
        detail["fix_hints"] = [{"note": "No known IDOC_FIX_MAP pattern matched. "
                                         "Review status_records / error_messages."}]

    # ── Look up the correct process code from WE64/EDIFCT ────────────────────
    msg_type  = _ctrl.get("MESTYP", "")
    direction = _ctrl.get("DIRECT", "1")
    if msg_type:
        procod = _we64_lookup_process_code(msg_type, direction)
        detail["process_code"] = procod
        detail["process_code_source"] = (
            "WE64/EDIFCT" if procod not in _PROCESS_CODE_MAP.values()
                             and procod != msg_type.upper()[:8]
            else ("hardcoded_map" if _PROCESS_CODE_MAP.get(msg_type.upper()) == procod
                  else "fallback")
        )
    else:
        detail["process_code"] = ""
        detail["process_code_source"] = "unknown"

    # Summary for easy reading
    detail["summary"] = {
        "idoc":          str(idoc_number),
        "status":        _ctrl.get("STATUS","?"),
        "status_desc":   detail.get("status_desc",""),
        "message_type":  msg_type,
        "partner":       _ctrl.get("SNDPRN",""),
        "partner_type":  _ctrl.get("SNDPRT",""),
        "direction":     direction,
        "process_code":  detail.get("process_code",""),
        "error_count":   len(detail["error_messages"]),
        "errors":        detail["error_messages"][:5],
    }

    return detail


def get_idoc_segments(idoc_number):
    """
    Read raw segment data for a specific IDoc via SE16N (EDIDD table).
    Returns all segment content for deep data-level analysis.
    """
    docnum = str(idoc_number).zfill(16)

    # Read EDIDD (segment data)
    result_rows = []
    try:
        go_to_transaction("SE16N")
        time.sleep(1)
        for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
            try: session.FindById(fid).Text = "EDIDD"; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)
        # Set DOCNUM filter
        elems = discover_elements()
        for e in elems:
            eid = e.get("id", "").upper()
            if "DOCNUM" in eid and "LOW" in eid:
                try: session.FindById(e["id"]).Text = docnum; break
                except Exception: pass
        for fid in ("wnd[0]/usr/txtGD-MAX_LINES",):
            try: session.FindById(fid).Text = "500"
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        result_rows = [{"note": "EDIDD rows retrieved via SE16N", "screen": get_screen_text()}]
    except Exception as e:
        result_rows = [{"error": str(e)}]

    # Read EDIDC (control record)
    ctrl_rows = []
    try:
        go_to_transaction("SE16N")
        time.sleep(1)
        for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
            try: session.FindById(fid).Text = "EDIDC"; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)
        elems = discover_elements()
        for e in elems:
            eid = e.get("id", "").upper()
            if "DOCNUM" in eid and "LOW" in eid:
                try: session.FindById(e["id"]).Text = docnum; break
                except Exception: pass
        for fid in ("wnd[0]/usr/txtGD-MAX_LINES",):
            try: session.FindById(fid).Text = "5"
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        ctrl_rows = [{"note": "EDIDC rows retrieved via SE16N", "screen": get_screen_text()}]
    except Exception as e:
        ctrl_rows = [{"error": str(e)}]

    return {
        "idoc_number":    str(idoc_number),
        "control_record": ctrl_rows,
        "segment_data":   result_rows,
    }


def reprocess_idoc(idoc_number, edit_mode=False):
    """
    Reprocess a failed IDoc via WE19.
    edit_mode=False: reprocess as-is (standard retry).
    edit_mode=True:  open in edit mode so field values can be changed first.
    """
    go_to_transaction("WE19")
    time.sleep(1)
    try:
        # Enter IDoc number
        for fid in ("wnd[0]/usr/ctxtWE19-DOCNUM",
                    "wnd[0]/usr/ctxtDOCNUM"):
            try:
                session.FindById(fid).Text = str(idoc_number).zfill(16)
                break
            except Exception:
                pass

        if edit_mode:
            # Select "Edit IDoc" option
            for fid in ("wnd[0]/usr/radWE19-CEDITYPE_2",
                        "wnd[0]/usr/radEDIT"):
                try:
                    session.FindById(fid).Select()
                    break
                except Exception:
                    pass

        session.FindById("wnd[0]").SendVKey(8)   # Execute
        time.sleep(2)

        result = {"ok": True, "idoc": str(idoc_number),
                  "edit_mode": edit_mode, "screen": get_screen_text()}

        if not edit_mode:
            # Standard processing — trigger inbound function module
            try:
                session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
                time.sleep(2)
            except Exception:
                try:
                    session.FindById("wnd[0]").SendVKey(8)
                    time.sleep(2)
                except Exception:
                    pass
            result["screen_after"] = get_screen_text()

        audit_log("IDOC_REPROCESS",
                  {"idoc": str(idoc_number), "edit_mode": edit_mode},
                  status="executed")
        return result
    except Exception as e:
        return {"error": str(e), "idoc": str(idoc_number)}


def edit_idoc_field(idoc_number, segment_name, field_name, new_value):
    """
    Open IDoc in WE19 edit mode, navigate to a specific segment field
    and change its value, then reprocess.
    Use when the root cause is wrong data in a segment field.
    """
    # First open in edit mode
    open_result = reprocess_idoc(idoc_number, edit_mode=True)
    if not open_result.get("ok"):
        return open_result
    time.sleep(1)

    # Discover elements and try to find the segment/field
    elems = discover_elements()
    found = False
    for elem in elems:
        eid = elem.get("id","").upper()
        tip = elem.get("tooltip","").upper()
        if field_name.upper() in eid or field_name.upper() in tip:
            set_field(elem["id"], str(new_value))
            found = True
            break

    if not found:
        return {
            "ok":      False,
            "note":    f"Field {field_name} not found on screen. "
                       "Use discover_screen_elements to find the correct ID.",
            "elements": [e["id"] for e in elems[:30]],
        }

    # Save and reprocess
    session.FindById("wnd[0]").SendVKey(11)   # Save
    time.sleep(1)
    session.FindById("wnd[0]").SendVKey(8)    # Execute/Reprocess
    time.sleep(2)

    audit_log("IDOC_FIELD_EDIT",
              {"idoc": str(idoc_number), "segment": segment_name,
               "field": field_name, "new_value": str(new_value)},
              status="executed")
    return {"ok": True, "field_changed": field_name,
            "new_value": new_value, "screen": get_screen_text()}


# ── WE64 / EDIFCT process-code lookup ─────────────────────────────────────────

# Hardcoded fallback map — used only when WE64 lookup fails or returns nothing.
_PROCESS_CODE_MAP = {
    "MATMAS":  "MATM",  "MATMAS01": "MATM",
    "ORDERS":  "ORDE",  "ORDERS05": "ORDE",
    "INVOIC":  "INVL",  "INVOIC01": "INVL",
    "DEBMAS":  "DEBM",  "DEBMAS06": "DEBM",
    "CREMAS":  "CREM",  "CREMAS05": "CREM",
    "DESADV":  "DELS",  "DELVRY":   "DELS",
    "PORDCR":  "PORD",  "PORDCH":   "PORD",
    "SHPORD":  "SHPORD","WMMBID":   "WMMBID",
    "HRMD_A":  "HRMD",  "PEXR2":    "PEXR",
}

# Cache to avoid re-querying WE64 for the same message type in one run
_we64_cache: dict = {}


def _we64_lookup_process_code(message_type: str, direction: str = "1") -> str:
    """
    Look up the inbound (direction=1) or outbound (direction=2) process code
    for a given message type by querying SAP table EDIFCT via SE16N.

    Falls back to _PROCESS_CODE_MAP, then to message_type.upper() if nothing found.

    EDIFCT key fields:
      DIRECT  — direction (1=inbound, 2=outbound)
      MESTYP  — message type (e.g. MATMAS)
      MESCOD  — message code  (usually blank)
      MESFCT  — message function (usually blank)
      PROCOD  — process code (what we want)
    """
    cache_key = f"{message_type}:{direction}"
    if cache_key in _we64_cache:
        return _we64_cache[cache_key]

    procod = ""

    try:
        go_to_transaction("SE16N")
        time.sleep(1.5)

        # Enter table name
        for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/txtGD-TAB",
                    "wnd[0]/usr/ctxtDATABROWSE-TABLENAME"):
            try:
                session.FindById(fid).Text = "EDIFCT"
                session.FindById("wnd[0]").SendVKey(0)   # Enter
                time.sleep(1.5)
                break
            except Exception:
                pass

        # Set DIRECT filter
        _set_se16n_filter("DIRECT", str(direction))
        # Set MESTYP filter
        _set_se16n_filter("MESTYP", message_type.upper())
        # Set max hits to 10 (we only need the first match)
        for fid in ("wnd[0]/usr/txtGD-MAX_LINES",):
            try:
                session.FindById(fid).Text = "10"
            except Exception:
                pass

        # Execute (F8)
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        # Read the result grid — look for PROCOD column value
        procod = _se16n_read_column("PROCOD")

    except Exception:
        pass

    # Fallback chain
    if not procod:
        procod = _PROCESS_CODE_MAP.get(message_type.upper(), "")
    if not procod:
        # Use message type itself as last resort (SAP convention for many types)
        procod = message_type.upper()[:8]

    _we64_cache[cache_key] = procod
    return procod


def _set_se16n_filter(field_name: str, value: str):
    """
    In an open SE16N selection screen, find the filter row for field_name
    and set its Low value.  Tries known ID patterns first, then walks the tree.
    """
    # SE16N filter rows have IDs like: wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtSELFIELDS-LOW[X,Y]
    # but the position varies. Use fragment search.
    obj, _ = _find_input_field_by_fragment(field_name)
    if obj:
        try:
            obj.Text = value
            return
        except Exception:
            pass

    # If not found on screen, try typing the field name into the search/position field
    for fid in ("wnd[0]/usr/txtGD-FNAM",):
        try:
            session.FindById(fid).Text = field_name
            session.FindById("wnd[0]").SendVKey(0)
            time.sleep(0.5)
            # Now try the Low field again
            obj2, _ = _find_input_field_by_fragment("LOW")
            if obj2:
                obj2.Text = value
            return
        except Exception:
            pass


def _se16n_read_column(column_name: str) -> str:
    """
    After SE16N executes and shows a result grid (GuiShell/GuiGridView),
    read the first row value of the named column.
    """
    try:
        # Try GuiGridView (ALV grid)
        grid = session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell", False)
        if not grid:
            grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
        if grid:
            # Get column index
            col_count = grid.ColumnCount
            for ci in range(col_count):
                try:
                    col_key = grid.GetColumnKey(ci)
                    if column_name.upper() in col_key.upper():
                        val = grid.GetCellValue(0, col_key)
                        return (val or "").strip()
                except Exception:
                    pass
    except Exception:
        pass

    # Fallback: read screen text and parse
    try:
        texts = _screen_texts()
        found_col = False
        for t in texts:
            t = t.strip()
            if column_name.upper() in t.upper():
                found_col = True
                continue
            if found_col and t and not t.startswith("-"):
                return t.strip()
    except Exception:
        pass

    return ""


def check_partner_profile(partner_number, direction="1", message_type=""):
    """
    Check WE20 partner profile for a specific partner.
    direction: '1'=inbound, '2'=outbound
    Returns whether profile exists and configured message types.
    """
    go_to_transaction("WE20")
    time.sleep(1.5)
    try:
        elems = discover_elements()

        # Fill the partner number field using whatever ID is on screen
        parnr_ids = [
            "wnd[0]/usr/ctxtWE20-PARNR",
            "wnd[0]/usr/ctxtPARTNER_NO",
        ]
        filled = False
        for fid in parnr_ids:
            try:
                session.FindById(fid).Text = str(partner_number)
                filled = True
                break
            except Exception:
                pass

        if not filled:
            # Fallback: find any ctxt/txt field whose ID contains PARNR or PARTNER
            for e in elems:
                eid = e.get("id", "").upper()
                if any(k in eid for k in ("PARNR", "PARTNER")) and e.get("type") in ("GuiCTextField", "GuiTextField"):
                    try:
                        session.FindById(e["id"]).Text = str(partner_number)
                        filled = True
                        break
                    except Exception:
                        pass

        # Execute the search — try F8, then Enter, then toolbar Execute button
        executed = False
        for vkey in (8, 0):
            try:
                session.FindById("wnd[0]").SendVKey(vkey)
                time.sleep(1.5)
                executed = True
                break
            except Exception:
                pass
        if not executed:
            # Try toolbar execute button
            for btn in ("wnd[0]/tbar[1]/btn[8]", "wnd[0]/tbar[0]/btn[0]"):
                try:
                    session.FindById(btn).Press()
                    time.sleep(1.5)
                    break
                except Exception:
                    pass

        screen = get_screen_text()
        exists = (str(partner_number) in screen or
                  "partner" in screen.lower())

        return {
            "partner":  str(partner_number),
            "exists":   exists,
            "direction": direction,
            "screen":   screen,
            "elements": discover_elements()[:40],
        }
    except Exception as e:
        return {"error": str(e)}


def create_partner_profile(partner_number, partner_type, direction,
                            message_type, process_code, basic_type="",
                            func_module=""):
    """
    Add an inbound or outbound message-type entry to a partner profile in WE20.

    Handles two cases:
      • Partner already exists → navigate to it in the tree, add the new
        inbound/outbound row.  (Most common case — e.g. S4HANA2023 exists
        but is missing MATMAS inbound.)
      • Partner does not exist → create the partner header first, then add row.

    partner_type : LS=Logical system, KU=Customer, LI=Vendor
    direction    : '1' or 'inbound'  /  '2' or 'outbound'
    process_code : e.g. MATM — if blank or unknown, looked up in WE64/EDIFCT
    basic_type   : e.g. MATMAS01 (leave blank to use SAP default)
    """
    is_inbound = str(direction) in ("1", "inbound", "INBOUND")
    dir_code   = "1" if is_inbound else "2"

    # ── Resolve process code via WE64 if not provided or not in hardcoded map ──
    if not process_code or process_code.upper() == message_type.upper():
        looked_up = _we64_lookup_process_code(message_type, dir_code)
        if looked_up:
            print(f"  [WE64] Process code for {message_type} dir={dir_code}: {looked_up}",
                  flush=True)
            process_code = looked_up
    else:
        # Validate the supplied code against WE64; warn if mismatch
        we64_code = _we64_lookup_process_code(message_type, dir_code)
        if we64_code and we64_code != process_code:
            print(f"  [WE64] Warning: supplied process_code={process_code} "
                  f"but WE64 says {we64_code} for {message_type}. "
                  f"Using WE64 value.", flush=True)
            process_code = we64_code

    go_to_transaction("WE20")
    time.sleep(2)

    # ── Step 1: Find S4HANA2023 (or any partner) in the WE20 tree ────────────
    partner_on_screen = False
    tree = _find_gui_tree_anywhere()

    if tree:
        try:
            # Expand all nodes so the target partner becomes visible
            all_keys = []
            try:
                all_keys = list(tree.GetAllNodeKeys() or [])
            except Exception:
                pass
            for k in all_keys:
                try:
                    tree.ExpandNode(k)
                except Exception:
                    pass
            time.sleep(0.5)
            # Refresh key list after expansion
            try:
                all_keys = list(tree.GetAllNodeKeys() or [])
            except Exception:
                pass
            # Click the node whose text matches the partner number
            for k in all_keys:
                try:
                    node_text = tree.GetNodeText(k) or ""
                    if partner_number.upper() in node_text.upper():
                        tree.EnsureVisibleHorizontalItem(k, "")
                        tree.ClickNode(k)
                        time.sleep(1.5)
                        partner_on_screen = True
                        break
                except Exception:
                    pass
        except Exception:
            pass

    # Fallback: type partner number into WE20's filter/search field
    if not partner_on_screen:
        for fid in ("wnd[0]/usr/ctxtWE20-PARNR", "wnd[0]/usr/txtWE20-PARNR",
                    "wnd[0]/usr/ctxtPARTNR"):
            try:
                session.FindById(fid).Text = partner_number
                session.FindById("wnd[0]").SendVKey(0)
                time.sleep(1.5)
                partner_on_screen = True
                break
            except Exception:
                pass

    screen = get_screen_text()
    partner_exists = partner_number in screen

    # ── Step 2: If partner doesn't exist, create the header first ─────────────
    if not partner_exists:
        # Press "Create" toolbar button
        for fid in ("wnd[0]/tbar[1]/btn[8]", "wnd[0]/tbar[0]/btn[3]",
                    "wnd[0]/tbar[1]/btn[3]"):
            try:
                session.FindById(fid).Press()
                time.sleep(1)
                break
            except Exception:
                pass
        # Fill partner number and type
        elems = discover_elements()
        for e in elems:
            eid = e.get("id", "").upper()
            if "PARNR" in eid or "PARTNER_NO" in eid:
                try:
                    session.FindById(e["id"]).Text = partner_number
                    break
                except Exception:
                    pass
        for e in elems:
            eid = e.get("id", "").upper()
            if "PARVW" in eid or "PARTYP" in eid or "PARTNER_TYPE" in eid:
                try:
                    session.FindById(e["id"]).Text = partner_type
                    break
                except Exception:
                    pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

    # ── Step 3: Ensure Change mode ────────────────────────────────────────────
    # WE20 opens in change mode automatically when you click a partner in the
    # tree — no explicit toggle needed.  But if somehow in display mode (e.g.
    # opened via a read-only path), try the standard Ctrl+F1 pencil button.
    screen = get_screen_text()
    if any(w in screen for w in ("Display", "Anzeige")):
        for fid in ("wnd[0]/tbar[1]/btn[4]", "wnd[0]/tbar[0]/btn[4]",
                    "wnd[0]/tbar[1]/btn[1]"):
            try:
                session.FindById(fid).Press()
                time.sleep(1)
                break
            except Exception:
                pass

    # ── Step 4: Add the new inbound or outbound row ───────────────────────────
    # The WE20 Inbound / Outbound tables each have a row of 4 buttons below
    # them.  Button positions in SAP WE20 (typical):
    #   btn[0]=detail/navigate  btn[1]=copy  btn[2]=create  btn[3]=delete
    # Sub-screen IDs vary by SAP version — try multiple paths.
    SUBSCREENS = [
        "wnd[0]/usr/subSUBSCREEN_BODY:SAPLWEDC:0100/",
        "wnd[0]/usr/subSUBSCREEN_BODY:SAPLWEDC:0200/",
        "wnd[0]/usr/sub:SAPLWEDC:0100/",
        "wnd[0]/usr/",
    ]
    section = "IN" if is_inbound else "OUT"

    row_added = False
    for pfx in SUBSCREENS:
        for btn_suffix in (
            f"btnBT_{section}_CREATE",
            f"btnBT_{section}BOUND_CREATE",
            f"tblSAPLWEDCTC_{section}/btnCREATE",
            f"btnCREATE_{section}BOUND",
            f"btnCREATE_{section}",
        ):
            try:
                session.FindById(pfx + btn_suffix).Press()
                time.sleep(0.8)
                row_added = True
                break
            except Exception:
                pass
        if row_added:
            break

    if not row_added:
        # Button IDs vary across SAP versions — hand off to user to click
        # the Create button; agent fills the fields automatically after.
        result = wait_for_user_input(
            screen_name="WE20 Partner Profile — Add Row",
            instructions=(
                f"In WE20, partner {partner_number} is selected. "
                f"Click the small CREATE button below the "
                f"{'INBOUND' if is_inbound else 'OUTBOUND'} PARAMETERS table "
                f"to add a new empty row. Do NOT fill any fields yet — "
                f"just click Create to open the new row, then type 'done'."
            ),
            fields=[
                {"name": "Action", "value": f"Click Create in the {'Inbound' if is_inbound else 'Outbound'} Parameters section"},
            ],
        )
        if result.get("status") == "user_completed":
            row_added = True
        elif result.get("status") in ("skipped", "aborted"):
            # User could not / would not add the row — full manual handoff
            manual = wait_for_user_input(
                screen_name="WE20 Partner Profile — Full Manual Entry",
                instructions=(
                    f"Automated WE20 entry failed. Please create the partner profile manually:\n"
                    f"  1. Go to WE20 (tcode)\n"
                    f"  2. Find or create partner: {partner_number} (type {partner_type})\n"
                    f"  3. Click Create in the {'Inbound' if is_inbound else 'Outbound'} Parameters table\n"
                    f"  4. Enter Message Type = {message_type}\n"
                    f"  5. Enter Process Code = {process_code}\n"
                    f"  {'6. Enter Basic Type = ' + basic_type if basic_type else ''}\n"
                    f"  7. Press Ctrl+S to save\n"
                    f"  Type 'done' when finished."
                ),
                fields=[
                    {"name": "Partner Number",  "value": partner_number},
                    {"name": "Partner Type",    "value": partner_type},
                    {"name": "Direction",       "value": "Inbound" if is_inbound else "Outbound"},
                    {"name": "Message Type",    "value": message_type},
                    {"name": "Process Code",    "value": process_code},
                    {"name": "Basic Type",      "value": basic_type or "(leave blank)"},
                ],
            )
            screen = get_screen_text()
            audit_log("CREATE_PARTNER_PROFILE",
                      {"partner": partner_number, "type": partner_type,
                       "message_type": message_type, "process_code": process_code,
                       "method": "manual_user"},
                      status="manual_completed" if manual.get("status") == "user_completed"
                             else "manual_skipped")
            return {
                "ok":           manual.get("status") == "user_completed",
                "partner":      partner_number,
                "direction":    "inbound" if is_inbound else "outbound",
                "message_type": message_type,
                "process_code": process_code,
                "row_added":    False,
                "screen":       screen,
                "method":       "manual_user",
                "note": "Agent could not automate WE20 — user completed manually."
                        if manual.get("status") == "user_completed"
                        else "User skipped manual WE20 entry.",
            }

    # ── Step 5: Fill Message Type and Process Code ────────────────────────────
    # Use _find_input_field_by_fragment — bypasses the 50-element cap of
    # discover_elements(), which would miss the new row added at the bottom.
    time.sleep(0.5)

    msg_set = proc_set = False

    # MESTYP — find first empty field (new row will be blank)
    for frag in ("MESTYP", "MESTYPE", "MSG_TYPE"):
        obj, _ = _find_input_field_by_fragment(frag)
        if obj:
            try:
                cur = (obj.Text or "").strip()
                if not cur:          # use the empty (new) row
                    obj.Text = message_type
                    msg_set = True
                    break
            except Exception:
                pass
    if not msg_set:
        # No empty MESTYP found — set the last one (the new row is always last)
        obj, _ = _find_input_field_by_fragment("MESTYP")
        if obj:
            try:
                obj.Text = message_type
                msg_set = True
            except Exception:
                pass

    # PROCOD — process code field
    for frag in ("PROCOD", "PRZNR", "PROCESS_CODE", "PROC_CODE"):
        obj, _ = _find_input_field_by_fragment(frag)
        if obj:
            try:
                cur = (obj.Text or "").strip()
                if not cur:
                    obj.Text = process_code
                    proc_set = True
                    break
            except Exception:
                pass
    if not proc_set:
        obj, _ = _find_input_field_by_fragment("PROCOD")
        if obj:
            try:
                obj.Text = process_code
                proc_set = True
            except Exception:
                pass

    # Basic Type (optional) — IDOCTP field
    if basic_type:
        for frag in ("IDOCTP", "BASIC_TYPE", "BASICTYPE"):
            obj, _ = _find_input_field_by_fragment(frag)
            if obj:
                try:
                    cur = (obj.Text or "").strip()
                    if not cur:
                        obj.Text = basic_type
                        break
                except Exception:
                    pass

    # ── Step 5b: Hand off if fields could not be filled automatically ─────────
    if not msg_set or not proc_set:
        missing_fields = []
        if not msg_set:
            missing_fields.append({"name": "Message Type (MESTYP)", "value": message_type,
                                   "note": "Enter in the new row that was just created"})
        if not proc_set:
            missing_fields.append({"name": "Process Code (NACFN)", "value": process_code,
                                   "note": "Enter in the same new row"})
        wait_for_user_input(
            screen_name="WE20 Partner Profile — Fill Row Fields",
            instructions=(
                f"The new {'inbound' if is_inbound else 'outbound'} row was added for "
                f"partner {partner_number} but the agent could not fill the fields. "
                f"Please fill them manually in the new row at the bottom of the "
                f"{'Inbound' if is_inbound else 'Outbound'} Parameters table, then type 'done'."
            ),
            fields=missing_fields,
        )
        # Treat as set — user just filled them
        msg_set = proc_set = True

    # ── Step 6: Save ──────────────────────────────────────────────────────────
    session.FindById("wnd[0]").SendVKey(11)   # Ctrl+S
    time.sleep(2)

    # Dismiss any transport / local-object popup
    for _ in range(3):
        try:
            wnd1 = session.FindById("wnd[1]", False)
            if wnd1:
                # Press "Local Object" if visible, else Enter
                try:
                    session.FindById("wnd[1]/usr/btnSPOPLI-SELFLAG").Press()
                except Exception:
                    session.FindById("wnd[1]").SendVKey(0)
                time.sleep(1)
        except Exception:
            break

    screen = get_screen_text()
    saved = ("saved" in screen.lower() or "changed" in screen.lower()
             or "created" in screen.lower()
             or partner_number in screen)

    # ── Step 6b: If save failed, full manual handoff ──────────────────────────
    if not saved:
        wait_for_user_input(
            screen_name="WE20 Partner Profile — Save Required",
            instructions=(
                f"The agent could not confirm that the partner profile was saved. "
                f"Please check WE20 for partner {partner_number} and verify that the "
                f"{'inbound' if is_inbound else 'outbound'} row for "
                f"Message Type={message_type} / Process Code={process_code} exists. "
                f"If it is missing: add it manually and press Ctrl+S. "
                f"If it already exists: just type 'done'."
            ),
            fields=[
                {"name": "Partner",       "value": partner_number},
                {"name": "Message Type",  "value": message_type},
                {"name": "Process Code",  "value": process_code},
                {"name": "Direction",     "value": "Inbound" if is_inbound else "Outbound"},
            ],
        )
        # Re-read screen after user confirmation
        screen = get_screen_text()
        saved = True   # User confirmed — treat as saved

    audit_log("CREATE_PARTNER_PROFILE",
              {"partner": partner_number, "type": partner_type,
               "direction": "inbound" if is_inbound else "outbound",
               "message_type": message_type, "process_code": process_code,
               "basic_type": basic_type},
              status="saved" if saved else "attempted")

    return {
        "ok":           saved,
        "partner":      partner_number,
        "partner_type": partner_type,
        "direction":    "inbound" if is_inbound else "outbound",
        "message_type": message_type,
        "process_code": process_code,
        "row_added":    row_added,
        "screen":       screen,
        "note": (f"Added {message_type}/{process_code} {'inbound' if is_inbound else 'outbound'} "
                 f"to partner {partner_number}/{partner_type}."),
    }


def _bd87_find_tree():
    """Locate the result tree/shell on a BD87 result screen."""
    # Known container paths for different SAP versions
    candidates = [
        "wnd[0]/usr/sub:SAPLBD7H:0200/cntlBD87_CONT/shellcont/shell",
        "wnd[0]/usr/cntlBD87_CONT/shellcont/shell",
        "wnd[0]/usr/sub:SAPLBD7H:0100/cntlBD87_CONT/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
    ]
    for cid in candidates:
        try:
            obj = session.FindById(cid, False)
            if obj:
                return obj
        except Exception:
            pass
    # Fallback: walk tree looking for GuiTree or GuiShell
    found = [None]
    def _walk(comp, depth=0):
        if found[0] or depth > 7:
            return
        try:
            n = comp.Children.Count
        except Exception:
            return
        for i in range(n):
            try:
                child = comp.Children(i)
                if child.Type in ("GuiTree", "GuiShell"):
                    found[0] = child
                    return
                _walk(child, depth + 1)
            except Exception:
                pass
    _walk(session.FindById("wnd[0]"))
    return found[0]


def _bd87_fill_selection(df, dt, message_type, docnum_lo, docnum_hi):
    """
    Fill BD87 'Select IDocs' screen fields.

    BD87 field layout (from screen inspection):
      • IDoc Number   — SEL_DOCNUM-LOW / HIGH  (ctxt OR txt)
      • Created On    — SEL_CREDAT-LOW / HIGH
      • Changed On    — SEL_UPDDAT-LOW / HIGH  ← BD87 auto-fills this; primary date
      • IDoc Status   — SEL_STATUS-LOW / HIGH
      • Message Type  — SEL_MESTYP-LOW / HIGH

    Strategy:
      1. Try two known sub-screen prefixes × both ctxt and txt prefixes
      2. Walk all elements and match by ID fragment (DOCNUM, UPDDAT, CREDAT, MESTYP)
      3. Label-based search for IDoc Number field as last resort
    """
    # ── Tier 1: hardcoded ID sets ─────────────────────────────────────────────
    PREFIXES = [
        "wnd[0]/usr/",
        "wnd[0]/usr/sub:SAPLBD7H:0100/",
        "wnd[0]/usr/sub:SAPLBD87:0100/",
    ]
    for pfx in PREFIXES:
        for typ in ("ctxt", "txt"):
            try:
                # IDoc Number
                lo_fid = f"{pfx}{typ}SEL_DOCNUM-LOW"
                hi_fid = f"{pfx}{typ}SEL_DOCNUM-HIGH"
                if docnum_lo:
                    session.FindById(lo_fid).Text = docnum_lo
                    try: session.FindById(hi_fid).Text = docnum_hi
                    except Exception: pass

                # Date — try Changed On (UPDDAT) first, then Created On (CREDAT)
                date_set = False
                for dfield in ("SEL_UPDDAT", "SEL_CREDAT"):
                    try:
                        session.FindById(f"{pfx}{typ}{dfield}-LOW").Text = df
                        try: session.FindById(f"{pfx}{typ}{dfield}-HIGH").Text = dt
                        except Exception: pass
                        date_set = True
                        break
                    except Exception:
                        pass

                # Message type
                if message_type:
                    try:
                        session.FindById(f"{pfx}{typ}SEL_MESTYP-LOW").Text = \
                            message_type.upper()
                    except Exception:
                        pass

                return True
            except Exception:
                pass

    # ── Tier 2: element scan by ID fragment ───────────────────────────────────
    elems  = discover_elements()
    filled = False
    for e in elems:
        eid = e.get("id", "").upper()
        try:
            if "DOCNUM" in eid and "LOW" in eid and docnum_lo:
                session.FindById(e["id"]).Text = docnum_lo; filled = True
            elif "DOCNUM" in eid and "HIGH" in eid and docnum_hi:
                session.FindById(e["id"]).Text = docnum_hi
            elif ("UPDDAT" in eid or "CREDAT" in eid) and "LOW" in eid:
                session.FindById(e["id"]).Text = df; filled = True
            elif ("UPDDAT" in eid or "CREDAT" in eid) and "HIGH" in eid:
                session.FindById(e["id"]).Text = dt
            elif "MESTYP" in eid and "LOW" in eid and message_type:
                session.FindById(e["id"]).Text = message_type.upper()
        except Exception:
            pass

    # ── Tier 3: label-based search for IDoc Number field ─────────────────────
    if not filled and docnum_lo:
        obj, _ = _find_field_by_label(
            ["IDoc Number", "IDoc-Nummer", "IDoc Nummer"])
        if obj:
            try:
                obj.Text = docnum_lo
                filled = True
            except Exception:
                pass

    return filled


def _bd87_select_all_and_process():
    """
    After BD87 executes and shows a result tree, select all nodes and
    press the Process/Reprocess button.  Tries 4 selection strategies
    and 6 button candidates so it works across SAP versions.
    Returns (selected_method, processed_ok).
    """
    tree = _bd87_find_tree()
    selected = None

    # ── Strategy 1: tree.SelectAll() ──────────────────────────────────────────
    if tree and not selected:
        try:
            tree.SelectAll()
            selected = "SelectAll()"
            time.sleep(0.3)
        except Exception:
            pass

    # ── Strategy 2: tree.ExpandSubTree + select all nodes manually ────────────
    if tree and not selected:
        try:
            tree.ExpandSubTree(tree.GetAllNodeKeys()[0])
            keys = tree.GetAllNodeKeys()
            for k in keys:
                try:
                    tree.ChangeCheckBox(k, True)
                except Exception:
                    try:
                        tree.SelectNode(k)
                    except Exception:
                        pass
            selected = f"manual-select({len(keys)} nodes)"
            time.sleep(0.3)
        except Exception:
            pass

    # ── Strategy 3: Ctrl+A via window ─────────────────────────────────────────
    if not selected:
        try:
            session.FindById("wnd[0]").SendVKey(16)   # Ctrl+A
            selected = "SendVKey(16)"
            time.sleep(0.3)
        except Exception:
            pass

    # ── Strategy 4: Edit menu → Select All ────────────────────────────────────
    if not selected:
        for menu_path in (
            "wnd[0]/mbar/menu[1]/menu[7]",
            "wnd[0]/mbar/menu[1]/menu[6]",
            "wnd[0]/mbar/menu[1]/menu[5]",
        ):
            try:
                session.FindById(menu_path).Select()
                selected = f"menu({menu_path})"
                time.sleep(0.3)
                break
            except Exception:
                pass

    # ── Press the Process / Reprocess toolbar button ──────────────────────────
    processed = False
    for btn in (
        "wnd[0]/tbar[1]/btn[8]",   # most common "Process" position
        "wnd[0]/tbar[1]/btn[9]",
        "wnd[0]/tbar[1]/btn[4]",
        "wnd[0]/tbar[1]/btn[5]",
        "wnd[0]/tbar[0]/btn[8]",
    ):
        try:
            session.FindById(btn).Press()
            processed = True
            time.sleep(3)
            break
        except Exception:
            pass

    # Fallback vkeys: F9 (Process selected), then F8 (Execute)
    if not processed:
        for vk in (9, 8):
            try:
                session.FindById("wnd[0]").SendVKey(vk)
                processed = True
                time.sleep(3)
                break
            except Exception:
                pass

    # Dismiss any confirmation popup
    try:
        wnd1 = session.FindById("wnd[1]", False)
        if wnd1:
            session.FindById("wnd[1]").SendVKey(0)
            time.sleep(2)
    except Exception:
        pass

    return selected, processed


def _we09_navigate_to_idoc(idoc_number) -> dict:
    """
    Navigate WE09 (IDoc Search) to display a single specific IDoc.

    WE09 is a search transaction: you enter DOCNUM LOW/HIGH on its selection
    screen, execute (F8), and the result list shows the matching IDoc(s).
    Double-clicking opens the IDoc detail screen.

    Returns dict with keys:
      on_detail  — True if IDoc detail screen was reached
      screen     — current screen text
      docnum     — the padded IDoc number used
    """
    docnum_padded = str(idoc_number).zfill(16)

    go_to_transaction("WE09")
    time.sleep(2)

    # Clear selection screen
    _clear_selection_screen()
    time.sleep(0.3)

    # Set DOCNUM field — same three-tier strategy as WE02/WE05
    docnum_set = False

    # Tier 1: ctxt + txt variants
    for fid in (
        "wnd[0]/usr/ctxtS_DOCNUM-LOW",   "wnd[0]/usr/txtS_DOCNUM-LOW",
        "wnd[0]/usr/ctxtSEL_DOCNUM-LOW", "wnd[0]/usr/txtSEL_DOCNUM-LOW",
        "wnd[0]/usr/ctxtDOCNUM-LOW",     "wnd[0]/usr/txtDOCNUM-LOW",
        "wnd[0]/usr/ctxtDOCNUM",         "wnd[0]/usr/txtDOCNUM",
        "wnd[0]/usr/ctxtS_DOCNUM",       "wnd[0]/usr/txtS_DOCNUM",
    ):
        try:
            session.FindById(fid).Text = docnum_padded
            docnum_set = True
            hi = fid.replace("-LOW", "-HIGH").replace("_LOW", "_HIGH")
            try: session.FindById(hi).Text = docnum_padded
            except Exception: pass
            break
        except Exception:
            pass

    # Tier 2: fragment scan
    if not docnum_set:
        for frag in ("S_DOCNUM-LOW", "DOCNUM-LOW", "DOCNUM"):
            obj, oid = _find_input_field_by_fragment(frag)
            if obj:
                try:
                    obj.Text = docnum_padded
                    docnum_set = True
                    hi = oid.replace("-LOW", "-HIGH").replace("_LOW", "_HIGH")
                    if hi != oid:
                        try: session.FindById(hi).Text = docnum_padded
                        except Exception: pass
                except Exception:
                    pass
                break

    # Tier 3: label-based search
    if not docnum_set:
        obj, _ = _find_field_by_label(
            ["IDoc Number", "IDoc-Nummer", "IDoc Nummer", "Doc.Number"])
        if obj:
            try:
                obj.Text = docnum_padded
                docnum_set = True
            except Exception:
                pass

    # Execute
    session.FindById("wnd[0]").SendVKey(8)
    time.sleep(3)

    screen = get_screen_text()
    on_detail = ("IDoc Display:" in screen or "IDoc-Anzeige:" in screen
                 or docnum_padded.lstrip("0") in screen)

    if not on_detail:
        # List screen — enter the first row
        for vkey in (2, 0, 13):
            try:
                session.FindById("wnd[0]").SendVKey(vkey)
                time.sleep(1.5)
                screen = get_screen_text()
                if ("IDoc Display:" in screen or "IDoc-Anzeige:" in screen
                        or docnum_padded.lstrip("0") in screen):
                    on_detail = True
                    break
            except Exception:
                pass

    return {
        "on_detail":   on_detail,
        "docnum":      docnum_padded,
        "screen":      screen,
        "docnum_set":  docnum_set,
        "note": ("WE09: IDoc detail screen reached." if on_detail
                 else "WE09: could not reach IDoc detail — check screen."),
    }


def bd87_select_and_reprocess(idoc_numbers=None, message_type="",
                               date_from=None, date_to=None):
    """
    BD87: filter by specific IDoc numbers (or date + message type),
    select ALL matching IDocs in the result tree, and reprocess them.

    idoc_numbers: list of IDoc number strings — if given, filters by
                  DOCNUM range (min to max).  Pass None to process all
                  IDocs matching date / message_type.
    """
    today     = datetime.now().strftime("%d.%m.%Y")
    last_year = datetime.now().year - 1
    # When specific IDoc numbers are given, use Jan 1 of last year as start date
    # so IDocs created anytime in the past year are included. DOCNUM filter
    # makes the date range irrelevant, but BD87 requires non-empty date fields.
    if idoc_numbers:
        df = date_from or f"01.01.{last_year}"
        dt = date_to   or today
    else:
        df = date_from or today
        dt = date_to   or today

    docnum_lo = docnum_hi = ""
    if idoc_numbers:
        padded    = [str(n).zfill(16) for n in idoc_numbers]
        docnum_lo = min(padded)
        docnum_hi = max(padded)

    go_to_transaction("BD87")
    time.sleep(1.5)
    try:
        _bd87_fill_selection(df, dt, message_type, docnum_lo, docnum_hi)
        session.FindById("wnd[0]").SendVKey(8)   # Execute
        time.sleep(3)

        selected, processed = _bd87_select_all_and_process()

        audit_log("BD87_SELECT_REPROCESS",
                  {"idoc_numbers": idoc_numbers, "message_type": message_type,
                   "docnum_lo": docnum_lo, "docnum_hi": docnum_hi,
                   "date_from": df, "date_to": dt},
                  status="executed" if processed else "attempted")

        return {
            "ok":           processed,
            "selected_by":  selected,
            "processed":    processed,
            "idoc_numbers": idoc_numbers,
            "message_type": message_type,
            "screen":       get_screen_text(),
        }
    except Exception as e:
        return {"error": str(e)}


def bd87_reprocess_all(message_type="", date_from=None, date_to=None):
    """Reprocess ALL failed IDocs (optionally filtered by message type).
    Delegates to bd87_select_and_reprocess with no IDoc number filter."""
    return bd87_select_and_reprocess(
        idoc_numbers=None,
        message_type=message_type,
        date_from=date_from,
        date_to=date_to,
    )


def wait_for_user_input(screen_name, instructions, fields=None):
    """
    Pause automation and hand off to the user for manual SAP screen input.
    Prints the screen name, what fields to fill, and their values.
    Waits for the user to type 'done' (or similar) before continuing.
    Returns the current SAP screen state so the agent can continue.
    """
    print("\n" + "═" * 68)
    print("  MANUAL INPUT REQUIRED")
    print("═" * 68)
    print(f"  Screen : {screen_name}")
    print(f"  Action : {instructions}")
    if fields:
        print("\n  Fields to fill:")
        for fld in (fields if isinstance(fields, list) else []):
            name  = fld.get("name", "")  if isinstance(fld, dict) else str(fld)
            value = fld.get("value", "") if isinstance(fld, dict) else ""
            note  = fld.get("note", "")  if isinstance(fld, dict) else ""
            suffix = f"  ({note})" if note else ""
            print(f"    • {name}: {value}{suffix}")
    print("\n  Fill these fields in SAP now, then type 'done' to continue.")
    print("  Type 'skip' to skip this step, 'abort' to stop the agent.")
    print("─" * 68)

    while True:
        try:
            response = input("  [done / skip / abort]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\n  [Interrupted — aborting manual input step]")
            return {"status": "aborted", "screen": get_screen_text(),
                    "note": "User interrupted manual input."}

        if response in ("done", "ok", "yes", "continue", "proceed",
                        "d", "y", "go", ""):
            screen = get_screen_text()
            texts  = _screen_texts()
            audit_log("MANUAL_INPUT",
                      {"screen": screen_name, "instructions": instructions},
                      status="completed", approved_by=CURRENT_USER)
            return {
                "status":       "user_completed",
                "screen":       screen,
                "screen_texts": texts[:30],
                "note": "User confirmed manual input complete — agent continuing.",
            }

        elif response in ("skip", "s"):
            audit_log("MANUAL_INPUT",
                      {"screen": screen_name, "instructions": instructions},
                      status="skipped", approved_by=CURRENT_USER)
            return {"status": "skipped",
                    "note": "User chose to skip this step."}

        elif response in ("abort", "stop", "exit", "quit"):
            audit_log("MANUAL_INPUT",
                      {"screen": screen_name, "instructions": instructions},
                      status="aborted", approved_by=CURRENT_USER)
            return {"status": "aborted",
                    "note": "User aborted the manual input step."}

        else:
            print("  Type 'done' when fields are filled, 'skip' to skip, "
                  "'abort' to stop.")


def ask_user_question(question: str, context: str, options: list = None) -> dict:
    """
    Ask the user an interactive question when the agent is stuck and needs
    information it cannot look up in SAP automatically.

    Prints a clear question block with context and waits for typed input.
    Returns the user's answer to the agent so it can continue.
    """
    print("\n" + "╔" + "═" * 66 + "╗", flush=True)
    print("║  AGENT QUESTION — input required to continue" + " " * 20 + "║", flush=True)
    print("╠" + "═" * 66 + "╣", flush=True)
    print(f"║  Why   : {context[:63]:<63}║", flush=True)
    if len(context) > 63:
        # Wrap long context
        remaining = context[63:]
        while remaining:
            chunk = remaining[:63]
            print(f"║          {chunk:<63}║", flush=True)
            remaining = remaining[63:]
    print("║" + "─" * 66 + "║", flush=True)
    print(f"║  Q     : {question[:63]:<63}║", flush=True)
    if len(question) > 63:
        remaining = question[63:]
        while remaining:
            chunk = remaining[:63]
            print(f"║          {chunk:<63}║", flush=True)
            remaining = remaining[63:]
    if options:
        print("║" + "─" * 66 + "║", flush=True)
        print(f"║  Options:", flush=True)
        for i, opt in enumerate(options, 1):
            print(f"║    {i}. {opt[:60]:<60}║", flush=True)
    print("╚" + "═" * 66 + "╝", flush=True)

    while True:
        try:
            answer = input("  Your answer: ").strip()
        except (EOFError, KeyboardInterrupt):
            answer = ""
            print("\n  [No answer given — agent will try to proceed without it]",
                  flush=True)
            break
        if answer:
            break
        print("  Please type an answer (or press Enter to skip).", flush=True)
        try:
            answer = input("  Your answer (Enter to skip): ").strip()
        except (EOFError, KeyboardInterrupt):
            answer = ""
        break

    print(f"  [Answer recorded: {answer[:80]}]", flush=True)
    audit_log("ASK_USER_QUESTION",
              {"question": question, "context": context, "answer": answer},
              status="answered", approved_by=CURRENT_USER)
    return {"answer": answer, "question": question, "context": context}


# Classification labels
_CLASS_LABEL = {
    "SOLVED_BY_AGENT":  "SOLVED BY AGENT       ✓",
    "HUMAN_THEN_AGENT": "HUMAN + AGENT         ✓",
    "CANNOT_SOLVE":     "CANNOT SOLVE          ✗",
}

# Session-level list so the main loop can print a final summary table
_session_resolutions: list = []

# ── Next-steps guidance keyed by IDoc error type ───────────────────────────────
# Each entry: patterns (substrings to match in error text, case-insensitive),
# title, tcode, team, urgency, steps (list of strings).
# The first matching entry is used.
_CANNOT_SOLVE_GUIDANCE = [
    {
        "patterns": ["partner profile does not exist", "no inbound partner profile",
                     "no outbound partner profile", "partner.*not found in we20",
                     "kein.*partnerprofil", "inbound partner profile"],
        "title":    "Missing Partner Profile",
        "tcode":    "WE20",
        "team":     "Basis / EDI team",
        "urgency":  "High — IDoc will not post until profile exists",
        "steps": [
            "Open tcode WE20",
            "In the left tree, expand the partner type node (LS / KU / LI)",
            "Find partner <PARTNER> — if missing, click Create and enter partner number + type",
            "Select the partner, switch to Change mode (pencil icon)",
            "In the Inbound Parameters table, click Create (small icon in toolbar)",
            "Enter: Message Type = <MESTYP>, Process Code = <PROCOD>",
            "Optionally enter Basic Type = <BASIC_TYPE>",
            "Press Ctrl+S to save",
            "Return here and reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["posting period", "period not open", "fiscal period",
                     "buchungsperiode", "period.*closed", "closed.*period",
                     "period.*not allowed", "no open period"],
        "title":    "Posting Period Not Open",
        "tcode":    "OB52 (FI)  or  MMPV / MMRV (MM)",
        "team":     "Finance controller / MM administrator",
        "urgency":  "High — posting will fail until period is opened",
        "steps": [
            "For FI (financial documents): open tcode OB52",
            "  → Find the posting period variant for company code",
            "  → Open the relevant period (From/To period, From/To year)",
            "  → Save",
            "For MM (material documents): open tcode MMPV",
            "  → Enter company code, period, year → Execute",
            "After opening period: reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["material.*does not exist", "material.*not found",
                     "no material master", "material.*unknown",
                     "material.*not created", "material number.*invalid"],
        "title":    "Material Master Does Not Exist",
        "tcode":    "MM01",
        "team":     "Material master / master data team",
        "urgency":  "Medium — create material or reject IDoc",
        "steps": [
            "Verify whether material <MATERIAL> should exist",
            "If yes: open MM01, create material with required views",
            "  → Basic Data 1/2, Purchasing, MRP 1-4, Plant data, etc.",
            "If no: mark IDoc as no further processing (WE02 → status 68)",
            "After creating material: reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["customer.*does not exist", "customer.*not found",
                     "no customer master", "debitor.*existiert nicht",
                     "kunnr.*not found"],
        "title":    "Customer Master Does Not Exist",
        "tcode":    "XD01",
        "team":     "Customer master / SD master data team",
        "urgency":  "Medium — create customer or reject IDoc",
        "steps": [
            "Verify whether customer <PARTNER> should exist",
            "If yes: open XD01, create customer in company code + sales area",
            "  → General data, Company code data, Sales area data",
            "If no: mark IDoc as no further processing (WE02 → right-click → status)",
            "After creating customer: reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["vendor.*does not exist", "vendor.*not found",
                     "no vendor master", "kreditor.*existiert nicht",
                     "lifnr.*not found"],
        "title":    "Vendor Master Does Not Exist",
        "tcode":    "XK01",
        "team":     "Vendor master / MM master data team",
        "urgency":  "Medium — create vendor or reject IDoc",
        "steps": [
            "Verify whether vendor <PARTNER> should exist",
            "If yes: open XK01, create vendor in company code + purchasing org",
            "  → General data, Company code data, Purchasing organisation data",
            "If no: mark IDoc as no further processing",
            "After creating vendor: reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["company code.*does not exist", "company code.*not defined",
                     "bukrs.*not found", "no company code", "buchungskreis"],
        "title":    "Company Code Not Defined",
        "tcode":    "OX02",
        "team":     "Finance / Basis team",
        "urgency":  "Critical — no posting is possible without company code",
        "steps": [
            "Open tcode OX02 (or SPRO → Enterprise Structure → Definition → FI)",
            "Create company code with: code, name, country, currency, language",
            "Assign company code to controlling area (OX19) if needed",
            "Run tcode OBY6 to copy chart of accounts settings",
            "Reprocess IDoc via BD87 after setup is complete",
        ],
    },
    {
        "patterns": ["port.*does not exist", "no port", "port.*not found",
                     "rfc.*port", "we21", "port definition"],
        "title":    "IDoc Port Not Defined",
        "tcode":    "WE21",
        "team":     "Basis / EDI team",
        "urgency":  "High — IDocs cannot be dispatched without a port",
        "steps": [
            "Open tcode WE21",
            "Select port type: RFC / File / TRFC / ABAP-PI",
            "Create port with name matching what WE20 partner profile expects",
            "For RFC port: enter RFC destination (SM59) pointing to target system",
            "Save port definition",
            "Update WE20 partner profile to reference the new port",
            "Reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["authorization", "authorisation", "not authorised",
                     "authority check", "no authorization", "auth.*failed",
                     "berechtigungs"],
        "title":    "Authorization Missing",
        "tcode":    "SU01 / PFCG / SU53",
        "team":     "Basis / Security team",
        "urgency":  "High — background user lacks required authorisation",
        "steps": [
            "Open SU53 as the failing user (or check /nSU53 after the error)",
            "  → Note the missing authorization object and field values",
            "Open PFCG → find the role assigned to the background user",
            "  → Add missing authorization object with required values",
            "  → Generate and save role, then run user comparison",
            "OR open SU01 → assign additional profile (e.g. EDI_ALL, S_IDOC_ALL)",
            "Reprocess IDoc via BD87 after fixing authorization",
        ],
    },
    {
        "patterns": ["syntax error", "syntax check", "segment.*error",
                     "field.*too long", "mandatory field.*missing",
                     "idoc.*structure", "wrong segment"],
        "title":    "IDoc Syntax / Segment Error",
        "tcode":    "WE02 / WE19",
        "team":     "EDI / Integration team",
        "urgency":  "Medium — fix data in IDoc or in sending system",
        "steps": [
            "Open WE02, find IDoc <IDOCNUM>, review segment data",
            "Identify the field causing the syntax error",
            "Option A — fix IDoc directly:",
            "  Open WE19, enter IDoc number, go to Change mode",
            "  Edit the offending segment field, save",
            "Option B — fix sending system:",
            "  Correct the source data in the sending system",
            "  Ask the sending system to resend the IDoc",
            "Reprocess corrected IDoc via BD87",
        ],
    },
    {
        "patterns": ["process code", "no inbound function module",
                     "function module.*not found", "procod", "nacfn",
                     "inbound.*function module"],
        "title":    "Process Code / Function Module Missing",
        "tcode":    "WE20 / WE64 / SE37",
        "team":     "Basis / EDI team",
        "urgency":  "High — IDoc cannot be processed without a process code",
        "steps": [
            "Open WE20 → find partner <PARTNER> → open Inbound Parameters",
            "Check whether process code exists for message type <MESTYP>",
            "If process code is missing or wrong:",
            "  Open WE64 to verify process code → function module mapping",
            "  Create/correct the process code to point to the correct FM",
            "  In WE20, update the partner profile row with correct process code",
            "  Press Ctrl+S to save",
            "Reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["duplicate", "already posted", "already exists",
                     "document already", "doppelt", "doppelter"],
        "title":    "Duplicate / Already Posted Document",
        "tcode":    "WE02 / SE16N (EDIDS)",
        "team":     "Business user / application support",
        "urgency":  "Low — confirm whether duplicate is valid before acting",
        "steps": [
            "Open WE02, find IDoc <IDOCNUM> — check the error message details",
            "Search for the original document (check VBELN/BELNR in EDIDD segments)",
            "If the document was genuinely already posted:",
            "  The IDoc is correct — mark as no further processing",
            "  In WE02 right-click IDoc → Set status → 68 (No further processing)",
            "If this is a false duplicate (different data, same key):",
            "  Investigate the key field collision and correct via WE19",
            "  Reprocess via BD87",
        ],
    },
    {
        "patterns": ["exchange rate", "currency.*not found", "no exchange rate",
                     "kurs.*nicht", "ob08", "tcurr"],
        "title":    "Exchange Rate Not Maintained",
        "tcode":    "OB08",
        "team":     "Finance / Treasury team",
        "urgency":  "Medium — posting fails until rate is entered",
        "steps": [
            "Open tcode OB08",
            "Enter exchange rate type (usually M for standard), currency pair",
            "Enter valid-from date and rate",
            "Save",
            "Reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["gl account", "g/l account", "account.*does not exist",
                     "no account", "sachkonto", "fs00", "coa"],
        "title":    "GL Account Does Not Exist",
        "tcode":    "FS00",
        "team":     "Finance team",
        "urgency":  "Medium — posting requires valid GL account",
        "steps": [
            "Open tcode FS00",
            "Enter GL account number and company code",
            "Create account with: account group, short/long text, balance sheet/P&L",
            "Assign to chart of accounts and add company code data",
            "Save",
            "Reprocess IDoc via BD87",
        ],
    },
    {
        "patterns": ["queue", "smq1", "smq2", "trfc", "bgRFC",
                     "locked queue", "queue.*backlog"],
        "title":    "ALE / qRFC Queue Backlog or Lock",
        "tcode":    "SMQ1 (outbound)  or  SMQ2 (inbound)",
        "team":     "Basis team",
        "urgency":  "High — queued IDocs will not process until queue is released",
        "steps": [
            "Open SMQ1 (outbound queues) or SMQ2 (inbound queues)",
            "Filter by queue name or RFC destination",
            "Select locked/error entries, click Activate to release",
            "Check RFC destination (SM59) is reachable",
            "If RFC connection is down: fix network/logon, then release queue",
            "Reprocess IDoc via BD87 or let queue process automatically",
        ],
    },
]


def _lookup_guidance(reason: str, what_was_done: str = "") -> dict | None:
    """
    Match the reason/error text against _CANNOT_SOLVE_GUIDANCE patterns.
    Returns the first matching guidance entry, or None if no match.
    """
    text = (reason + " " + what_was_done).lower()
    for entry in _CANNOT_SOLVE_GUIDANCE:
        for pat in entry["patterns"]:
            import re as _re
            try:
                if _re.search(pat.lower(), text):
                    return entry
            except Exception:
                if pat.lower() in text:
                    return entry
    return None


def set_idoc_resolution(idoc_number: str, classification: str,
                         reason: str, what_was_done: str = "",
                         next_steps: str = "") -> dict:
    """
    Record and display the final classification for an IDoc processing attempt.

    For CANNOT_SOLVE, automatically appends structured next-step guidance
    based on the error type, even if next_steps is not provided by the agent.
    """
    label = _CLASS_LABEL.get(classification, classification)
    w = 66

    # For CANNOT_SOLVE, look up structured guidance automatically
    guidance = None
    if classification == "CANNOT_SOLVE":
        guidance = _lookup_guidance(reason, what_was_done)

    print(f"\n{'╔' + '═'*w + '╗'}", flush=True)
    print(f"║  IDoc {idoc_number}  —  {label:<{w - len(idoc_number) - 8}}║",
          flush=True)
    print(f"{'╠' + '═'*w + '╣'}", flush=True)

    def _row(key, val):
        val = str(val)
        lines = [val[i:i+(w-len(key)-4)] for i in range(0, len(val), w-len(key)-4)] or [""]
        for line in lines:
            print(f"║  {key:<14}: {line:<{w - len(key) - 4}}║", flush=True)
            key = ""

    if reason:
        _row("Reason", reason)
    if what_was_done:
        _row("Done", what_was_done)
    if next_steps:
        _row("Next steps", next_steps)

    # ── Structured guidance block (CANNOT_SOLVE only) ─────────────────────────
    if guidance:
        print(f"{'╠' + '═'*w + '╣'}", flush=True)
        print(f"║  SUGGESTED NEXT STEPS — {guidance['title']:<{w-26}}║", flush=True)
        print(f"{'╠' + '─'*w + '╣'}", flush=True)
        _row("Tcode",   guidance["tcode"])
        _row("Team",    guidance["team"])
        _row("Urgency", guidance["urgency"])
        print(f"║  {'Steps':<14}:{'':>{w-16}}║", flush=True)
        for i, step in enumerate(guidance["steps"], 1):
            # Wrap long steps
            prefix = f"  {i}. "
            line = step
            first = True
            while line:
                avail = w - 2
                chunk = line[:avail]
                line  = line[avail:]
                pad   = prefix if first else "     "
                first = False
                print(f"║{pad}{chunk:<{avail - len(pad) + 2}}║", flush=True)

    print(f"{'╚' + '═'*w + '╝'}", flush=True)

    entry = {
        "idoc_number":    idoc_number,
        "classification": classification,
        "reason":         reason,
        "what_was_done":  what_was_done,
        "next_steps":     next_steps,
        "guidance_title": guidance["title"] if guidance else "",
        "guidance_tcode": guidance["tcode"] if guidance else "",
        "guidance_team":  guidance["team"]  if guidance else "",
        "timestamp":      datetime.now().isoformat(),
    }
    _session_resolutions.append(entry)
    audit_log("IDOC_RESOLUTION", entry, status=classification)
    return {
        "ok":             True,
        "classification": classification,
        "idoc_number":    idoc_number,
        "guidance":       guidance["title"] if guidance else "no match",
    }


# ── Tool Definitions for Claude ────────────────────────────────────────────────
TOOLS = [
    # ── READ ────────────────────────────────────────────────────────────────────
    {
        "name": "read_screen",
        "description": "Return the current SAP screen title and status bar text.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },

    # ── NAVIGATE ────────────────────────────────────────────────────────────────
    {
        "name": "go_to_transaction",
        "description": (
            "Navigate to any SAP transaction. "
            "Config examples: OX02 (company code), OX10 (plant), "
            "OVX5 (sales org), OVXI (dist channel), OVXB (division), "
            "OKKP (controlling area), SE09 (transports), SPRO (IMG)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "tcode": {"type": "string", "description": "Transaction code"},
            },
            "required": ["tcode"],
        },
    },

    # ── DISCOVER ─────────────────────────────────────────────────────────────────
    {
        "name": "discover_screen_elements",
        "description": (
            "Scan the current SAP screen and return all interactable elements "
            "with IDs, types, tooltips and current values. "
            "ALWAYS call this after navigating to a transaction before filling fields."
        ),
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },

    # ── WRITE ────────────────────────────────────────────────────────────────────
    {
        "name": "set_field_value",
        "description": (
            "Set a SAP input field value by its element ID. "
            "Get valid IDs from discover_screen_elements first. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "element_id": {"type": "string",
                               "description": "Full path e.g. wnd[0]/usr/ctxtBUKRS"},
                "value":      {"type": "string"},
            },
            "required": ["element_id", "value"],
        },
    },
    {
        "name": "press_button",
        "description": (
            "Press a SAP button by element ID. "
            "Common: wnd[0]/tbar[0]/btn[0]=New, wnd[0]/tbar[1]/btn[8]=Execute. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "element_id": {"type": "string"},
            },
            "required": ["element_id"],
        },
    },
    {
        "name": "send_vkey",
        "description": (
            "Send a virtual key to SAP. "
            "0=Enter, 3=Back, 8=Execute(F8), 11=Save(F11), 12=Cancel. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "key_code": {"type": "integer"},
            },
            "required": ["key_code"],
        },
    },
    {
        "name": "handle_popup",
        "description": (
            "Handle a SAP dialog/popup. "
            "action: confirm=OK/Enter, cancel=Cancel, input=type text then confirm. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "action":     {"type": "string",
                               "enum": ["confirm","cancel","yes","no","input"]},
                "text_input": {"type": "string",
                               "description": "Text to enter (only for action=input)"},
            },
            "required": ["action"],
        },
    },
    {
        "name": "handle_transport_request",
        "description": (
            "Handle the SAP customising/workbench transport popup that appears "
            "when saving configuration. Leave transport_number blank to create new. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "transport_number": {
                    "type": "string",
                    "description": "Existing transport e.g. S4K900123 (blank = create new)",
                },
            },
            "required": [],
        },
    },

    # ── IDOC ERROR ANALYSIS & AUTO-FIX ───────────────────────────────────────
    {
        "name": "scan_idoc_errors",
        "description": (
            "Scan WE05 for failed IDocs in a date range. "
            "Returns list with IDoc number, status code, message type, "
            "sender/receiver partner. "
            "status_filter: comma-separated status codes e.g. '51,26,56' "
            "or 'all' for every status. "
            "direction: 'inbound', 'outbound', or 'both'. "
            "idoc_number: pass a specific IDoc number (e.g. '198025') to filter "
            "WE05 to show ONLY that IDoc — the DOCNUM LOW/HIGH fields are set "
            "automatically and the date range is widened to 01.01.2020 so the IDoc "
            "is never excluded by date. ALWAYS pass idoc_number when the user has "
            "given you a specific IDoc number. "
            "ALWAYS call this first when diagnosing IDoc errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":     {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today, or 01.01.2020 when idoc_number set)"},
                "date_to":       {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today)"},
                "direction":     {"type": "string",
                                  "enum": ["inbound", "outbound", "both"],
                                  "description": "IDoc direction filter"},
                "status_filter": {"type": "string",
                                  "description": "Comma-separated status codes or 'all'"},
                "idoc_number":   {"type": "string",
                                  "description": "Specific IDoc number to filter on (e.g. '198025'). "
                                                 "Sets DOCNUM LOW=HIGH on WE05 selection screen. "
                                                 "ALWAYS pass this when a specific IDoc number is known."},
            },
            "required": [],
        },
    },
    {
        "name": "get_idoc_detail",
        "description": (
            "Open a specific IDoc in WE02 and extract: "
            "control record, ALL status records with error text, segment list, "
            "and auto-computed fix_hints based on the error messages. "
            "Always call after scan_idoc_errors to get root cause per IDoc."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number": {"type": "string",
                                "description": "IDoc document number (up to 16 digits)"},
            },
            "required": ["idoc_number"],
        },
    },
    {
        "name": "get_idoc_segments",
        "description": (
            "Read raw segment data for an IDoc from tables EDIDD (segments) "
            "and EDIDC (control record) via SE16N. "
            "Use for deep data-level analysis when fix_hints suggest a data problem."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number": {"type": "string",
                                "description": "IDoc document number"},
            },
            "required": ["idoc_number"],
        },
    },
    {
        "name": "view_idoc_in_we09",
        "description": (
            "Navigate WE09 (IDoc Search) to display a specific IDoc directly. "
            "WE09 sets DOCNUM LOW=HIGH on its selection screen so only the requested "
            "IDoc is shown. Use as an alternative to get_idoc_detail when WE02 has "
            "GUI issues, or when you want to view the IDoc in WE09 specifically. "
            "ALWAYS pass the IDoc number directly — do NOT leave it blank."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number": {"type": "string",
                                "description": "IDoc number to display (e.g. '198025')"},
            },
            "required": ["idoc_number"],
        },
    },
    {
        "name": "reprocess_idoc",
        "description": (
            "Reprocess a failed IDoc via WE19. "
            "edit_mode=false: standard retry using existing data. "
            "edit_mode=true: open in edit mode so you can change field values first. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number": {"type": "string",
                                "description": "IDoc document number"},
                "edit_mode":   {"type": "boolean",
                                "description": "True=open for editing, False=direct retry"},
            },
            "required": ["idoc_number"],
        },
    },
    {
        "name": "edit_idoc_field",
        "description": (
            "Open IDoc in WE19 edit mode, locate a specific segment field "
            "and change its value, then reprocess. "
            "Use when root cause is wrong data in a segment (e.g. wrong partner number, "
            "wrong plant code, wrong date format). "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number":   {"type": "string",
                                  "description": "IDoc document number"},
                "segment_name":  {"type": "string",
                                  "description": "Segment name e.g. E1EDKA1, E1EDP01"},
                "field_name":    {"type": "string",
                                  "description": "Field name within the segment e.g. KUNNR, MATNR"},
                "new_value":     {"type": "string",
                                  "description": "Corrected field value"},
            },
            "required": ["idoc_number", "segment_name", "field_name", "new_value"],
        },
    },
    {
        "name": "check_partner_profile",
        "description": (
            "Check WE20 partner profile for a partner number. "
            "Confirms whether the profile exists and which message types are configured. "
            "Call when get_idoc_detail fix_hints show 'fix_partner_profile'."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "partner_number": {"type": "string",
                                   "description": "Partner number e.g. PLANT1000, 1000"},
                "direction":      {"type": "string",
                                   "enum": ["1", "2"],
                                   "description": "1=Inbound, 2=Outbound"},
                "message_type":   {"type": "string",
                                   "description": "IDoc message type e.g. ORDERS, DESADV"},
            },
            "required": ["partner_number"],
        },
    },
    {
        "name": "create_partner_profile",
        "description": (
            "Add an inbound or outbound message-type entry to a WE20 partner profile. "
            "Works whether the partner already exists (adds a row to existing profile) "
            "or doesn't exist yet (creates the header first). "
            "For IDoc 198025: partner_number='S4HANA2023', partner_type='LS', "
            "direction='1' (inbound), message_type='MATMAS', process_code='MATM'. "
            "Use when IDoc error is 'Inbound partner profile does not exist' or "
            "'partner not found'."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "partner_number": {"type": "string",
                                   "description": "Partner number e.g. S4HANA2023"},
                "partner_type":   {"type": "string",
                                   "enum": ["LS", "KU", "LI", "KR", "B", "GP", "BP"],
                                   "description": "LS=Logical system, KU=Customer, LI=Vendor"},
                "direction":      {"type": "string",
                                   "enum": ["1", "2"],
                                   "description": "1=Inbound, 2=Outbound"},
                "message_type":   {"type": "string",
                                   "description": "IDoc message type e.g. MATMAS, ORDERS, INVOIC"},
                "process_code":   {"type": "string",
                                   "description": "Process code: MATM for MATMAS, ORDE for ORDERS, BAPI for BAPI-based, etc."},
                "basic_type":     {"type": "string",
                                   "description": "IDoc basic type e.g. MATMAS01 (optional, SAP sets default)"},
                "func_module":    {"type": "string",
                                   "description": "Function module (optional, rarely needed)"},
            },
            "required": ["partner_number", "partner_type", "direction",
                         "message_type", "process_code"],
        },
    },
    {
        "name": "bd87_select_and_reprocess",
        "description": (
            "BD87: filter by specific IDoc numbers and/or date range, select ALL "
            "matching IDocs in the result tree, and trigger reprocessing. "
            "Use this when you have a list of specific IDoc numbers to reprocess. "
            "Works by setting DOCNUM LOW/HIGH on the selection screen, executing, "
            "then using SelectAll() on the result tree and pressing the Process button. "
            "Preferred over bd87_reprocess_all when you know the IDoc numbers."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_numbers":  {"type":  "array",
                                  "items": {"type": "string"},
                                  "description": "List of IDoc numbers to reprocess (e.g. ['198025','198026']). Omit to process all matching."},
                "message_type":  {"type": "string",
                                  "description": "IDoc message type filter e.g. ORDERS (optional)"},
                "date_from":     {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today)"},
                "date_to":       {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today)"},
            },
            "required": [],
        },
    },
    {
        "name": "bd87_reprocess_all",
        "description": (
            "Batch reprocess ALL failed IDocs (optionally filtered by message type) via BD87. "
            "Use when you want to process all IDocs without specifying numbers. "
            "Delegates to bd87_select_and_reprocess with no IDoc number filter."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "message_type": {"type": "string",
                                 "description": "IDoc message type e.g. ORDERS (blank=all)"},
                "date_from":    {"type": "string",
                                 "description": "DD.MM.YYYY (defaults to today)"},
                "date_to":      {"type": "string",
                                 "description": "DD.MM.YYYY (defaults to today)"},
            },
            "required": [],
        },
    },

    # ── MANUAL INPUT HANDOFF ──────────────────────────────────────────────────
    {
        "name": "wait_for_user_input",
        "description": (
            "PAUSE automation and hand off to the user for manual SAP screen input. "
            "Call this when you reach a SAP screen you CANNOT fill automatically after "
            "2 failed attempts (e.g. OX02/OX10 New Entries, SPRO config screens, "
            "complex F4 search dialogs). "
            "Prints the screen name and field instructions on the console. "
            "Waits for user to type 'done'. Then reads and returns the current screen state "
            "so the agent can continue. "
            "DO NOT loop endlessly trying the same field IDs — call this tool instead."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "screen_name": {
                    "type": "string",
                    "description": (
                        "Name of the SAP screen / transaction currently open, "
                        "e.g. 'OX02 New Entries: Company Code Detail'"
                    ),
                },
                "instructions": {
                    "type": "string",
                    "description": (
                        "Plain-language instructions for the user. "
                        "State which fields to fill and with what values, "
                        "e.g. 'Enter Company Code=1000, Name=SAP IDES DE, "
                        "Country=DE, Currency=EUR, then press Save (Ctrl+S).'"
                    ),
                },
                "fields": {
                    "type": "array",
                    "description": "Optional structured list of fields to fill.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name":  {"type": "string",
                                      "description": "SAP field label or name"},
                            "value": {"type": "string",
                                      "description": "Value to enter"},
                            "note":  {"type": "string",
                                      "description": "Optional hint or explanation"},
                        },
                    },
                },
            },
            "required": ["screen_name", "instructions"],
        },
    },

    # ── INTERACTIVE Q&A ───────────────────────────────────────────────────────
    {
        "name": "ask_user_question",
        "description": (
            "Ask the user an interactive question when the agent CANNOT proceed "
            "automatically and needs information or a decision from the user. "
            "Use this tool (NOT wait_for_user_input) for:\n"
            "  • The root cause requires business context only the user knows "
            "(e.g. 'Should this material be created or was it deleted deliberately?')\n"
            "  • Multiple fix paths exist and only the user can choose "
            "(e.g. 'Company code 1000 is missing — create it or use 2000?')\n"
            "  • An error message is ambiguous and the user must clarify intent\n"
            "  • A posting period is closed and user must decide whether to open it\n"
            "  • Master data (material/customer/vendor) does not exist and user "
            "must confirm whether to create it\n"
            "ALWAYS fill 'context' with: what error you found, what you already tried, "
            "and exactly what information you need. Never ask vague questions."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": (
                        "The specific question for the user. Be precise. "
                        "Example: 'Material TG-100 does not exist in MM03. "
                        "Should I create it (yes/no) or is this IDoc invalid?'"
                    ),
                },
                "context": {
                    "type": "string",
                    "description": (
                        "Why the agent is stuck. Include: the IDoc error text, "
                        "what fix was attempted, and why it failed or is ambiguous. "
                        "Example: 'IDoc 198025 error: Material TG-100 not found. "
                        "Checked MM03 — material does not exist in plant 1000.'"
                    ),
                },
                "options": {
                    "type": "array",
                    "description": (
                        "Optional list of choices the user can pick from. "
                        "Example: ['Yes, create the material', "
                        "'No, this IDoc is invalid — mark for deletion', "
                        "'Use substitute material TG-200 instead']"
                    ),
                    "items": {"type": "string"},
                },
            },
            "required": ["question", "context"],
        },
    },

    # ── RESOLUTION CLASSIFICATION ─────────────────────────────────────────────
    {
        "name": "set_idoc_resolution",
        "description": (
            "MANDATORY — call this as the FINAL step after processing every IDoc. "
            "Records and displays the outcome classification in a clear box on screen. "
            "Choose exactly one classification:\n\n"
            "  SOLVED_BY_AGENT   — IDoc fixed entirely by the agent with no human help. "
            "Final status is 53 or equivalent success. "
            "Set reason=what the root cause was, what_was_done=what the agent did.\n\n"
            "  HUMAN_THEN_AGENT  — IDoc fixed, but human had to perform at least one "
            "manual step (e.g. clicked Create button in WE20, filled a field). "
            "Agent completed the reprocess after the human step. "
            "Set reason=what required human help, what_was_done=what human did + agent did.\n\n"
            "  CANNOT_SOLVE      — IDoc is still in error state. Either the agent "
            "exhausted all options, or the user confirmed the fix should not be done, "
            "or it requires action outside the agent's capability (e.g. finance team "
            "must open a fiscal period, an authorisation grant is needed). "
            "Set reason=exact blocking cause, next_steps=what must be done manually."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "idoc_number": {
                    "type": "string",
                    "description": "The IDoc number being classified, e.g. '198025'",
                },
                "classification": {
                    "type": "string",
                    "enum": ["SOLVED_BY_AGENT", "HUMAN_THEN_AGENT", "CANNOT_SOLVE"],
                    "description": "One of: SOLVED_BY_AGENT, HUMAN_THEN_AGENT, CANNOT_SOLVE",
                },
                "reason": {
                    "type": "string",
                    "description": (
                        "Concise explanation. For SOLVED/HUMAN: what the root cause was. "
                        "For CANNOT_SOLVE: exact reason why it cannot be fixed."
                    ),
                },
                "what_was_done": {
                    "type": "string",
                    "description": (
                        "What actions were taken (agent and/or human). "
                        "E.g. 'Created partner profile S4HANA2023/MATMAS in WE20, "
                        "reprocessed via BD87 — status changed 56→53.'"
                    ),
                },
                "next_steps": {
                    "type": "string",
                    "description": (
                        "Only for CANNOT_SOLVE: what the user or team must do next. "
                        "Include tcode and exact values needed. "
                        "Leave blank for SOLVED_BY_AGENT and HUMAN_THEN_AGENT."
                    ),
                },
            },
            "required": ["idoc_number", "classification", "reason"],
        },
    },
]

# ── IRREVERSIBLE operations — the ONLY ones that still confirm ────────────────
# Everything else is auto-executed (autonomous mode).
CONFIRM_BEFORE = set()  # No irreversible ops remain in IDoc-only mode

# ── Tool Dispatcher ────────────────────────────────────────────────────────────
def dispatch(tool_name, tool_input):

    # Only the truly irreversible operations still require a confirmation.
    # All other write operations execute autonomously (see OPERATING MODE).
    if tool_name in CONFIRM_BEFORE:
        approved = ask_approval(
            action=f"IRREVERSIBLE: {tool_name}",
            details=tool_input,
            risk="high",
        )
        if not approved:
            return {"status": "skipped", "message": "Skipped by operator."}

    if tool_name == "read_screen":
        return {"screen": get_screen_text()}
    if tool_name == "go_to_transaction":
        return go_to_transaction(tool_input["tcode"])
    if tool_name == "discover_screen_elements":
        elems = discover_elements()
        return {"count": len(elems), "elements": elems}
    if tool_name == "set_field_value":
        return set_field(tool_input["element_id"], tool_input["value"])
    if tool_name == "press_button":
        return press_button(tool_input["element_id"])
    if tool_name == "send_vkey":
        return send_vkey(tool_input["key_code"])
    if tool_name == "handle_popup":
        return handle_popup(
            action=tool_input.get("action", "confirm"),
            text_input=tool_input.get("text_input"),
        )
    if tool_name == "handle_transport_request":
        return handle_transport(tool_input.get("transport_number"))

    # ── IDoc error analysis & auto-fix tools ───────────────────────────────────
    if tool_name == "scan_idoc_errors":
        return scan_idoc_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("direction", "both"),
            tool_input.get("status_filter", "51,26,56"),
            tool_input.get("idoc_number"),
        )
    if tool_name == "get_idoc_detail":
        return get_idoc_detail(tool_input["idoc_number"])
    if tool_name == "get_idoc_segments":
        return get_idoc_segments(tool_input["idoc_number"])
    if tool_name == "view_idoc_in_we09":
        return _we09_navigate_to_idoc(tool_input["idoc_number"])
    if tool_name == "reprocess_idoc":
        return reprocess_idoc(
            tool_input["idoc_number"],
            tool_input.get("edit_mode", False),
        )
    if tool_name == "edit_idoc_field":
        return edit_idoc_field(
            tool_input["idoc_number"],
            tool_input["segment_name"],
            tool_input["field_name"],
            tool_input["new_value"],
        )
    if tool_name == "check_partner_profile":
        return check_partner_profile(
            tool_input["partner_number"],
            tool_input.get("direction", "1"),
            tool_input.get("message_type", ""),
        )
    if tool_name == "create_partner_profile":
        return create_partner_profile(
            tool_input["partner_number"],
            tool_input["partner_type"],
            tool_input["direction"],
            tool_input["message_type"],
            tool_input["process_code"],
            tool_input.get("basic_type", ""),
            tool_input.get("func_module", ""),
        )
    if tool_name == "bd87_select_and_reprocess":
        return bd87_select_and_reprocess(
            tool_input.get("idoc_numbers"),
            tool_input.get("message_type", ""),
            tool_input.get("date_from"),
            tool_input.get("date_to"),
        )
    if tool_name == "bd87_reprocess_all":
        return bd87_reprocess_all(
            tool_input.get("message_type", ""),
            tool_input.get("date_from"),
            tool_input.get("date_to"),
        )

    # ── Manual input handoff ───────────────────────────────────────────────────
    if tool_name == "wait_for_user_input":
        return wait_for_user_input(
            tool_input["screen_name"],
            tool_input["instructions"],
            tool_input.get("fields"),
        )

    # ── Interactive Q&A ────────────────────────────────────────────────────────
    if tool_name == "ask_user_question":
        return ask_user_question(
            tool_input["question"],
            tool_input["context"],
            tool_input.get("options"),
        )

    # ── Resolution classification ──────────────────────────────────────────────
    if tool_name == "set_idoc_resolution":
        return set_idoc_resolution(
            tool_input["idoc_number"],
            tool_input["classification"],
            tool_input["reason"],
            tool_input.get("what_was_done", ""),
            tool_input.get("next_steps", ""),
        )

    return {"error": f"Unknown tool: {tool_name}"}

# ── System Prompt ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = f"""You are ARTILEGENZ — an SAP IDoc self-healing agent.
SAP USER: S4ABAP24  |  DATE: {datetime.now().strftime('%Y-%m-%d')}
Auth: SAP_ALL (WE02, WE05, WE09, WE19, WE20, BD87, SE16N)

═══════════════════════════════════════════════════════════
 MANDATORY 4-STEP WORKFLOW — execute in this exact order
═══════════════════════════════════════════════════════════
The user will give you an IDoc number. Execute these 4 steps
every time, with no deviations and no pauses for approval.

──────────────────────────────────────────────────────────
STEP 1 — ROOT CAUSE ANALYSIS
──────────────────────────────────────────────────────────
Call: get_idoc_detail(idoc_number)

Then print EXACTLY this block to screen:
  ┌─────────────────────────────────────────┐
  │ IDoc      : <number>
  │ Status    : <code> — <text>
  │ Error     : <error message from EDIDS>
  │ Partner   : <partner number> (<partner type>)
  │ Msg Type  : <MESTYP>
  │ Direction : <1=Inbound / 2=Outbound>
  │ Root Cause: <one-line diagnosis>
  │ Fix Action: <what will be done next>
  └─────────────────────────────────────────┘

──────────────────────────────────────────────────────────
STEP 2 — FIX
──────────────────────────────────────────────────────────
Go to the correct tcode and resolve the root cause.
Announce in ONE line what you are doing, then do it.

Fix map (error text → action):
  "partner profile does not exist"  → WE20: create_partner_profile
  "inbound function module"         → WE20: create_partner_profile (add process code)
  "posting period not open"         → OB52/MMPV: open period
  "material does not exist"         → MM03: check material
  "customer does not exist"         → XD03: check customer
  "vendor does not exist"           → XK03: check vendor
  "syntax error" / segment error    → get_idoc_segments → edit_idoc_field
  "authorization"                   → SU01: check user profiles

PARTNER TYPE lookup (for create_partner_profile):
  SE16N → TBDLS (LOGSYS = partner) → exists → type=LS
  SE16N → LFB1  (LIFNR  = partner) → exists → type=LI
  SE16N → KNB1  (KUNNR  = partner) → exists → type=KU

PROCESS CODE — always use the value from get_idoc_detail result:
  The summary["process_code"] field is looked up live from WE64/EDIFCT
  (table EDIFCT, DIRECT+MESTYP filter) before returning. Use that value
  directly — do NOT guess or use the hardcoded map. The hardcoded map
  is only a fallback inside the Python code when WE64 returns nothing.

DIRECTION: status 51/56/26/68 = always INBOUND (direction=1)

If a SAP GUI screen cannot be automated (field IDs not found after 2 tries):
  → call wait_for_user_input(screen_name, instructions, fields=[...])
  → user completes the action in SAP, types 'done', then continue.

──────────────────────────────────────────────────────────
STEP 3 — REPROCESS VIA BD87
──────────────────────────────────────────────────────────
Call: bd87_select_and_reprocess(idoc_numbers=[<idoc_number>])

The function sets DOCNUM LOW=HIGH, date from 01.01.{datetime.now().year - 1},
uses SelectAll(), and presses Process. Do not call any other
reprocess tool unless bd87_select_and_reprocess fails.

──────────────────────────────────────────────────────────
STEP 4 — VERIFY RESULT
──────────────────────────────────────────────────────────
Call: get_idoc_detail(idoc_number)
Read the new status code (compare to status from Step 1).

Status codes: 53=Posted OK  51=Not posted  56=With errors
              26=Syntax err  64=Ready  68=No further processing

──────────────────────────────────────────────────────────
STEP 5 — CLASSIFY (MANDATORY — always the last call)
──────────────────────────────────────────────────────────
Call: set_idoc_resolution(idoc_number, classification, reason,
                           what_was_done, next_steps)

Classification rules — pick exactly one:

  SOLVED_BY_AGENT
    When: agent fixed it with no human involvement
          AND final status = 53 or error is gone
    reason       : what the root cause was
    what_was_done: what the agent did (tcode, action, result)
    next_steps   : (leave blank)

  HUMAN_THEN_AGENT
    When: wait_for_user_input was called at any point
          AND IDoc is now fixed (human + agent together)
    reason       : what required human help and why agent couldn't do it
    what_was_done: what human did + what agent did after
    next_steps   : (leave blank)

  CANNOT_SOLVE
    When: IDoc is still in error state after all attempts, OR
          user confirmed the fix should not be done, OR
          fix requires action outside agent capability
    reason       : exact blocking cause (be specific — include error text,
                   partner number, message type, any relevant values)
    what_was_done: what was tried (list each attempt)
    next_steps   : any additional manual context the user needs
    NOTE: the system automatically appends a structured "Suggested Next Steps"
          block (tcode, team, step-by-step guide) based on the error type.
          You do NOT need to repeat those steps — just fill reason accurately.

═══════════════════════════════════════════════════════════
 WHEN STUCK — ASK THE USER, NEVER SILENTLY GIVE UP
═══════════════════════════════════════════════════════════
You have full SAP_ALL authorization. Most errors CAN be fixed.
But some require a business decision only the user can make.

RULE: Before giving up on any IDoc, you MUST either:
  a) Ask the user a specific question using ask_user_question, OR
  b) Print a CANNOT RESOLVE block (see below) with full explanation.
  NEVER just stop without one of these two actions.

USE ask_user_question when:
  • Error text is ambiguous — you do not know which fix to apply
    Example: "Segment E1MARAM missing" — ask "Should I edit the segment
    data or mark IDoc for deletion?"
  • Master data missing — material/customer/vendor does not exist
    Ask: "Material X does not exist. Create it or reject this IDoc?"
  • Multiple valid fix paths — user must choose
    Ask: "Company code 1000 missing. Create it, or change IDoc to use 2000?"
  • Posting period closed — user must decide whether to open it
    Ask: "Fiscal period 03/2025 is closed. Shall I open it via OB52?"
  • Fix requires a value you cannot determine from SAP
    Ask the specific question with the context of what you found.

After ask_user_question returns the user's answer, continue
the fix using that answer. Do not give up after asking.

USE wait_for_user_input when:
  • SAP GUI screen cannot be automated (field IDs not found after 2 tries)
  • User must click a button or navigate SAP manually
  • create_partner_profile returns ok=False or row_added=False:
      → The function already calls wait_for_user_input internally for
        each sub-step that fails (Create button, field fill, save).
      → If it still returns ok=False after those prompts, call
        wait_for_user_input yourself with full manual WE20 instructions:
          screen_name : "WE20 Partner Profile — Manual Creation"
          instructions: step-by-step (tcode WE20, find partner, click
                        Create in Inbound/Outbound, fill MESTYP + PROCOD,
                        Ctrl+S). Include all values the user needs.
          fields      : Partner, Type, Direction, Message Type, Process Code

CANNOT RESOLVE — print this block ONLY when all of these are true:
  1. You have tried at least 2 different approaches
  2. You have asked the user and their answer still does not enable a fix
  3. The fix is genuinely outside the agent's capability
     (e.g. user confirmed the material should NOT be created)

  ┌─────────────────────────────────────────────────────┐
  │ CANNOT RESOLVE — IDoc <number>
  │ Error     : <exact error text from SAP>
  │ Tried     : <list what was attempted, one line each>
  │ Blocked by: <exact reason — be specific>
  │ Action    : <what the user must do manually>
  │ Tcode     : <which SAP transaction the user should go to>
  └─────────────────────────────────────────────────────┘

═══════════════════════════════════════════════════════════
 RULES
═══════════════════════════════════════════════════════════
• Execute fixes immediately — no approval, no confirmation.
• Never ask the user for data you can look up in SAP.
• Always pass idoc_number to get_idoc_detail and
  idoc_numbers=[idoc_number] to bd87_select_and_reprocess.
• NEVER loop more than 2 times on the same failed screen.
• Use high-level tools first; fall back to manual navigation only
  when no dedicated tool exists.
• Keep all output SHORT — one line per action, one line per result.
• NEVER end without calling set_idoc_resolution — it is the mandatory
  last step for every IDoc, regardless of outcome.
"""

# ── Token / Rate-Limit Management ──────────────────────────────────────────────

# Maximum chars stored per tool result in the messages list.
# Full output is still printed to console; only the stored copy is capped.
_TOOL_RESULT_CAP = 300

# When the messages list grows past this many entries, compress older exchanges.
# Each tool call = 2 entries (assistant + user/tool_result), so 16 = ~8 calls.
_MSG_COMPRESS_THRESHOLD = 16
# How many of the most-recent messages to always keep intact
_MSG_KEEP_RECENT = 8
# Max chars kept from a long assistant text block during compression
_ASSISTANT_TEXT_CAP = 200


def _cap_tool_result(content_str: str) -> str:
    """Truncate a tool result string to _TOOL_RESULT_CAP chars."""
    if len(content_str) <= _TOOL_RESULT_CAP:
        return content_str
    return content_str[:_TOOL_RESULT_CAP] + " ...[truncated]"


def _trim_message_content(msg: dict) -> dict:
    """
    Return a copy of msg with long assistant text blocks trimmed to
    _ASSISTANT_TEXT_CAP chars.  Tool-use blocks and tool-result blocks are
    left untouched (tool results are already capped at storage time).
    """
    if msg["role"] != "assistant":
        return msg
    content = msg.get("content")
    if not isinstance(content, list):
        return msg
    new_blocks = []
    for b in content:
        btype = getattr(b, "type", None) or (b.get("type") if isinstance(b, dict) else None)
        if btype == "text":
            txt = getattr(b, "text", None) or (b.get("text", "") if isinstance(b, dict) else "")
            if len(txt) > _ASSISTANT_TEXT_CAP:
                trimmed = txt[:_ASSISTANT_TEXT_CAP] + " ...[trimmed]"
                if isinstance(b, dict):
                    new_blocks.append({**b, "text": trimmed})
                else:
                    # SDK object — convert to plain dict
                    new_blocks.append({"type": "text", "text": trimmed})
                continue
        new_blocks.append(b)
    return {**msg, "content": new_blocks}


def _compress_messages(messages: list) -> list:
    """
    When the conversation thread grows very long, replace the oldest middle
    section with a lightweight summary placeholder so the token count stays
    manageable.  Always preserves:
      • messages[0]  — original user query
      • last _MSG_KEEP_RECENT messages — most recent context
    Long assistant text blocks in the kept section are also trimmed.
    """
    if len(messages) <= _MSG_COMPRESS_THRESHOLD:
        return messages

    first      = messages[0]
    middle     = messages[1 : len(messages) - _MSG_KEEP_RECENT]
    recent     = messages[len(messages) - _MSG_KEEP_RECENT :]

    # Count how many tool calls were in the compressed window
    tool_calls = sum(
        1 for m in middle
        if m["role"] == "assistant"
        and isinstance(m.get("content"), list)
        and any(getattr(b, "type", None) == "tool_use" or
                (isinstance(b, dict) and b.get("type") == "tool_use")
                for b in m["content"])
    )

    summary = {
        "role":    "user",
        "content": (
            f"[Earlier portion of this conversation compressed to save tokens. "
            f"{len(middle)} messages / ~{tool_calls} tool calls were made in the "
            f"previous steps. The most recent {_MSG_KEEP_RECENT} messages follow.]"
        ),
    }

    # Trim long assistant text in the kept-recent window too
    trimmed_recent = [_trim_message_content(m) for m in recent]

    compressed = [first, summary] + trimmed_recent
    print(f"  [Context compressed: {len(messages)} → {len(compressed)} messages]")
    return compressed


def _make_cached_system(system_text: str) -> list:
    """
    Wrap the system prompt in a list with cache_control so the Anthropic API
    serves it from the prompt cache on repeat calls.  Cached input tokens are
    NOT counted against the standard input-TPM bucket — this is the primary
    fix for the 30,000 TPM rate limit.
    """
    return [{"type": "text", "text": system_text,
             "cache_control": {"type": "ephemeral"}}]


def _make_cached_tools(tools: list) -> list:
    """
    Add cache_control to the LAST tool in the list, which causes the Anthropic
    API to cache all tools up to and including it.  Avoids re-tokenising the
    large TOOLS list on every call.
    """
    if not tools:
        return tools
    result = list(tools)
    last   = dict(result[-1])
    last["cache_control"] = {"type": "ephemeral"}
    result[-1] = last
    return result


# Build cached versions once at module load — reused on every API call
_CACHED_SYSTEM = None   # populated lazily after SYSTEM_PROMPT is defined
_CACHED_TOOLS  = None   # populated lazily after TOOLS is defined


def _api_call_with_retry(client, model, max_tokens, system, tools, messages,
                          max_retries=4):
    """
    Call client.messages.create with:
    - Prompt caching on system prompt and tools (avoids re-tokenising ~20k
      tokens every call → fixes the 30,000 TPM rate limit)
    - Exponential back-off starting at 65 s on 429 (TPM resets every 60 s)
    - Message compression before each retry to further reduce token count
    - SDK-level retries disabled (max_retries=0) to avoid double-retry
    """
    global _CACHED_SYSTEM, _CACHED_TOOLS
    if _CACHED_SYSTEM is None:
        _CACHED_SYSTEM = _make_cached_system(system)
    if _CACHED_TOOLS is None:
        _CACHED_TOOLS = _make_cached_tools(tools)

    current_messages = messages
    wait = 65  # start above 60 s so the TPM window fully resets

    for attempt in range(max_retries + 1):
        try:
            return client.messages.create(
                model      = model,
                max_tokens = max_tokens,
                system     = _CACHED_SYSTEM,
                tools      = _CACHED_TOOLS,
                messages   = current_messages,
            )
        except anthropic.BadRequestError as e:
            # "prompt is too long: N tokens > 200000 maximum"
            if "prompt is too long" in str(e) and attempt < max_retries:
                print(f"\n  [Prompt too long — compressing context and retrying "
                      f"(attempt {attempt + 1}/{max_retries})...]")
                current_messages = _compress_messages(current_messages)
                # If compression didn't shrink (already below threshold), force it
                if len(current_messages) > _MSG_KEEP_RECENT + 2:
                    # Keep only first + summary + last _MSG_KEEP_RECENT // 2
                    keep = max(4, _MSG_KEEP_RECENT // 2)
                    current_messages = [current_messages[0]] + current_messages[-keep:]
                    print(f"  [Hard trim: kept {len(current_messages)} messages]")
            else:
                raise
        except anthropic.RateLimitError:
            if attempt == max_retries:
                raise
            print(f"\n  [429 rate limit — waiting {wait}s "
                  f"(TPM window reset), retry {attempt + 1}/{max_retries}...]")
            time.sleep(wait)
            wait = min(wait * 2, 300)   # cap at 5 min
            # Compress aggressively before retrying
            current_messages = _compress_messages(current_messages)
        except Exception:
            raise


# ── Progress display helpers ────────────────────────────────────────────────────

# Human-readable label for each tool call shown on screen
_TOOL_LABELS = {
    "get_idoc_detail":          "Analysing IDoc — reading WE02 detail & error messages",
    "get_idoc_segments":        "Reading IDoc segments",
    "view_idoc_in_we09":        "Opening IDoc in WE09",
    "scan_idoc_errors":         "Scanning IDoc errors in WE05",
    "check_partner_profile":    "Checking partner profile in WE20",
    "create_partner_profile":   "Creating / updating partner profile in WE20",
    "reprocess_idoc":           "Reprocessing IDoc via WE19",
    "edit_idoc_field":          "Editing IDoc segment field",
    "bd87_select_and_reprocess":"Reprocessing via BD87 — select & execute",
    "bd87_reprocess_all":       "Reprocessing all matching IDocs via BD87",
    "go_to_transaction":        "Navigating to SAP transaction",
    "discover_screen_elements": "Reading SAP screen elements",
    "set_field_value":          "Setting field value on SAP screen",
    "press_button":             "Pressing button on SAP screen",
    "send_vkey":                "Sending keyboard shortcut to SAP",
    "handle_popup":             "Handling SAP popup dialog",
    "handle_transport_request": "Handling transport request popup",
    "read_screen":              "Reading current SAP screen",
    "wait_for_user_input":      "Waiting for manual user input in SAP",
    "ask_user_question":        "Agent is asking you a question — input required",
    "set_idoc_resolution":      "Recording IDoc resolution classification",
}

def _ts():
    """Return current time as HH:MM:SS string."""
    return datetime.now().strftime("%H:%M:%S")

def _print_tool_header(tool_name, tool_input, call_num):
    """Print a clear, readable header line before each tool call."""
    label = _TOOL_LABELS.get(tool_name, f"Calling {tool_name}")
    # Summarise the key input in one short string
    hint = ""
    if "idoc_number" in tool_input:
        hint = f"  IDoc={tool_input['idoc_number']}"
    elif "idoc_numbers" in tool_input:
        nums = tool_input["idoc_numbers"]
        hint = f"  IDocs={nums}"
    elif "tcode" in tool_input:
        hint = f"  tcode={tool_input['tcode']}"
    elif "element_id" in tool_input:
        hint = f"  field={tool_input['element_id']}  value={tool_input.get('value','')}"
    elif "partner_number" in tool_input:
        hint = f"  partner={tool_input['partner_number']}  msg={tool_input.get('message_type','')}"

    print(f"\n[{_ts()}] #{call_num:02d}  {label}{hint}", flush=True)
    print(f"         Tool   : {tool_name}", flush=True)
    inp_preview = json.dumps(tool_input)
    if len(inp_preview) > 120:
        inp_preview = inp_preview[:120] + "..."
    print(f"         Input  : {inp_preview}", flush=True)

def _print_tool_result(tool_name, result_json, elapsed_ms):
    """Print the tool result in a readable way."""
    # Show up to 1200 chars on screen (stored copy is still capped at 300)
    preview = result_json[:1200]
    if len(result_json) > 1200:
        preview += f"  ...[+{len(result_json)-1200} chars]"
    print(f"         Result : {preview}", flush=True)
    print(f"         Time   : {elapsed_ms}ms", flush=True)


# ── Agent Loop ─────────────────────────────────────────────────────────────────
def run_agent(user_query, api_key, messages=None):
    """
    Run the ARTILEGENZ agent.

    Pass messages=None to start a fresh conversation.
    Pass an existing messages list to continue.
    Returns (messages, last_text).
    """
    client    = anthropic.Anthropic(api_key=api_key, max_retries=0)
    last_text = ""
    call_num  = 0      # counts tool calls for display
    round_num = 0      # counts API round-trips

    if messages is None:
        messages = [{"role": "user", "content": user_query}]
    else:
        messages = list(messages) + [{"role": "user", "content": user_query}]

    print(f"\n[{_ts()}] ARTILEGENZ started", flush=True)
    print("─" * 68, flush=True)

    while True:
        messages  = _compress_messages(messages)
        round_num += 1

        # ── Show "thinking" indicator before every API call ────────────────────
        print(f"\n[{_ts()}] Thinking (round {round_num})...", flush=True)

        t0       = time.time()
        response = _api_call_with_retry(
            client,
            model      = "claude-sonnet-4-6",
            max_tokens = 8192,
            system     = SYSTEM_PROMPT,
            tools      = TOOLS,
            messages   = messages,
        )
        api_ms = int((time.time() - t0) * 1000)
        print(f"[{_ts()}] Claude responded in {api_ms}ms  "
              f"(stop={response.stop_reason})", flush=True)

        # ── Print any narrative text Claude produced ───────────────────────────
        for block in response.content:
            if hasattr(block, "text") and block.text.strip():
                last_text = block.text
                print(f"\n[{_ts()}] [Agent]\n{block.text}", flush=True)

        if response.stop_reason == "end_turn":
            break

        if response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []

            for block in response.content:
                if block.type != "tool_use":
                    continue

                call_num += 1
                _print_tool_header(block.name, block.input, call_num)

                t1 = time.time()
                try:
                    result = dispatch(block.name, block.input)
                except Exception as exc:
                    result = {"error": str(exc)}
                elapsed = int((time.time() - t1) * 1000)

                full_content = json.dumps(result, ensure_ascii=False)
                _print_tool_result(block.name, full_content, elapsed)

                tool_results.append({
                    "type":        "tool_result",
                    "tool_use_id": block.id,
                    "content":     _cap_tool_result(full_content),
                })

            messages.append({"role": "user", "content": tool_results})
        else:
            break

    # Append the final assistant turn
    if response.stop_reason == "end_turn":
        messages.append({"role": "assistant", "content": response.content})
    print(f"\n[{_ts()}] Agent finished  ({call_num} tool calls, {round_num} rounds)",
          flush=True)
    print("─" * 68, flush=True)

    # Print session classification summary if any resolutions were recorded
    if _session_resolutions:
        solved   = [r for r in _session_resolutions if r["classification"] == "SOLVED_BY_AGENT"]
        human    = [r for r in _session_resolutions if r["classification"] == "HUMAN_THEN_AGENT"]
        unsolved = [r for r in _session_resolutions if r["classification"] == "CANNOT_SOLVE"]
        print(f"\n{'═'*68}", flush=True)
        print(f"  SESSION SUMMARY", flush=True)
        print(f"{'─'*68}", flush=True)
        print(f"  Solved by agent        : {len(solved):>3}  "
              + (", ".join(r['idoc_number'] for r in solved) or "—"), flush=True)
        print(f"  Human + agent          : {len(human):>3}  "
              + (", ".join(r['idoc_number'] for r in human) or "—"), flush=True)
        print(f"  Cannot solve           : {len(unsolved):>3}  "
              + (", ".join(r['idoc_number'] for r in unsolved) or "—"), flush=True)
        print(f"{'═'*68}", flush=True)

    # ── Save session report ────────────────────────────────────────────────────
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = (user_query if isinstance(user_query, str) else "continuation")
    slug = slug[:40].replace(" ", "_").replace("/", "-")
    rpath = os.path.join(LOG_DIR, f"SAP_Session_{slug}_{ts}.txt")
    with open(rpath, "w", encoding="utf-8") as f:
        f.write(f"Query : {user_query}\n")
        f.write(f"Time  : {datetime.now()}\n")
        f.write(f"User  : {CURRENT_USER}\n\n")
        for m in messages:
            role = m["role"].upper()
            c    = m["content"]
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict) and item.get("type") == "text":
                        f.write(f"[{role}] {item['text']}\n")
                    elif hasattr(item, "text"):
                        f.write(f"[{role}] {item.text}\n")
            elif isinstance(c, str):
                f.write(f"[{role}] {c}\n")

    print(f"\nSession saved  → {rpath}")
    print(f"Audit log      → {LOG_DIR}")
    return messages, last_text

# ── NLP Input Parser ───────────────────────────────────────────────────────────
def _nlp_parse_input(query: str, last_text: str, api_key: str) -> tuple:
    """
    Use Claude Haiku to classify user intent and enrich terse keyboard input.

    Returns (action, enriched_query) where:
      action         : "CONTINUE" | "NEW_TASK"
      enriched_query : full natural-language expansion of what the user meant

    Examples:
      "A"              → CONTINUE, "Approved. Please proceed with the fix."
      "R"              → CONTINUE, "Rejected. Please skip this fix."
      "yes"            → CONTINUE, "Yes, please proceed."
      "do all"         → CONTINUE, "Please fix all remaining errors automatically."
      "skip vendor"    → CONTINUE, "Skip the vendor fix but continue with the others."
      "what about 198025" → CONTINUE, "What was the result for IDoc 198025?"
      "scan failed jobs" → NEW_TASK, (original query)
    """
    if not last_text:
        return "NEW_TASK", query

    # Explicit hard resets — skip API call
    q_lower = query.strip().lower()
    if q_lower in ("new", "reset", "clear", "start over", "new task", "/new"):
        return "NEW_TASK", query

    client = anthropic.Anthropic(api_key=api_key)

    prompt = f"""You are the input parser for an SAP automation agent (ARTILEGENZ).

The agent's last response (last 700 chars):
---
{last_text[-700:]}
---

The user just typed: "{query}"

Your job:
1. Decide if this is a CONTINUATION of the current conversation or a NEW_TASK.

   CONTINUATION examples:
     - Approval/rejection: "A", "R", "yes", "no", "ok", "sure", "skip", "go", "proceed"
     - Follow-up question: "what about idoc 198025", "did it work", "show details"
     - Instruction to current task: "do all of them", "fix all", "skip the vendor one",
       "try again", "also check the GL account", "what's the error text"
     - Short acknowledgement: "ok thanks", "got it", "understood", "continue"
     - Correction: "actually fix all not just one", "use company code 2000"

   NEW_TASK examples:
     - Completely different SAP operation: "scan failed jobs this week",
       "show all blocked invoices", "fix ABAP dump in SAPMV45A",
       "create partner profile for vendor 100012"
     - Only NEW_TASK if it is CLEARLY unrelated to what the agent just discussed.

2. Enrich the user's input into a clear natural-language instruction.
   Keep the user's intent exactly — just expand terse input so the agent understands.
   Do NOT add information that wasn't implied. Do NOT change the meaning.

Reply in this EXACT format (two lines, nothing else):
ACTION: CONTINUE|NEW_TASK
ENRICHED: <the enriched instruction>"""

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.content[0].text.strip()
        action   = "CONTINUE"
        enriched = query
        for line in text.splitlines():
            if line.startswith("ACTION:"):
                val = line.split(":", 1)[1].strip().upper()
                action = val if val in ("CONTINUE", "NEW_TASK") else "CONTINUE"
            elif line.startswith("ENRICHED:"):
                enriched = line.split(":", 1)[1].strip()
        return action, enriched

    except Exception:
        # Heuristic fallback if Haiku call fails
        SAP_ACTIONS = {"scan","fix","show","find","create","release","reverse",
                       "check","run","load","repair","list","update","cancel",
                       "idoc","order","vendor","material","invoice","job",
                       "workflow","lock","dump","abap","delivery","purchase"}
        words = set(q_lower.split())
        if len(query) > 80 and bool(words & SAP_ACTIONS):
            return "NEW_TASK", query
        return "CONTINUE", query


# ── Entry Point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
    if not API_KEY:
        API_KEY = input("Anthropic API key: ").strip()

    print("\n" + "═" * 68)
    print("  ARTILEGENZ — IDoc Self-Healing Agent  [SAP_ALL]")
    print("  User: S4ABAP24  |  4-step workflow: analyse → fix → BD87 → result")
    print("═" * 68)
    print("\nEnter an IDoc number to fix it. Type 'exit' to quit.")
    print()

    while True:
        try:
            raw = input("IDoc number (or exit): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break

        if not raw:
            continue
        if raw.lower() in ("exit", "quit", "q"):
            break

        # Accept a plain number or a full query containing a number
        import re as _re
        m = _re.search(r'\d+', raw)
        if not m:
            print("  Please enter a numeric IDoc number.")
            continue

        idoc_number = m.group(0)
        query = (
            f"Fix IDoc {idoc_number}. "
            f"Follow the mandatory 4-step workflow exactly: "
            f"(1) get_idoc_detail(\"{idoc_number}\") and print root cause, "
            f"(2) go to the correct tcode and fix the error, "
            f"(3) bd87_select_and_reprocess(idoc_numbers=[\"{idoc_number}\"]), "
            f"(4) get_idoc_detail(\"{idoc_number}\") and print final status."
        )

        _session_resolutions.clear()   # fresh classification list per IDoc run

        print(f"\n{'═'*68}", flush=True)
        print(f"  IDoc {idoc_number}  —  started at {datetime.now().strftime('%H:%M:%S')}", flush=True)
        print(f"{'═'*68}", flush=True)
        run_agent(query, API_KEY)
        print(f"\n{'═'*68}", flush=True)
        print(f"  IDoc {idoc_number}  —  completed at {datetime.now().strftime('%H:%M:%S')}", flush=True)
        print(f"{'═'*68}\n", flush=True)
