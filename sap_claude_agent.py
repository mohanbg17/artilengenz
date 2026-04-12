"""
ARTILEGENZ SAP Claude Agent v9.0
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
    """Walk current screen; return up to 100 interactable elements with IDs."""
    results = []
    def walk(comp, depth=0):
        if depth > 7:
            return
        try:
            n = comp.Children.Count
        except Exception:
            return
        for i in range(n):
            try:
                child = comp.Children(i)
                t = child.Type
                if t in ("GuiTextField","GuiCTextField","GuiComboBox",
                         "GuiRadioButton","GuiCheckBox","GuiButton","GuiTab"):
                    try:
                        results.append({
                            "id":      child.Id,
                            "type":    t,
                            "value":   child.Text,
                            "tooltip": child.Tooltip,
                        })
                    except Exception:
                        pass
                walk(child, depth + 1)
            except Exception:
                pass
    walk(session.FindById("wnd[0]"))
    return results[:100]

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

# ── Read Functions ─────────────────────────────────────────────────────────────
def read_sap_table(table_name, max_rows=200):
    go_to_transaction("SE16N")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = table_name
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1)
        try:
            session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = str(max_rows)
        except Exception:
            pass
        session.FindById("wnd[0]").SendVKey(8)   # Execute
        time.sleep(2)
        return {"table": table_name, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}

def maintain_table(table_name):
    """Open SM30 table maintenance for direct table writes."""
    go_to_transaction("SM30")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtVIEWNAME").Text = table_name
        # Choose Maintain (not Display)
        try:
            session.FindById("wnd[0]/usr/btnMAINTAIN").Press()
        except Exception:
            session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)
        return {"table": table_name, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}

def execute_abap(program_name=None, abap_code=None):
    """
    Execute an existing ABAP program via SA38, or
    create+run a temporary program via SE38.
    """
    if program_name:
        go_to_transaction("SA38")
        time.sleep(1)
        try:
            session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = program_name
            session.FindById("wnd[0]").SendVKey(8)  # Execute
            time.sleep(2)
            return {"program": program_name, "screen": get_screen_text()}
        except Exception as e:
            return {"error": str(e)}
    return {"note": "Provide program_name to execute existing ABAP program."}

def create_transport(description="ARTILEGENZ Change"):
    """Create a new Workbench or Customising transport request via SE09."""
    go_to_transaction("SE09")
    time.sleep(1)
    try:
        # Click Create button
        session.FindById("wnd[0]/tbar[1]/btn[13]").Press()
        time.sleep(1)
        popup = session.FindById("wnd[1]", False)
        if popup:
            # Set description if field available
            try:
                popup.FindById("usr/txtKO007-AS4TEXT").Text = description
            except Exception:
                pass
            popup.SendVKey(0)   # Enter/Create
            time.sleep(1.5)
        return {"screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}

def release_transport(transport_number):
    """Release a transport request via SE09."""
    go_to_transaction("SE09")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtKO007-TRKORR").Text = transport_number
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1)
        # Press Release button
        session.FindById("wnd[0]/tbar[1]/btn[19]").Press()
        time.sleep(1)
        handle_popup("confirm")
        return {"transport": transport_number, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}

# ── Table Data Loading ────────────────────────────────────────────────────────
def read_table_structure(table_name):
    """
    Read field definitions of any SAP table from SE11.
    Returns field names, types, lengths and key flags so Claude
    knows exactly which columns to populate when loading data.
    """
    table = table_name.strip().upper()
    go_to_transaction("SE11")
    time.sleep(1)
    try:
        # Select Database Table radio button
        for fid in ("wnd[0]/usr/radRB_DBTB", "wnd[0]/usr/rad_DBTB",
                    "wnd[0]/usr/radRB_TABL"):
            try:
                session.FindById(fid).Select()
                break
            except Exception:
                pass
        # Enter table name
        for fid in ("wnd[0]/usr/ctxtRS38M-TABNAME",
                    "wnd[0]/usr/ctxtOBJECT_NAME"):
            try:
                session.FindById(fid).Text = table
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(7)   # Display
        time.sleep(1.5)

        # Collect field rows from the Fields tab
        fields = []
        try:
            grid = session.FindById(
                "wnd[0]/usr/tabsTAB_STRIP/tabpFIELD/ssubSUB:SAPLSD11:2100"
                "/tblSAPLSD11TC_DD03", False)
            if grid is None:
                raise Exception("try ALV")
            rows = grid.RowCount
            for i in range(min(rows, 200)):
                row = {}
                for col in ["FIELDNAME","DATATYPE","LENG","DECIMALS",
                            "KEYFLAG","NOTNULL","DDTEXT"]:
                    try:
                        row[col] = grid.GetCellValue(i, col)
                    except Exception:
                        pass
                if row.get("FIELDNAME"):
                    fields.append(row)
        except Exception:
            # Fallback: read screen text
            pass

        screen = get_screen_text()
        return {"table": table, "fields": fields,
                "field_count": len(fields), "screen": screen}
    except Exception as e:
        return {"error": str(e), "table": table}


def load_data_to_table(table_name, records, mode="MODIFY"):
    """
    Load data into ANY SAP table by auto-generating an ABAP INSERT/MODIFY
    program, uploading it to SE38, activating and executing it.

    table_name : SAP table e.g. T001, VBAK, ZTABLE
    records    : list of dicts  [{"FIELD1": "VAL1", "FIELD2": "VAL2"}, ...]
    mode       : INSERT (new rows only) | MODIFY (upsert) | UPDATE (existing only)
    """
    table = table_name.strip().upper()
    # Safe program name (max 40 chars, Z prefix)
    safe = table.replace("/","_")[:8]
    prog = f"Z_ARTLGZ_LD_{safe}"
    ts   = datetime.now().strftime("%H%M%S")
    prog = f"{prog}_{ts}"[:30]

    # ── Generate ABAP source ────────────────────────────────────────────────
    lines = [
        f"*&{'─'*50}",
        f"*& ARTILEGENZ Data Load",
        f"*& Table  : {table}",
        f"*& Records: {len(records)}",
        f"*& Mode   : {mode}",
        f"*& Created: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        f"*& User   : {CURRENT_USER}",
        f"*&{'─'*50}",
        f"REPORT {prog.lower()}.",
        "",
        f"DATA: lt_data  TYPE STANDARD TABLE OF {table.lower()},",
        f"      ls_data  TYPE {table.lower()},",
        f"       lv_count TYPE i,",
        f"       lv_err   TYPE i.",
        "",
        "START-OF-SELECTION.",
        "",
    ]

    for i, rec in enumerate(records, 1):
        lines.append(f"  \"--- Record {i} ---")
        lines.append( "  CLEAR ls_data.")
        for field, value in rec.items():
            f_name = field.strip().lower()
            # Numeric types: no quotes; everything else: quotes
            val_str = str(value)
            if val_str.lstrip("-").replace(".","",1).isdigit():
                lines.append(f"  ls_data-{f_name} = {val_str}.")
            else:
                escaped = val_str.replace("'","''")
                lines.append(f"  ls_data-{f_name} = '{escaped}'.")
        lines.append( "  APPEND ls_data TO lt_data.")
        lines.append( "")

    _ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    lines += [
        f"  {mode} {table.lower()} FROM TABLE lt_data.",
        "  IF sy-subrc = 0.",
        "    lv_count = sy-dbcnt.",
        f"    WRITE: / 'ARTILEGENZ: ' && lv_count && ' record(s) loaded to {table}'.",
        "  ELSE.",
        "    lv_err = sy-subrc.",
        f"    WRITE: / 'ERROR loading {table}. SY-SUBRC:' && lv_err.",
        "  ENDIF.",
        f"  WRITE: / 'Done: {_ts} User:{CURRENT_USER}'.",
    ]

    source = "\n".join(lines)

    # ── Create, upload, activate, execute ───────────────────────────────────
    r_create = create_abap_program(prog, f"Load {table} data"[:40], "1", "$TMP")
    if not r_create.get("ok"):
        return {"error": "create_abap_program failed", "detail": r_create}

    r_upload = upload_abap_source(prog, source)
    if not r_upload.get("ok"):
        return {"error": "upload_abap_source failed", "detail": r_upload}

    r_syntax = check_abap_syntax(prog)
    r_act    = activate_abap_object(prog)
    r_run    = execute_abap(program_name=prog)

    audit_log("TABLE_DATA_LOAD",
              {"table": table, "records": len(records), "mode": mode,
               "program": prog},
              status="executed")

    return {
        "ok":           True,
        "table":        table,
        "program":      prog,
        "records":      len(records),
        "mode":         mode,
        "source_lines": len(lines),
        "syntax":       r_syntax.get("screen",""),
        "activated":    r_act.get("ok", False),
        "execution":    r_run.get("screen",""),
    }


def load_csv_to_table(csv_path, table_name, mode="MODIFY", delimiter=","):
    """
    Read a CSV file from the local Windows machine and load its
    contents into any SAP table using the ABAP load approach.

    csv_path  : Full local path  e.g. C:\\Users\\mohan\\Downloads\\data.csv
    table_name: Target SAP table e.g. T001, ZTABLE
    mode      : INSERT | MODIFY | UPDATE
    delimiter : , or ; or TAB
    """
    import csv as csv_mod
    try:
        delim = "\t" if delimiter.upper() in ("TAB", "\\T") else delimiter
        records = []
        with open(csv_path, "r", encoding="utf-8-sig", errors="replace") as fh:
            reader = csv_mod.DictReader(fh, delimiter=delim)
            for row in reader:
                records.append({k.strip(): v.strip() for k, v in row.items()})

        if not records:
            return {"error": "CSV empty or unreadable", "path": csv_path}

        return load_data_to_table(table_name, records, mode)

    except FileNotFoundError:
        return {"error": f"File not found: {csv_path}"}
    except Exception as e:
        return {"error": str(e), "csv_path": csv_path}


def sm30_load_entries(table_name, entries):
    """
    Load entries into a configuration/customising table via SM30.
    Best for tables that have a maintenance view (V_ prefix).
    entries: list of dicts with field:value pairs.
    """
    go_to_transaction("SM30")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtVIEWNAME").Text = table_name.strip().upper()
        # Click Maintain
        try:
            session.FindById("wnd[0]/usr/btnMAINTAIN").Press()
        except Exception:
            session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        results = []
        for entry in entries:
            # Click New Entries button
            try:
                session.FindById("wnd[0]/tbar[1]/btn[8]").Press()  # New entries
            except Exception:
                try:
                    session.FindById("wnd[0]/tbar[1]/btn[2]").Press()
                except Exception:
                    session.FindById("wnd[0]").SendVKey(4)
            time.sleep(1)

            # Discover fields and fill them
            elems = discover_elements()
            filled = []
            for field, value in entry.items():
                for elem in elems:
                    if field.lower() in elem.get("id","").lower() or \
                       field.lower() in elem.get("tooltip","").lower():
                        set_field(elem["id"], value)
                        filled.append(field)
                        break

            # Save row (Enter)
            session.FindById("wnd[0]").SendVKey(0)
            time.sleep(0.5)
            results.append({"entry": entry, "filled_fields": filled,
                            "screen": get_screen_text()})

        # Save all (F11)
        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1)
        handle_transport()

        return {"ok": True, "table": table_name,
                "entries_loaded": len(results), "detail": results}
    except Exception as e:
        return {"error": str(e)}


# ── ABAP Custom Code Creation ─────────────────────────────────────────────────
def create_abap_program(program_name, title, prog_type="1", package="$TMP"):
    """
    Create a new ABAP program in SE38.
    prog_type: 1=Executable, M=Module pool, F=Function group,
               S=Subroutine pool, I=Include, J=Interface pool
    package: $TMP = local (no transport), Z* = customer package
    After creation, call upload_abap_source to add the code.
    """
    prog = program_name.strip().upper()
    go_to_transaction("SE38")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = prog
        session.FindById("wnd[0]").SendVKey(5)   # F5 = Create
        time.sleep(1.5)

        # Dialog 1: set title + type
        wnd1 = session.FindById("wnd[1]", False)
        if wnd1:
            for fid in ("usr/txtRS38M-DBAPL", "usr/ctxtRS38M-DBAPL",
                        "usr/txtTITLE", "usr/ctxtTITLE"):
                try:
                    wnd1.FindById(fid).Text = title[:40]
                    break
                except Exception:
                    pass
            for fid in ("usr/ctxtRS38M-SUBC", "usr/radSUBC_1"):
                try:
                    wnd1.FindById(fid).Text = prog_type
                    break
                except Exception:
                    pass
            wnd1.SendVKey(0)
            time.sleep(1)

        # Dialog 2: package / transport
        wnd1 = session.FindById("wnd[1]", False)
        if wnd1:
            title2 = ""
            try:
                title2 = wnd1.Text
            except Exception:
                pass
            if "package" in title2.lower() or "paket" in title2.lower() \
                    or "object" in title2.lower():
                for fid in ("usr/ctxtRS38M-DEVCLASS", "usr/ctxtDEVCLASS",
                            "usr/ctxtOBJECT_PACKAGE"):
                    try:
                        wnd1.FindById(fid).Text = package
                        break
                    except Exception:
                        pass
                wnd1.SendVKey(0)
                time.sleep(1)

        # Transport popup (if package is not local)
        if package != "$TMP":
            handle_transport()

        return {"ok": True, "program": prog, "title": title,
                "type": prog_type, "package": package,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def create_function_module(fm_name, func_group, short_text, package="$TMP"):
    """
    Create a new Function Module in SE37.
    The function group must already exist (or create it first with SE80).
    After creation, call upload_abap_source with the FM source.
    """
    fm = fm_name.strip().upper()
    fg = func_group.strip().upper()
    go_to_transaction("SE37")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38L-NAME").Text = fm
        session.FindById("wnd[0]").SendVKey(5)   # Create
        time.sleep(1.5)

        wnd1 = session.FindById("wnd[1]", False)
        if wnd1:
            for fid in ("usr/ctxtRS38L-AREA", "usr/ctxtFUNCTION_GROUP",
                        "usr/ctxtAREA"):
                try:
                    wnd1.FindById(fid).Text = fg
                    break
                except Exception:
                    pass
            for fid in ("usr/txtRS38L-STEXT", "usr/txtSHORT_TEXT",
                        "usr/txtSHORT_DESCRIPT"):
                try:
                    wnd1.FindById(fid).Text = short_text[:40]
                    break
                except Exception:
                    pass
            wnd1.SendVKey(0)
            time.sleep(1.5)

        return {"ok": True, "fm": fm, "func_group": fg,
                "short_text": short_text, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def create_function_group(fg_name, short_text, package="$TMP"):
    """
    Create a new Function Group via SE80 (needed before creating FMs).
    """
    go_to_transaction("SE80")
    time.sleep(1.5)
    try:
        # Select "Function Group" from dropdown
        for fid in ("wnd[0]/usr/cmbTREE_MAIN-OCSLT",
                    "wnd[0]/usr/cmbOBJECT_SEL",
                    "wnd[0]/usr/cmbBROWSE_SEL"):
            try:
                cb = session.FindById(fid)
                cb.Key = "F"   # Function Group key
                break
            except Exception:
                pass
        # Enter name
        for fid in ("wnd[0]/usr/ctxtTREE_MAIN-OCNAME",
                    "wnd[0]/usr/ctxtOBJECT_NAME"):
            try:
                session.FindById(fid).Text = fg_name.strip().upper()
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1)
        # Handle "Create?" popup
        wnd1 = session.FindById("wnd[1]", False)
        if wnd1:
            for fid in ("usr/txtSHORT_TEXT", "usr/txtRS38L-STEXT"):
                try:
                    wnd1.FindById(fid).Text = short_text[:40]
                    break
                except Exception:
                    pass
            wnd1.SendVKey(0)
            time.sleep(1)
        return {"ok": True, "func_group": fg_name, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def add_fm_parameter(fm_name, param_name, param_type,
                     direction="import", type_ref="", optional=False):
    """
    Add an import/export/changing/tables parameter to an FM in SE37.
    direction: import | export | changing | tables | exceptions
    """
    fm = fm_name.strip().upper()
    go_to_transaction("SE37")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38L-NAME").Text = fm
        session.FindById("wnd[0]").SendVKey(6)   # Change
        time.sleep(1.5)

        # Navigate to the right tab
        tab_map = {
            "import":     "wnd[0]/usr/tabsTABSTRIP1/tabpTAB1",
            "export":     "wnd[0]/usr/tabsTABSTRIP1/tabpTAB2",
            "changing":   "wnd[0]/usr/tabsTABSTRIP1/tabpTAB3",
            "tables":     "wnd[0]/usr/tabsTABSTRIP1/tabpTAB4",
            "exceptions": "wnd[0]/usr/tabsTABSTRIP1/tabpTAB5",
        }
        try:
            session.FindById(tab_map.get(direction,
                tab_map["import"])).Select()
            time.sleep(0.5)
        except Exception:
            pass

        return {"ok": True, "fm": fm, "param": param_name,
                "direction": direction, "screen": get_screen_text(),
                "note": "Navigate to parameter tab and add row manually if needed."}
    except Exception as e:
        return {"error": str(e)}


def se11_create_data_element(element_name, short_text, domain=None,
                              built_in_type=None, length=None):
    """Create a Data Element in SE11."""
    go_to_transaction("SE11")
    time.sleep(1)
    try:
        # Select Data Element radio
        for fid in ("wnd[0]/usr/radRB_DTEL", "wnd[0]/usr/rad_DTEL"):
            try:
                session.FindById(fid).Select()
                break
            except Exception:
                pass
        # Enter name
        for fid in ("wnd[0]/usr/ctxtRS38M-DTEL", "wnd[0]/usr/ctxtOBJECT_NAME"):
            try:
                session.FindById(fid).Text = element_name.strip().upper()
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(5)   # Create
        time.sleep(1.5)
        return {"ok": True, "element": element_name, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def se11_create_table(table_name, short_text, package="$TMP"):
    """Create a transparent database table in SE11."""
    go_to_transaction("SE11")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/radRB_TABL", "wnd[0]/usr/rad_TABL"):
            try:
                session.FindById(fid).Select()
                break
            except Exception:
                pass
        for fid in ("wnd[0]/usr/ctxtRS38M-TABNAME",
                    "wnd[0]/usr/ctxtOBJECT_NAME"):
            try:
                session.FindById(fid).Text = table_name.strip().upper()
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(5)   # Create
        time.sleep(1.5)
        return {"ok": True, "table": table_name, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


# ── ABAP Dump Functions (ST22) ────────────────────────────────────────────────
def scan_st22_dumps(date_from=None, date_to=None):
    """List ABAP runtime errors from ST22."""
    go_to_transaction("ST22")
    time.sleep(1.5)
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today
    try:
        for fid in ("wnd[0]/usr/ctxtSEL_DATUM-LOW", "wnd[0]/usr/ctxtP_DATUM"):
            try:
                session.FindById(fid).Text = df
                break
            except Exception:
                pass
        for fid in ("wnd[0]/usr/ctxtSEL_DATUM-HIGH", "wnd[0]/usr/ctxtP_DATUM2"):
            try:
                session.FindById(fid).Text = dt
                break
            except Exception:
                pass
    except Exception:
        pass
    session.FindById("wnd[0]").SendVKey(8)   # Execute
    time.sleep(2)

    dumps = []
    # Try ALV grid first
    try:
        shell = session.FindById(
            "wnd[0]/usr/cntlST22_CONTAINER/shellcont/shell", False)
        if shell:
            rows = shell.RowCount
            for i in range(min(rows, 50)):
                row = {}
                for col in ["DATUM","UZEIT","UNAME","REPID","ERTYP","MANDT"]:
                    try:
                        row[col] = shell.GetCellValue(i, col)
                    except Exception:
                        pass
                if any(row.values()):
                    row["index"] = i
                    dumps.append(row)
    except Exception:
        pass

    return {"date": df, "dumps": dumps, "count": len(dumps),
            "screen": get_screen_text()}


def get_dump_detail(dump_index=0):
    """
    Open a specific dump from ST22 (by row index) and extract:
    error type, program, include, line number, error text, call stack, variables.
    """
    try:
        # Try ALV double-click
        try:
            shell = session.FindById(
                "wnd[0]/usr/cntlST22_CONTAINER/shellcont/shell", False)
            if shell:
                shell.SetCurrentCell(dump_index, "DATUM")
                shell.DoubleClickCurrentCell()
                time.sleep(2)
            else:
                raise Exception("no shell")
        except Exception:
            # Fallback: position cursor and press Enter
            session.FindById("wnd[0]").SendVKey(2)
            time.sleep(2)

        # Walk all text from dump detail screen
        texts = []
        def walk_text(comp, depth=0):
            if depth > 9:
                return
            try:
                n = comp.Children.Count
            except Exception:
                return
            for i in range(n):
                try:
                    child = comp.Children(i)
                    if child.Type in ("GuiTextField","GuiCTextField",
                                     "GuiLabel","GuiStatusbar","GuiTitlebar"):
                        try:
                            t = child.Text.strip()
                            if t:
                                texts.append(t)
                        except Exception:
                            pass
                    walk_text(child, depth + 1)
                except Exception:
                    pass

        walk_text(session.FindById("wnd[0]"))
        raw = "\n".join(texts[:300])

        # Parse key fields from raw text
        result = {"dump_index": dump_index, "raw": raw,
                  "screen": get_screen_text()}

        # Extract program name from dump text
        for line in texts:
            if "Program" in line or "REPID" in line:
                result["program_hint"] = line
                break

        return result

    except Exception as e:
        return {"error": str(e)}


def _clipboard_copy_from_sap():
    """Select all text in current SAP editor and copy to clipboard."""
    try:
        import win32api, win32con, win32clipboard
        # Ctrl+A
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(0x41, 0, 0, 0)
        win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.4)
        # Ctrl+C
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(0x43, 0, 0, 0)
        win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.5)
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
        win32clipboard.CloseClipboard()
        if isinstance(data, bytes):
            data = data.decode("latin-1", errors="replace")
        return data
    except Exception as e:
        return None


def _clipboard_paste_to_sap(text):
    """Put text on clipboard then Ctrl+A / Ctrl+V into current SAP editor."""
    try:
        import win32api, win32con, win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
        time.sleep(0.3)
        # Ctrl+A
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(0x41, 0, 0, 0)
        win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.3)
        # Ctrl+V
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(0x56, 0, 0, 0)
        win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.8)
        return True
    except Exception:
        return False


def download_abap_source(program_name):
    """
    Download ABAP source from SE38 to local file + return as string.
    Tries clipboard first, then menu download as fallback.
    """
    prog = program_name.strip().upper()
    filepath = os.path.join(TEMP_DIR, f"{prog}.abap")

    go_to_transaction("SE38")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = prog
        session.FindById("wnd[0]").SendVKey(7)   # F7 = Display
        time.sleep(1.5)
    except Exception as e:
        return {"error": f"SE38 navigate: {e}"}

    # Method 1 — clipboard
    source = _clipboard_copy_from_sap()
    if source and len(source) > 10:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(source)
        return {"program": prog, "source": source,
                "lines": len(source.splitlines()), "file": filepath,
                "method": "clipboard"}

    # Method 2 — SE38 menu download
    menu_candidates = [
        "wnd[0]/mbar/menu[0]/menu[7]",
        "wnd[0]/mbar/menu[0]/menu[6]",
        "wnd[0]/mbar/menu[4]/menu[9]/menu[0]",
    ]
    for mp in menu_candidates:
        try:
            session.FindById(mp).Select()
            time.sleep(1)
            wnd1 = session.FindById("wnd[1]", False)
            if wnd1:
                for fid in ("usr/ctxtDY_FILENAME","usr/ctxtFILENAME","usr/txtFILENAME"):
                    try:
                        wnd1.FindById(fid).Text = filepath
                        break
                    except Exception:
                        pass
                wnd1.SendVKey(0)
                time.sleep(1)
                handle_popup("confirm")
                time.sleep(1)
                if os.path.exists(filepath):
                    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                        source = f.read()
                    return {"program": prog, "source": source,
                            "lines": len(source.splitlines()), "file": filepath,
                            "method": "menu_download"}
        except Exception:
            continue

    return {"program": prog, "error": "Could not download source automatically. "
            "Please manually save from SE38 → Program → Download.",
            "file_would_be": filepath}


def upload_abap_source(program_name, source_code, is_new=False):
    """
    Upload fixed ABAP source to SE38.
    Saves to temp file then pastes via clipboard (Ctrl+A / Ctrl+V).
    """
    prog = program_name.strip().upper()
    filepath = os.path.join(TEMP_DIR, f"{prog}_fixed.abap")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(source_code)

    go_to_transaction("SE38")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = prog
        if is_new:
            session.FindById("wnd[0]").SendVKey(5)   # F5 = Create
            time.sleep(1.5)
            wnd1 = session.FindById("wnd[1]", False)
            if wnd1:
                wnd1.SendVKey(0)
                time.sleep(1)
        else:
            session.FindById("wnd[0]").SendVKey(6)   # F6 = Change
            time.sleep(1.5)
    except Exception as e:
        return {"error": f"SE38 open: {e}"}

    # Paste via clipboard
    ok = _clipboard_paste_to_sap(source_code)
    if ok:
        return {"ok": True, "program": prog, "method": "clipboard_paste",
                "lines": len(source_code.splitlines()),
                "file": filepath, "screen": get_screen_text()}

    # Fallback: menu upload
    menu_candidates = [
        "wnd[0]/mbar/menu[0]/menu[8]",
        "wnd[0]/mbar/menu[0]/menu[7]",
    ]
    for mp in menu_candidates:
        try:
            session.FindById(mp).Select()
            time.sleep(1)
            wnd1 = session.FindById("wnd[1]", False)
            if wnd1:
                for fid in ("usr/ctxtDY_FILENAME","usr/ctxtFILENAME","usr/txtFILENAME"):
                    try:
                        wnd1.FindById(fid).Text = filepath
                        break
                    except Exception:
                        pass
                wnd1.SendVKey(0)
                time.sleep(1)
                return {"ok": True, "program": prog, "method": "menu_upload",
                        "file": filepath, "screen": get_screen_text()}
        except Exception:
            continue

    return {"ok": False, "program": prog,
            "error": "Upload failed. Source saved locally: " + filepath}


def check_abap_syntax(program_name):
    """Syntax check in SE38 (must be in change mode first)."""
    prog = program_name.strip().upper()
    go_to_transaction("SE38")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = prog
        session.FindById("wnd[0]").SendVKey(6)   # Change
        time.sleep(1.5)
        # Syntax check button (toolbar btn[2]) or Ctrl+F2
        try:
            session.FindById("wnd[0]/tbar[1]/btn[2]").Press()
        except Exception:
            try:
                session.FindById("wnd[0]/mbar/menu[0]/menu[1]").Select()
            except Exception:
                pass
        time.sleep(2)
        return {"program": prog, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def activate_abap_object(program_name):
    """Activate ABAP program in SE38 after fix is uploaded."""
    prog = program_name.strip().upper()
    go_to_transaction("SE38")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = prog
        session.FindById("wnd[0]").SendVKey(6)   # Change
        time.sleep(1.5)
        # Activate: toolbar btn[3] or menu Program → Activate
        try:
            session.FindById("wnd[0]/tbar[1]/btn[3]").Press()
        except Exception:
            try:
                session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select()
            except Exception:
                pass
        time.sleep(2)
        # Handle any activation dialog (e.g. "inactive includes")
        try:
            wnd1 = session.FindById("wnd[1]", False)
            if wnd1:
                wnd1.SendVKey(0)
                time.sleep(1)
        except Exception:
            pass
        return {"ok": True, "program": prog, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


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
                     status_filter="51"):
    """
    Scan for IDoc errors using WE05 (IDoc list).
    Returns list of failed IDocs with number, status, message type, partner.
    status_filter: comma-separated status codes to search e.g. '51,26,56'
    direction: 'inbound', 'outbound', or 'both'
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("WE05")
    time.sleep(1.5)

    try:
        # Set date range
        for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-LOW",
                    "wnd[0]/usr/ctxtLOW_DATE"):
            try:
                session.FindById(fid).Text = df
                break
            except Exception:
                pass
        for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-HIGH",
                    "wnd[0]/usr/ctxtHIGH_DATE"):
            try:
                session.FindById(fid).Text = dt
                break
            except Exception:
                pass

        # Set direction (inbound/outbound)
        if direction == "inbound":
            for fid in ("wnd[0]/usr/radRB_DIRECT_1",
                        "wnd[0]/usr/radINBOUND"):
                try:
                    session.FindById(fid).Select()
                    break
                except Exception:
                    pass
        elif direction == "outbound":
            for fid in ("wnd[0]/usr/radRB_DIRECT_2",
                        "wnd[0]/usr/radOUTBOUND"):
                try:
                    session.FindById(fid).Select()
                    break
                except Exception:
                    pass

        # Execute
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        # Read ALV grid
        idocs = []
        try:
            shell = session.FindById(
                "wnd[0]/usr/cntlWE05_CONTAINER/shellcont/shell", False)
            if shell:
                rows = shell.RowCount
                for i in range(min(rows, 200)):
                    row = {}
                    for col in ["DOCNUM","STATUS","MANDT","DIRECT",
                                "MESTYP","MESCOD","MESFCT","SNDPRT",
                                "SNDPRN","RCVPRT","RCVPRN","CREDAT",
                                "CRETIM","UPDDAT","STATXT"]:
                        try:
                            row[col] = shell.GetCellValue(i, col)
                        except Exception:
                            pass
                    if row.get("DOCNUM"):
                        row["status_desc"] = IDOC_STATUS.get(
                            row.get("STATUS",""), "Unknown")
                        # Filter by status
                        if status_filter == "all" or \
                           row.get("STATUS","") in status_filter.split(","):
                            idocs.append(row)
        except Exception:
            pass

        return {
            "date_from":   df,
            "date_to":     dt,
            "idoc_errors": idocs,
            "count":       len(idocs),
            "screen":      get_screen_text(),
            "note": (f"Found {len(idocs)} IDocs matching status {status_filter}. "
                     "Call get_idoc_detail for each to analyse root cause."),
        }
    except Exception as e:
        return {"error": str(e)}


def get_idoc_detail(idoc_number):
    """
    Display a specific IDoc in WE02 and extract:
    - Control record (partner, message type, status, direction)
    - ALL status records with timestamps and error text
    - ALL segment data (key fields from each segment)
    Returns structured dict for root cause analysis.
    """
    go_to_transaction("WE02")
    time.sleep(1)
    try:
        # Enter IDoc number
        for fid in ("wnd[0]/usr/ctxtSEL_DOCNUM-LOW",
                    "wnd[0]/usr/ctxtDOCNUM"):
            try:
                session.FindById(fid).Text = str(idoc_number).zfill(16)
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(8)   # Execute
        time.sleep(2)

        detail = {
            "idoc_number": str(idoc_number),
            "status_records": [],
            "segments": [],
            "error_messages": [],
            "screen": get_screen_text(),
        }

        # Try to read tree/ALV structure
        try:
            shell = session.FindById(
                "wnd[0]/usr/cntlWE02_CONTAINER/shellcont/shell", False)
            if shell:
                rows = shell.RowCount
                for i in range(min(rows, 300)):
                    row = {}
                    for col in ["DOCNUM","STATUS","LOGDAT","LOGTIM",
                                "STAMQU","STATXT","SEGNAM","HLEVEL",
                                "DTINT2"]:
                        try:
                            row[col] = shell.GetCellValue(i, col)
                        except Exception:
                            pass
                    if row.get("STATXT"):
                        detail["status_records"].append(row)
                        txt = row.get("STATXT","").lower()
                        if any(e in txt for e in ["error","fehler","not found",
                                                   "nicht","invalid","missing"]):
                            detail["error_messages"].append(row["STATXT"])
                    if row.get("SEGNAM"):
                        detail["segments"].append(row)
        except Exception:
            pass

        # Also walk all text elements for any missed error text
        texts = []
        def walk_t(comp, depth=0):
            if depth > 8:
                return
            try:
                n = comp.Children.Count
            except Exception:
                return
            for i in range(n):
                try:
                    child = comp.Children(i)
                    if child.Type in ("GuiTextField","GuiCTextField","GuiLabel"):
                        try:
                            t = child.Text.strip()
                            if t and len(t) > 5:
                                texts.append(t)
                        except Exception:
                            pass
                    walk_t(child, depth + 1)
                except Exception:
                    pass
        walk_t(session.FindById("wnd[0]"))
        detail["all_screen_text"] = "\n".join(texts[:200])

        # Determine fix hint from error messages
        fix_hints = []
        combined_error = " ".join(detail["error_messages"]).lower()
        combined_error += " " + detail["all_screen_text"].lower()
        for pattern, (fix_action, risk) in IDOC_FIX_MAP.items():
            if pattern in combined_error:
                fix_hints.append({
                    "pattern":    pattern,
                    "fix_action": fix_action,
                    "risk":       risk,
                })
        detail["fix_hints"] = fix_hints

        return detail

    except Exception as e:
        return {"error": str(e), "idoc_number": str(idoc_number)}


def get_idoc_segments(idoc_number):
    """
    Read raw segment data from IDoc via EDIDD table (SE16N).
    Returns all segment content for deep data-level analysis.
    """
    result = read_sap_table("EDIDD", max_rows=500)
    # Also read control record
    ctrl  = read_sap_table("EDIDC", max_rows=10)
    return {
        "idoc_number": str(idoc_number),
        "control_record": ctrl,
        "segment_data":   result,
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


def check_partner_profile(partner_number, direction="1", message_type=""):
    """
    Check WE20 partner profile for a specific partner.
    direction: '1'=inbound, '2'=outbound
    Returns whether profile exists and configured message types.
    """
    go_to_transaction("WE20")
    time.sleep(1.5)
    try:
        # Try to search for the partner
        for fid in ("wnd[0]/usr/ctxtWE20-PARNR",
                    "wnd[0]/usr/ctxtPARTNER_NO"):
            try:
                session.FindById(fid).Text = str(partner_number)
                break
            except Exception:
                pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(1.5)
        return {"partner": str(partner_number),
                "screen":  get_screen_text(),
                "elements": discover_elements()[:40]}
    except Exception as e:
        return {"error": str(e)}


def create_partner_profile(partner_number, partner_type, direction,
                            message_type, process_code, func_module=""):
    """
    Create or fix a partner profile entry in WE20.
    partner_type: LS=Logical system, KU=Customer, LI=Vendor
    direction: 1=Inbound, 2=Outbound
    process_code: e.g. ORDE, DELVRY, INVOIC, DESADV
    """
    go_to_transaction("WE20")
    time.sleep(1.5)
    try:
        # Click Create / New
        for fid in ("wnd[0]/tbar[1]/btn[8]",
                    "wnd[0]/tbar[0]/btn[3]"):
            try:
                session.FindById(fid).Press()
                time.sleep(1)
                break
            except Exception:
                pass

        elems = discover_elements()
        field_map = {
            "PARNR": partner_number,
            "PARVW": partner_type,
        }
        for field, val in field_map.items():
            for e in elems:
                if field in e.get("id","").upper():
                    set_field(e["id"], val)
                    break

        session.FindById("wnd[0]").SendVKey(0)   # Enter
        time.sleep(1)

        return {"ok": True, "partner": partner_number,
                "type": partner_type, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def bd87_reprocess_all(message_type="", date_from=None, date_to=None):
    """
    Reprocess ALL failed IDocs of a given message type via BD87.
    Leave message_type blank to reprocess all error IDocs.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("BD87")
    time.sleep(1.5)
    try:
        # Set date
        for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-LOW",
                    "wnd[0]/usr/ctxtLOW_DATE"):
            try:
                session.FindById(fid).Text = df
                break
            except Exception:
                pass
        for fid in ("wnd[0]/usr/ctxtSEL_CREDAT-HIGH",
                    "wnd[0]/usr/ctxtHIGH_DATE"):
            try:
                session.FindById(fid).Text = dt
                break
            except Exception:
                pass
        if message_type:
            for fid in ("wnd[0]/usr/ctxtSEL_MESTYP-LOW",
                        "wnd[0]/usr/ctxtMESTYP"):
                try:
                    session.FindById(fid).Text = message_type.upper()
                    break
                except Exception:
                    pass

        session.FindById("wnd[0]").SendVKey(8)   # Execute
        time.sleep(2.5)

        # Select all IDocs and trigger reprocessing
        session.FindById("wnd[0]").SendVKey(16)  # Select all
        time.sleep(0.5)

        # Click Execute/Reprocess
        try:
            session.FindById("wnd[0]/tbar[1]/btn[9]").Press()
            time.sleep(2)
        except Exception:
            session.FindById("wnd[0]").SendVKey(9)
            time.sleep(2)

        audit_log("IDOC_BATCH_REPROCESS",
                  {"message_type": message_type or "ALL",
                   "date_from": df, "date_to": dt},
                  status="executed")
        return {"ok": True, "message_type": message_type or "ALL",
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def get_sales_orders(date_from, date_to):
    go_to_transaction("VA05")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtSD_VBAK-AUDAT_LOW").Text  = date_from
        session.FindById("wnd[0]/usr/ctxtSD_VBAK-AUDAT_HIGH").Text = date_to
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        time.sleep(2)
    except Exception as e:
        return {"error": str(e)}
    return {"screen": get_screen_text()}

def get_document_flow(sales_order):
    go_to_transaction("VA03")
    time.sleep(1)
    try:
        session.FindById("wnd[0]/usr/ctxtVBAK-VBELN").Text = str(sales_order)
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1)
        # Environment → Document flow
        session.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select()
        time.sleep(1.5)
    except Exception as e:
        return {"error": str(e)}
    return {"order": sales_order, "screen": get_screen_text()}

# ── Tool Definitions for Claude ────────────────────────────────────────────────
TOOLS = [
    # ── READ ────────────────────────────────────────────────────────────────────
    {
        "name": "get_sales_orders",
        "description": "Run VA05 and retrieve sales orders between two dates.",
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
            },
            "required": ["date_from", "date_to"],
        },
    },
    {
        "name": "read_sap_table",
        "description": "Read any SAP table via SE16N. Use for T001, T001W, TVKO, VBAK, EDIDC, etc.",
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name": {"type": "string", "description": "e.g. T001, VBAK, EDIDC"},
                "max_rows":   {"type": "integer", "description": "Default 200"},
            },
            "required": ["table_name"],
        },
    },
    {
        "name": "get_document_flow",
        "description": "Read O2C document flow for a sales order from VA03.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string", "description": "10-digit SAP order number"},
            },
            "required": ["sales_order"],
        },
    },
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
    {
        "name": "maintain_table",
        "description": (
            "Open SM30 table maintenance view for direct table write access. "
            "SAP_ALL grants write to all tables: T001, T001W, TVKO, EDIDC, etc. "
            "Use discover_screen_elements after opening to find entry fields. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name": {
                    "type": "string",
                    "description": "Table or view name e.g. T001, V_001, T001W",
                },
            },
            "required": ["table_name"],
        },
    },
    {
        "name": "execute_abap_program",
        "description": (
            "Execute an existing ABAP report/program via SA38. "
            "SAP_ALL allows running any program. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {
                    "type": "string",
                    "description": "ABAP program name e.g. RGSPAR00, Z_MY_REPORT",
                },
            },
            "required": ["program_name"],
        },
    },
    {
        "name": "create_transport_request",
        "description": (
            "Create a new Workbench or Customising transport request via SE09. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "description": {
                    "type": "string",
                    "description": "Transport description e.g. 'ARTILEGENZ Org Structure'",
                },
            },
            "required": ["description"],
        },
    },
    {
        "name": "release_transport_request",
        "description": (
            "Release an existing transport request via SE09 so it can be imported. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "transport_number": {
                    "type": "string",
                    "description": "Transport number e.g. S4K900123",
                },
            },
            "required": ["transport_number"],
        },
    },

    # ── TABLE DATA LOADING ────────────────────────────────────────────────────
    {
        "name": "read_table_structure",
        "description": (
            "Read the field definitions (name, type, length, key flag) of any "
            "SAP database table from SE11. ALWAYS call this before load_data_to_table "
            "so you know the exact field names and types to populate."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name": {"type": "string",
                               "description": "e.g. T001, VBAK, ZTABLE"},
            },
            "required": ["table_name"],
        },
    },
    {
        "name": "load_data_to_table",
        "description": (
            "Load records into ANY SAP table by auto-generating, uploading and "
            "executing an ABAP INSERT/MODIFY program. "
            "Works for ALL tables: config (T001, T001W), master data, "
            "transaction data, and custom Z-tables. "
            "Call read_table_structure first to know the correct field names. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name": {
                    "type": "string",
                    "description": "Target SAP table name e.g. T001, LFA1, ZTABLE",
                },
                "records": {
                    "type": "array",
                    "description": (
                        "List of records to insert. Each record is a dict of "
                        "field_name: value pairs matching the table structure. "
                        "e.g. [{\"BUKRS\":\"1000\",\"BUTXT\":\"IDES AG\",\"WAERS\":\"EUR\"}]"
                    ),
                    "items": {"type": "object"},
                },
                "mode": {
                    "type": "string",
                    "enum": ["INSERT","MODIFY","UPDATE"],
                    "description": (
                        "INSERT=new rows only (fails if exists), "
                        "MODIFY=upsert insert+update (recommended), "
                        "UPDATE=update existing rows only"
                    ),
                },
            },
            "required": ["table_name", "records"],
        },
    },
    {
        "name": "load_csv_to_table",
        "description": (
            "Read a CSV file from the local Windows machine and load all rows "
            "into a SAP table. Column headers in the CSV must match SAP field names. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "csv_path": {
                    "type": "string",
                    "description": r"Full local path e.g. C:\Users\mohan\Downloads\data.csv",
                },
                "table_name": {"type": "string",
                               "description": "Target SAP table"},
                "mode": {
                    "type": "string",
                    "enum": ["INSERT","MODIFY","UPDATE"],
                    "description": "INSERT|MODIFY(upsert)|UPDATE",
                },
                "delimiter": {
                    "type": "string",
                    "description": "CSV delimiter: , or ; or TAB",
                },
            },
            "required": ["csv_path", "table_name"],
        },
    },
    {
        "name": "sm30_load_entries",
        "description": (
            "Load entries into a customising/config table via SM30 table maintenance. "
            "Best for tables that have a view (V_ prefix) like V_T001, V_TVKO. "
            "For large volumes use load_data_to_table instead. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name": {
                    "type": "string",
                    "description": "Table or view name e.g. V_T001, T001W",
                },
                "entries": {
                    "type": "array",
                    "description": "List of entry dicts with field:value pairs",
                    "items": {"type": "object"},
                },
            },
            "required": ["table_name", "entries"],
        },
    },

    # ── ABAP CUSTOM CODE CREATION ────────────────────────────────────────────
    {
        "name": "create_abap_program",
        "description": (
            "Create a NEW ABAP program/report object in SE38. "
            "Use for custom reports, module pools, subroutine pools. "
            "After creation, call upload_abap_source to insert the code, "
            "then check_abap_syntax and activate_abap_object. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {
                    "type": "string",
                    "description": "Name e.g. Z_ARTILEGENZ_REPORT (must start with Z or Y)",
                },
                "title": {
                    "type": "string",
                    "description": "Short description of the program (max 40 chars)",
                },
                "prog_type": {
                    "type": "string",
                    "enum": ["1","M","F","S","I","J"],
                    "description": (
                        "1=Executable report (default), M=Module pool, "
                        "F=Function group, S=Subroutine pool, I=Include"
                    ),
                },
                "package": {
                    "type": "string",
                    "description": "$TMP=local no transport, Z*=customer package",
                },
            },
            "required": ["program_name", "title"],
        },
    },
    {
        "name": "create_function_module",
        "description": (
            "Create a NEW Function Module in SE37. "
            "The function group must already exist. "
            "After creation go to Source Code tab and call upload_abap_source. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "fm_name": {
                    "type": "string",
                    "description": "FM name e.g. Z_ARTILEGENZ_CALC (must start with Z or Y)",
                },
                "func_group": {
                    "type": "string",
                    "description": "Existing function group name e.g. ZARTILEGENZ",
                },
                "short_text": {
                    "type": "string",
                    "description": "Short description (max 40 chars)",
                },
                "package": {
                    "type": "string",
                    "description": "$TMP=local, Z*=customer package",
                },
            },
            "required": ["fm_name", "func_group", "short_text"],
        },
    },
    {
        "name": "create_function_group",
        "description": (
            "Create a new Function Group in SE80 "
            "(required before creating function modules). "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "fg_name": {
                    "type": "string",
                    "description": "Function group name e.g. ZARTILEGENZ",
                },
                "short_text": {"type": "string"},
                "package":    {"type": "string", "description": "$TMP or Z-package"},
            },
            "required": ["fg_name", "short_text"],
        },
    },
    {
        "name": "create_data_element",
        "description": (
            "Create a Data Element in SE11 for custom fields/tables. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "element_name": {"type": "string",
                                 "description": "e.g. ZARTILEGENZ_STATUS"},
                "short_text":   {"type": "string"},
            },
            "required": ["element_name", "short_text"],
        },
    },
    {
        "name": "create_database_table",
        "description": (
            "Create a custom transparent database table in SE11. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "table_name":  {"type": "string",
                                "description": "e.g. ZARTILEGENZ_LOG"},
                "short_text":  {"type": "string"},
                "package":     {"type": "string"},
            },
            "required": ["table_name", "short_text"],
        },
    },

    # ── ABAP DUMP DEBUG & FIX ─────────────────────────────────────────────────
    {
        "name": "scan_abap_dumps",
        "description": (
            "Scan ST22 for ABAP runtime errors (dumps). "
            "Returns list of dumps with date, time, user, program, error type. "
            "Always call this first when asked to fix ABAP dumps."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string",
                              "description": "DD.MM.YYYY (defaults to today)"},
                "date_to":   {"type": "string",
                              "description": "DD.MM.YYYY (defaults to today)"},
            },
            "required": [],
        },
    },
    {
        "name": "get_dump_detail",
        "description": (
            "Open a specific ABAP dump from ST22 by its row index and extract: "
            "error type, program name, include, line number, error message, "
            "call stack, and variable values. Call scan_abap_dumps first to "
            "get the index."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "dump_index": {"type": "integer",
                               "description": "Row index from scan_abap_dumps (0-based)"},
            },
            "required": ["dump_index"],
        },
    },
    {
        "name": "download_abap_source",
        "description": (
            "Download the full ABAP source code of a program from SE38. "
            "Use after get_dump_detail identifies the failing program. "
            "Returns complete source code for analysis."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {"type": "string",
                                 "description": "ABAP program name e.g. Z_MY_PROG, SAPMV45A"},
            },
            "required": ["program_name"],
        },
    },
    {
        "name": "upload_abap_source",
        "description": (
            "Upload corrected ABAP source code to SE38, replacing the existing code. "
            "Use after generating the fix. Pastes via clipboard. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {"type": "string"},
                "source_code":  {"type": "string",
                                 "description": "Complete corrected ABAP source code"},
                "is_new":       {"type": "boolean",
                                 "description": "True if creating a new program"},
            },
            "required": ["program_name", "source_code"],
        },
    },
    {
        "name": "check_abap_syntax",
        "description": (
            "Run syntax check on an ABAP program in SE38. "
            "Call after upload_abap_source and before activation. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {"type": "string"},
            },
            "required": ["program_name"],
        },
    },
    {
        "name": "activate_abap_object",
        "description": (
            "Activate an ABAP program in SE38 after uploading the fix. "
            "Only call after check_abap_syntax returns no errors. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "program_name": {"type": "string"},
            },
            "required": ["program_name"],
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
            "ALWAYS call this first when diagnosing IDoc errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":     {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today)"},
                "date_to":       {"type": "string",
                                  "description": "DD.MM.YYYY (defaults to today)"},
                "direction":     {"type": "string",
                                  "enum": ["inbound", "outbound", "both"],
                                  "description": "IDoc direction filter"},
                "status_filter": {"type": "string",
                                  "description": "Comma-separated status codes or 'all'"},
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
            "Create or fix a partner profile entry in WE20. "
            "Use when IDoc fails with 'partner not found' or 'no partner agreement'. "
            "partner_type: LS=Logical system, KU=Customer, LI=Vendor. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "partner_number": {"type": "string",
                                   "description": "Partner number"},
                "partner_type":   {"type": "string",
                                   "enum": ["LS", "KU", "LI", "KR"],
                                   "description": "Partner type: LS=Logical system, KU=Customer, LI=Vendor"},
                "direction":      {"type": "string",
                                   "enum": ["1", "2"],
                                   "description": "1=Inbound, 2=Outbound"},
                "message_type":   {"type": "string",
                                   "description": "IDoc message type e.g. ORDERS, INVOIC"},
                "process_code":   {"type": "string",
                                   "description": "Process code e.g. ORDE, DELVRY, INVOIC"},
                "func_module":    {"type": "string",
                                   "description": "Function module (optional)"},
            },
            "required": ["partner_number", "partner_type", "direction",
                         "message_type", "process_code"],
        },
    },
    {
        "name": "bd87_reprocess_all",
        "description": (
            "Batch reprocess ALL failed IDocs of a message type via BD87. "
            "Leave message_type blank to reprocess all error IDocs. "
            "Use after fixing the root cause (e.g. partner profile, posting period) "
            "to reprocess the backlog in one shot. "
            "REQUIRES human approval."
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
]

# ── Write operations that need approval ───────────────────────────────────────
WRITE_TOOLS = {
    "set_field_value", "press_button", "send_vkey",
    "handle_popup", "handle_transport_request",
    "maintain_table", "execute_abap_program",
    "create_transport_request", "release_transport_request",
    "upload_abap_source", "check_abap_syntax", "activate_abap_object",
    "create_abap_program", "create_function_module", "create_function_group",
    "create_data_element", "create_database_table",
    "load_data_to_table", "load_csv_to_table", "sm30_load_entries",
    # IDoc write tools
    "reprocess_idoc", "edit_idoc_field",
    "create_partner_profile", "bd87_reprocess_all",
}

# ── Tool Dispatcher ────────────────────────────────────────────────────────────
def dispatch(tool_name, tool_input):

    # Gate all write operations behind human approval
    if tool_name in WRITE_TOOLS:
        approved = ask_approval(
            action=f"SAP Write: {tool_name}",
            details=tool_input,
            risk="high" if tool_name == "handle_transport_request" else "medium",
        )
        if not approved:
            return {"status": "rejected", "message": "Rejected by operator."}

    if tool_name == "get_sales_orders":
        return get_sales_orders(
            tool_input.get("date_from", "01.01.2025"),
            tool_input.get("date_to",   "31.12.2026"),
        )
    if tool_name == "read_sap_table":
        return read_sap_table(
            tool_input["table_name"],
            tool_input.get("max_rows", 200),
        )
    if tool_name == "get_document_flow":
        return get_document_flow(tool_input["sales_order"])
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
    if tool_name == "maintain_table":
        return maintain_table(tool_input["table_name"])
    if tool_name == "execute_abap_program":
        return execute_abap(program_name=tool_input.get("program_name"))
    if tool_name == "create_transport_request":
        return create_transport(tool_input.get("description", "ARTILEGENZ Change"))
    if tool_name == "release_transport_request":
        return release_transport(tool_input["transport_number"])

    # ── ABAP dump debug tools ──────────────────────────────────────────────────
    if tool_name == "scan_abap_dumps":
        return scan_st22_dumps(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
        )
    if tool_name == "get_dump_detail":
        return get_dump_detail(tool_input.get("dump_index", 0))
    if tool_name == "download_abap_source":
        return download_abap_source(tool_input["program_name"])
    if tool_name == "upload_abap_source":
        return upload_abap_source(
            tool_input["program_name"],
            tool_input["source_code"],
            tool_input.get("is_new", False),
        )
    if tool_name == "check_abap_syntax":
        return check_abap_syntax(tool_input["program_name"])
    if tool_name == "activate_abap_object":
        return activate_abap_object(tool_input["program_name"])

    # ── Table data loading tools ───────────────────────────────────────────────
    if tool_name == "read_table_structure":
        return read_table_structure(tool_input["table_name"])
    if tool_name == "load_data_to_table":
        return load_data_to_table(
            tool_input["table_name"],
            tool_input["records"],
            tool_input.get("mode", "MODIFY"),
        )
    if tool_name == "load_csv_to_table":
        return load_csv_to_table(
            tool_input["csv_path"],
            tool_input["table_name"],
            tool_input.get("mode", "MODIFY"),
            tool_input.get("delimiter", ","),
        )
    if tool_name == "sm30_load_entries":
        return sm30_load_entries(
            tool_input["table_name"],
            tool_input["entries"],
        )

    # ── ABAP custom code creation tools ────────────────────────────────────────
    if tool_name == "create_abap_program":
        return create_abap_program(
            tool_input["program_name"],
            tool_input["title"],
            tool_input.get("prog_type", "1"),
            tool_input.get("package", "$TMP"),
        )
    if tool_name == "create_function_module":
        return create_function_module(
            tool_input["fm_name"],
            tool_input["func_group"],
            tool_input["short_text"],
            tool_input.get("package", "$TMP"),
        )
    if tool_name == "create_function_group":
        return create_function_group(
            tool_input["fg_name"],
            tool_input["short_text"],
            tool_input.get("package", "$TMP"),
        )
    if tool_name == "create_data_element":
        return se11_create_data_element(
            tool_input["element_name"],
            tool_input["short_text"],
        )
    if tool_name == "create_database_table":
        return se11_create_table(
            tool_input["table_name"],
            tool_input["short_text"],
            tool_input.get("package", "$TMP"),
        )

    # ── IDoc error analysis & auto-fix tools ───────────────────────────────────
    if tool_name == "scan_idoc_errors":
        return scan_idoc_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("direction", "both"),
            tool_input.get("status_filter", "51,26,56"),
        )
    if tool_name == "get_idoc_detail":
        return get_idoc_detail(tool_input["idoc_number"])
    if tool_name == "get_idoc_segments":
        return get_idoc_segments(tool_input["idoc_number"])
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
            tool_input.get("func_module", ""),
        )
    if tool_name == "bd87_reprocess_all":
        return bd87_reprocess_all(
            tool_input.get("message_type", ""),
            tool_input.get("date_from"),
            tool_input.get("date_to"),
        )

    return {"error": f"Unknown tool: {tool_name}"}

# ── System Prompt ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = f"""You are ARTILEGENZ — an expert SAP S/4 HANA consultant
with FULL SYSTEM ACCESS running as user S4ABAP24.

DATE/TIME    : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
SAP USER     : S4ABAP24
AUTHORIZATION: SAP_ALL + SAP_NEW (verified in SU01)
               → Unrestricted access to ALL transactions,
                 ALL tables, ALL configuration, ALL development,
                 ALL transport functions.

═══════════════════════════════════════════════════════════
 WHAT YOU CAN DO  (SAP_ALL — no restrictions)
═══════════════════════════════════════════════════════════

READ ANY DATA
  • Sales/Purchase/Finance docs    VA05, ME23N, FB03 …
  • Any table (read)               SE16N — T001, VBAK, EDIDC, BKPF …
  • Document flow                  VA03 → Environment → Document Flow
  • System logs                    SM21, ST22, SLG1

WRITE / CHANGE ANY DATA (human approval required per action)
  • Any SAP table directly         SM30 / SM31 (maintain_table tool)
  • Sales orders                   VA02
  • Purchase orders                ME22N
  • FI postings                    FB01, F-02
  • Master data                    MM02, XD02, XK02, CS02 …
  • IDocs                          WE19, BD87

CONFIGURATION (human approval required)
  • Company Code                   OX02  BUKRS / BUTXT / ORT01 / LAND1 / WAERS
  • Plant                          OX10  WERKS / NAME1 / LAND1 / ORT01
  • Sales Organisation             OVX5  VKORG / VTEXT / BUKRS / WAERS
  • Distribution Channel           OVXI  VTWEG / VTEXT
  • Division                       OVXB  SPART / VTEXT
  • Controlling Area               OKKP  KOKRS / BEZEI / BUKRS / WAERS
  • All other SPRO IMG             go_to_transaction → discover → fill → save
  • Credit Control Area            OB45
  • Chart of Accounts              OB13
  • Fiscal Year Variant            OB29
  • Posting Period Variant         OBBO

ABAP DEVELOPMENT (human approval required)
  • Create / edit programs         SE38 / SE80
  • Create function modules        SE37
  • Data dictionary                SE11
  • Execute programs               SA38 (execute_abap_program tool)

TRANSPORT MANAGEMENT (human approval required)
  • Create transport               SE09 (create_transport_request tool)
  • Release transport              SE09 (release_transport_request tool)
  • Import transport               STMS → go_to_transaction

SYSTEM ADMINISTRATION (human approval required)
  • User management                SU01, PFCG
  • System parameters              RZ10, RZ11
  • Process management             SM50, SM51
  • Lock management                SM12
  • Job scheduling                 SM36, SM37

═══════════════════════════════════════════════════════════
 STANDARD WRITE SEQUENCE (follow every time)
═══════════════════════════════════════════════════════════
  1. go_to_transaction(tcode)
  2. discover_screen_elements        ← always do this first
  3. set_field_value(id, value)      ← one field at a time, each approved
  4. send_vkey(0) or send_vkey(8)   ← Enter or Execute
  5. send_vkey(11)                   ← Save (F11)
  6. handle_transport_request        ← assign to transport (config changes)

═══════════════════════════════════════════════════════════
 IDES ORG STRUCTURE REFERENCE VALUES
═══════════════════════════════════════════════════════════
Company Code  : 1000  IDES AG               DE  EUR
Plant         : 1000  Werk Hamburg          DE
               1100  Werk Berlin           DE
Sales Org     : 1000  Deutschland           1000  EUR
               2000  Europe Export         1000  USD
Dist Channel  : 10    Endkundenverkauf
               12    Wiederverkäufer
Division      : 00    Prod.übergreifend
               01    Pumpen
Controlling   : 1000  Kostenrechnungskreis 1000
Credit Ctrl   : 1000
Chart/Accounts: INT
Fiscal Year   : K4
Purch Org     : 1000  IDES Deutschland

═══════════════════════════════════════════════════════════
 TABLE DATA LOADING — ALL TABLES
═══════════════════════════════════════════════════════════
You can load data into ANY SAP table. Three methods:

METHOD 1 — load_data_to_table  (recommended, all tables)
  Works by generating + running an ABAP INSERT/MODIFY program.
  ┌─────────────────────────────────────────────────────┐
  │ 1. read_table_structure(table) ← get field names    │
  │ 2. Ask user to provide the data if not given        │
  │ 3. load_data_to_table(table, records, mode)         │
  │    → auto-generates ABAP, uploads, activates, runs  │
  │    → returns success/error count                    │
  └─────────────────────────────────────────────────────┘

METHOD 2 — load_csv_to_table  (file upload)
  ┌─────────────────────────────────────────────────────┐
  │ CSV file headers MUST match SAP field names exactly │
  │ e.g.  BUKRS,BUTXT,WAERS,LAND1                      │
  │       1000,IDES AG,EUR,DE                           │
  │ load_csv_to_table(path, table, mode)                │
  └─────────────────────────────────────────────────────┘

METHOD 3 — sm30_load_entries  (config/customising tables)
  ┌─────────────────────────────────────────────────────┐
  │ Best for tables with SM30 maintenance views         │
  │ e.g. V_T001, V_TVKO, V_001                         │
  │ sm30_load_entries(table, [{field:value,...},...])    │
  └─────────────────────────────────────────────────────┘

MODE choices:
  INSERT → new rows only (fails on duplicate key)
  MODIFY → INSERT + UPDATE (upsert — use this by default)
  UPDATE → update existing rows only

COMMON TABLE LOADS:
  T001   Company codes      BUKRS,BUTXT,ORT01,LAND1,WAERS,SPRAS
  T001W  Plants             WERKS,NAME1,LAND1,ORT01,REGIO,ADRNR
  TVKO   Sales orgs         VKORG,VTEXT,BUKRS,WAERS,KNDNR
  LFA1   Vendor master      LIFNR,KTOKK,LAND1,NAME1,ORT01
  KNA1   Customer master    KUNNR,KTOKD,LAND1,NAME1,ORT01
  MARA   Material master    MATNR,MTART,MBRSH,MEINS,MATKL
  EKKO   Purchase orders    EBELN,BUKRS,BSTYP,AEDAT,LIFNR
  VBAK   Sales orders       VBELN,AUDAT,AUART,KUNNR,VKORG
  BKPF   FI documents       BUKRS,BELNR,GJAHR,BLART,BLDAT

ALWAYS:
  1. Call read_table_structure first to confirm field names
  2. Include MANDT (client) = sy-mandt if table has it
  3. Show the data to user for review before loading
  4. Use MODIFY mode by default to avoid duplicate key errors
  5. Audit log is written automatically after every load

═══════════════════════════════════════════════════════════
 ABAP CUSTOM CODE GENERATION WORKFLOW
═══════════════════════════════════════════════════════════
When asked to create a custom ABAP program or FM:

STEP 1  CLARIFY (ask if not provided)
        → What should the program DO?
        → Input parameters / selection screen fields?
        → Output: ALV report / file / IDoc / RFC call?
        → Which SAP tables to read/write?
        → Naming: Z_<PREFIX>_<DESCRIPTION>

STEP 2  GENERATE complete ABAP source in your response:
        Always include:
        ┌─────────────────────────────────────────────────┐
        │ *&---------------------------------------------*│
        │ *& Program: Z_ARTILEGENZ_XXXX                  *│
        │ *& Author : ARTILEGENZ AI                       *│
        │ *& Date   : {datetime.now().strftime('%d.%m.%Y')}             *│
        │ *& Purpose: <description>                       *│
        │ *&---------------------------------------------*│
        │ REPORT Z_ARTILEGENZ_XXXX.                       │
        │                                                 │
        │ *--- Data Declarations ---------------------------│
        │ TYPES: BEGIN OF ty_data, ...                    │
        │ DATA:  lt_data TYPE STANDARD TABLE OF ty_data,  │
        │        lv_var  TYPE string.                     │
        │                                                 │
        │ *--- Selection Screen ---------------------------│
        │ SELECT-OPTIONS: so_bukrs FOR t001-bukrs.        │
        │ PARAMETERS:     p_date   TYPE sy-datum.         │
        │                                                 │
        │ *--- Start of Selection -------------------------│
        │ START-OF-SELECTION.                             │
        │   PERFORM get_data.                             │
        │   PERFORM display_results.                      │
        │                                                 │
        │ *--- Subroutines --------------------------------│
        │ FORM get_data.                                  │
        │   TRY.                                          │
        │     SELECT ...                                  │
        │   CATCH cx_root INTO DATA(lx_exc).              │
        │     MESSAGE lx_exc->get_text() TYPE 'E'.        │
        │   ENDTRY.                                       │
        │ ENDFORM.                                        │
        └─────────────────────────────────────────────────┘

ABAP CODE STANDARDS (always follow):
  Naming     : Programs=Z_*, FMs=Z_*, Tables=Z*, Data Elements=Z*
  Variables  : lv_=local var, lt_=local table, ls_=local struct,
               gv_=global var, gt_=global table, lc_=constant
  Error handling: TRY...CATCH cx_root for all DB operations
  Messages   : MESSAGE 'text' TYPE 'S'/'E'/'W'/'I'/'X'
  ALV output : Use cl_salv_table for modern ALV grids
  Comments   : *--- Section ---, "-- inline comment
  No obsolete: avoid MOVE TO (use =), avoid WRITE with NEW-LINE in reports

CREATION SEQUENCE FOR NEW PROGRAM:
  1. create_abap_program(name, title, type, package)
  2. upload_abap_source(name, full_source_code)
  3. check_abap_syntax(name)           ← fix errors if any
  4. activate_abap_object(name)
  5. handle_transport_request()        ← if not $TMP

CREATION SEQUENCE FOR FUNCTION MODULE:
  1. [If new FG needed] create_function_group(fg, text)
  2. create_function_module(fm, fg, text)
  3. [Add parameters via SE37 tabs manually or via set_field_value]
  4. upload_abap_source(fm_name, source)
  5. check_abap_syntax + activate_abap_object
  6. handle_transport_request

CREATION SEQUENCE FOR CUSTOM TABLE:
  1. create_database_table(name, text)
  2. discover_screen_elements → add fields via set_field_value
  3. Set table category (Transparent), delivery class
  4. Save + activate + transport

═══════════════════════════════════════════════════════════
 ABAP DUMP DEBUG & FIX WORKFLOW  (ST22)
═══════════════════════════════════════════════════════════
When asked to fix/debug ABAP dumps, follow this exact sequence:

STEP 1  scan_abap_dumps(date_from, date_to)
        → Returns list of dumps with index, program, error type

STEP 2  get_dump_detail(dump_index)
        → Returns full error text, failing program/include/line,
          call stack, variable values at time of crash

STEP 3  download_abap_source(program_name)
        → Returns complete ABAP source code of the failing program

STEP 4  ANALYSE (no tool needed)
        → Read the dump detail + source code
        → Identify the exact line causing the error
        → Common errors and fixes:
          COMPUTE_BCD_OVERFLOW    → add CHECK / TRY CATCH around arithmetic
          CONVT_NO_NUMBER         → add CHECK sy-subrc / CONDENSE / validate input
          DATA_LENGTH_0           → check internal table is not empty before access
          GETWA_NOT_ASSIGNED      → check FIELD-SYMBOL is assigned before use
          MESSAGE_TYPE_X          → find the MESSAGE ... TYPE 'X' line and fix logic
          MOVE_CAST_ERROR         → check type compatibility before MOVE/CAST
          RAISE_EXCEPTION         → find unhandled RAISE and add TRY...CATCH
          CONNE_IMPORT_WRONG_VER  → program/transport version mismatch
        → Generate the corrected ABAP source in full

STEP 5  Show proposed fix to user:
        - Which line is wrong
        - What the original code does
        - What the fix changes and why
        → Wait for user approval

STEP 6  upload_abap_source(program_name, corrected_source)
        → REQUIRES approval → pastes corrected source into SE38

STEP 7  check_abap_syntax(program_name)
        → REQUIRES approval → verifies no syntax errors before activate

STEP 8  activate_abap_object(program_name)
        → REQUIRES approval → activates the fixed program

STEP 9  handle_transport_request()
        → REQUIRES approval → assigns fix to transport for migration

STEP 10 scan_abap_dumps(today, today)
        → Verify dump no longer appears

═══════════════════════════════════════════════════════════
 IDOC ROOT CAUSE ANALYSIS & AUTO-FIX WORKFLOW
═══════════════════════════════════════════════════════════
When asked to find/fix IDoc errors, follow this exact sequence:

STEP 1  scan_idoc_errors(date_from, date_to, direction, status_filter)
        → Returns list of failed IDocs with number, status, partner, message type
        → Default status_filter covers the most common error codes:
          51 = Application document not posted
          26 = Error during syntax check
          56 = IDoc with errors added
          64 = IDoc ready to be transferred (stuck)
          68 = Error — no further processing

STEP 2  get_idoc_detail(idoc_number) — for EACH failed IDoc
        → Returns error_messages[], status_records[], fix_hints[]
        → fix_hints are auto-computed from the error text; act on them:

        fix_hint → fix_action              → what to do
        ─────────────────────────────────────────────────────────
        fix_partner_profile   → check_partner_profile → if missing: create_partner_profile
        open_posting_period   → go to OB52 or MMPV/MMRV → maintain_table or set_field
        check_master_data     → read_sap_table(KNA1/LFA1/MARA) → fix or create master data
        fix_segment_data      → get_idoc_segments → identify bad field → edit_idoc_field
        fix_idoc_syntax       → get_idoc_segments → correct and edit_idoc_field
        check_authorization   → read_sap_table(UST12) → SU01 → add profile
        check_exchange_rate   → go to OB08 → maintain exchange rate
        check_gl_account      → go to FS00 → verify/create GL account
        check_tax_config      → go to FTXP → verify/create tax code
        check_fm_exists       → read_sap_table(TFDIR) → fix/create function module

STEP 3  PRESENT findings to user:
        - List all failed IDocs with status and error text
        - For each: state the root cause identified
        - Propose the specific fix action
        - ASK BEFORE FIXING: "I found X IDocs with error Y.
          Root cause: Z. Proposed fix: [action]. Approve? [A/R]"

STEP 4  Execute fix (only after user approves per IDoc or batch):
        PARTNER PROFILE MISSING:
          create_partner_profile(partner, type, direction, msg_type, process_code)
        WRONG SEGMENT DATA:
          edit_idoc_field(idoc, segment, field, corrected_value)
        STANDARD RETRY (data was fixed externally):
          reprocess_idoc(idoc_number, edit_mode=False)
        BATCH REPROCESS (after fixing root cause):
          bd87_reprocess_all(message_type, date_from, date_to)

STEP 5  Verify: scan_idoc_errors again to confirm count dropped to zero

COMMON ROOT CAUSES (memorize these patterns):
  "Partner ... not found"       → WE20 partner profile missing → create_partner_profile
  "No inbound function module"  → WE20 missing process code → create_partner_profile
  "Posting period ... not open" → OB52/MMPV/MMRV → open fiscal period
  "Company code ... not defined"→ OX02 → check/create company code
  "Material ... does not exist" → MM03/MM01 → check/create material
  "Customer ... does not exist" → XD03/XD01 → check/create customer
  "Vendor ... does not exist"   → XK03/XK01 → check/create vendor
  "Segment ... error"           → edit_idoc_field to correct segment data
  "Syntax error"                → get_idoc_segments → fix and edit_idoc_field
  "Authorization"               → SU01 → verify user has correct profiles

IDOC STATUS CODE QUICK REFERENCE:
  01=Generated  03=Dispatched  12=Dispatch OK  53=Posted OK
  26=Syntax err 51=Not posted  52=Partial post 56=With errors
  64=Ready      65=ALE error   68=No further   71=Edited copy

═══════════════════════════════════════════════════════════
 RULES
═══════════════════════════════════════════════════════════
• ALWAYS discover_screen_elements after every navigation.
• NEVER guess field IDs — discover first, then fill.
• Before creating any config object, read its table first
  (e.g. read T001 before OX02) to avoid duplicate key errors.
• After every config save, call handle_transport_request.
• If a field is protected/greyed, note it and move on.
• Confirm each step with read_screen to verify success.
• For SM30 table writes: navigate → discover → New Entry button
  → fill fields → Save → transport.
• For ABAP fixes: NEVER skip syntax check before activation.
• Always show the user what changed in the code before uploading.
• If source download fails, tell the user the program name so
  they can paste it manually.
"""

# ── Agent Loop ─────────────────────────────────────────────────────────────────
def run_agent(user_query, api_key):
    client   = anthropic.Anthropic(api_key=api_key)
    messages = [{"role": "user", "content": user_query}]

    print("\nARTILEGENZ agent running...\n")

    while True:
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=8192,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        # Show narrative text
        for block in response.content:
            if hasattr(block, "text") and block.text.strip():
                print(f"\n[Claude]\n{block.text}")

        if response.stop_reason == "end_turn":
            break

        if response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []

            for block in response.content:
                if block.type != "tool_use":
                    continue

                print(f"\n>> Tool : {block.name}")
                print(f"   Input: {json.dumps(block.input)[:200]}")

                try:
                    result = dispatch(block.name, block.input)
                except Exception as exc:
                    result = {"error": str(exc)}

                preview = json.dumps(result, ensure_ascii=False)[:300]
                print(f"   Out  : {preview}")

                tool_results.append({
                    "type":        "tool_result",
                    "tool_use_id": block.id,
                    "content":     json.dumps(result, ensure_ascii=False),
                })

            messages.append({"role": "user", "content": tool_results})
        else:
            break

    # ── Save session report ────────────────────────────────────────────────────
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug  = user_query[:40].replace(" ", "_").replace("/", "-")
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

# ── Entry Point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
    if not API_KEY:
        API_KEY = input("Anthropic API key: ").strip()

    print("\n" + "═" * 68)
    print("  ARTILEGENZ SAP Agent v9.0  —  User: S4ABAP24  [SAP_ALL]")
    print("  Authorization: FULL SYSTEM ACCESS")
    print("  All write operations require your approval first.")
    print("═" * 68)
    print("\nExample queries:")
    print('  "Scan all IDoc errors from today and tell me the root cause"')
    print('  "Find all failed IDocs from this week and fix them"')
    print('  "IDoc 0000000000123456 is failing — diagnose and propose a fix"')
    print('  "Reprocess all failed ORDERS IDocs from 01.04.2026 to 12.04.2026"')
    print('  "Check partner profile for partner 1000 inbound ORDERS"')
    print('  "Load these company codes into T001: 1000=IDES AG DE EUR, 2000=IDES US USD"')
    print('  "Load data from C:\\Users\\mohan\\Downloads\\vendors.csv into LFA1"')
    print('  "Show me the structure of table KNA1 then load 3 test customers"')
    print('  "Create an ABAP report that lists all open sales orders with ALV"')
    print('  "Create a function module to validate customer credit limit"')
    print('  "Create a custom Z-table to log all AI changes with timestamp"')
    print('  "Scan all ABAP dumps from today and fix them"')
    print('  "Fix the ABAP dump in program SAPMV45A"')
    print('  "Create the full IDES org structure with transports"')
    print('  "Show all sales orders from January 2025"')

    while True:
        query = input("\nQuery (or exit): ").strip()
        if query.lower() in ("exit", "quit", "q", ""):
            break
        run_agent(query, API_KEY)
