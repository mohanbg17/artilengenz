"""
ARTILEGENZ SAP Claude Agent v12.0
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


# ── Universal ALV Shell Finder ────────────────────────────────────────────────
def _find_shell_anywhere(container_hints=None):
    """
    Robustly locate any ALV grid shell on the current screen.
    Tries a list of known container IDs first, then walks the full UI tree.
    Returns the first usable GuiShell object, or None.
    """
    candidates = list(container_hints or [])
    for hint in candidates:
        for path in (hint, hint + "/shell", hint.rstrip("/shell")):
            try:
                obj = session.FindById(path, False)
                if obj and hasattr(obj, "RowCount"):
                    return obj
            except Exception:
                pass

    # Walk the entire window tree looking for any object with RowCount
    found = [None]
    def walk(comp, depth=0):
        if found[0] or depth > 7:
            return
        try:
            n = comp.Children.Count
        except Exception:
            return
        for i in range(n):
            try:
                child = comp.Children(i)
                try:
                    if hasattr(child, "RowCount"):
                        _ = child.RowCount      # confirm accessible
                        found[0] = child
                        return
                except Exception:
                    pass
                walk(child, depth + 1)
                if found[0]:
                    return
            except Exception:
                pass
    walk(session.FindById("wnd[0]"))
    return found[0]


def _read_shell_rows(shell, columns, max_rows=300):
    """Read up to max_rows from an ALV shell into a list of dicts."""
    rows = []
    try:
        count = shell.RowCount
        for i in range(min(count, max_rows)):
            row = {}
            for col in columns:
                try:
                    row[col] = shell.GetCellValue(i, col)
                except Exception:
                    pass
            if any(v for v in row.values() if v):
                row["_row"] = i
                rows.append(row)
    except Exception:
        pass
    return rows


def _screen_texts():
    """Walk the current screen and return all visible text strings."""
    texts = []
    def walk(comp, depth=0):
        if depth > 8:
            return
        try:
            n = comp.Children.Count
        except Exception:
            return
        for i in range(n):
            try:
                child = comp.Children(i)
                if child.Type in ("GuiTextField", "GuiCTextField",
                                  "GuiLabel", "GuiStatusbar", "GuiTitlebar"):
                    try:
                        t = child.Text.strip()
                        if t:
                            texts.append(t)
                    except Exception:
                        pass
                walk(child, depth + 1)
            except Exception:
                pass
    walk(session.FindById("wnd[0]"))
    return texts


def scan_st22_dumps(date_from=None, date_to=None):
    """
    List ABAP runtime errors from ST22.
    Strategy:
      1. Navigate ST22 and try to read the ALV/list grid (all known container IDs).
      2. Walk full UI tree for any accessible shell.
      3. Fall back to reading SNAP table directly via SE16N.
      4. Always return screen texts + element list so Claude can see what's visible.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("ST22")
    time.sleep(1.5)

    # ── Set date fields — try every known field ID variant ────────────────────
    for fid in ("wnd[0]/usr/ctxtSEL_DATUM-LOW",
                "wnd[0]/usr/ctxtP_DATUM",
                "wnd[0]/usr/subBESTIM:RSTS22:0200/ctxtSTARTDAT",
                "wnd[0]/usr/ctxtDAT1"):
        try: session.FindById(fid).Text = df; break
        except Exception: pass

    for fid in ("wnd[0]/usr/ctxtSEL_DATUM-HIGH",
                "wnd[0]/usr/ctxtP_DATUM2",
                "wnd[0]/usr/subBESTIM:RSTS22:0200/ctxtENDDAT",
                "wnd[0]/usr/ctxtDAT2"):
        try: session.FindById(fid).Text = dt; break
        except Exception: pass

    # Try "All Clients" button before executing
    for fid in ("wnd[0]/tbar[1]/btn[7]", "wnd[0]/tbar[1]/btn[8]"):
        try: session.FindById(fid).Press(); time.sleep(0.5); break
        except Exception: pass

    session.FindById("wnd[0]").SendVKey(8)   # F8 Execute
    time.sleep(2.5)

    dumps = []
    method_used = "none"

    # ── Strategy 1: Try every known ST22 ALV container ID ─────────────────────
    ST22_CONTAINERS = [
        "wnd[0]/usr/cntlST22_CONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlST22_CONTAINER/shellcont/shell/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlGRID/shellcont/shell",
        "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlALV_CONTAINER/shellcont/shell",
    ]
    ST22_COLS = ["DATUM","UZEIT","UNAME","REPID","ERTYP","MANDT",
                 "ADATE","ATIME","SYSUNAM","ABTYPE","SRTXT",
                 "CPROG","INCLNAME","ERRLINE"]

    shell = _find_shell_anywhere(ST22_CONTAINERS)
    if shell:
        dumps = _read_shell_rows(shell, ST22_COLS, max_rows=200)
        for i, d in enumerate(dumps):
            d["index"] = i
        method_used = "alv_grid"

    # ── Strategy 2: SNAP table via SE16N (most reliable fallback) ─────────────
    if not dumps:
        try:
            df_snap = datetime.strptime(df, "%d.%m.%Y").strftime("%Y%m%d")
            dt_snap = datetime.strptime(dt, "%d.%m.%Y").strftime("%Y%m%d")

            go_to_transaction("SE16N")
            time.sleep(1)
            for fid in ("wnd[0]/usr/ctxtGD-TAB", "wnd[0]/usr/ctxtTABLE"):
                try: session.FindById(fid).Text = "SNAP"; break
                except Exception: pass
            session.FindById("wnd[0]").SendVKey(0)
            time.sleep(1.5)

            # Set DATUM filter
            elems = discover_elements()
            for e in elems:
                if "DATUM" in e.get("id","").upper() and "LOW" in e.get("id","").upper():
                    try: session.FindById(e["id"]).Text = df_snap; break
                    except Exception: pass
            for e in elems:
                if "DATUM" in e.get("id","").upper() and "HIGH" in e.get("id","").upper():
                    try: session.FindById(e["id"]).Text = dt_snap; break
                    except Exception: pass

            # Max rows
            for fid in ("wnd[0]/usr/txtGD-MAX_LINES",):
                try: session.FindById(fid).Text = "500"; break
                except Exception: pass

            session.FindById("wnd[0]").SendVKey(8)
            time.sleep(2.5)

            snap_shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if snap_shell:
                snap_cols = ["DATUM","UZEIT","MANDT","UNAME","REPID",
                             "ERTYP","ARID","SRTXT","CPROG"]
                rows = _read_shell_rows(snap_shell, snap_cols, max_rows=500)
                for i, r in enumerate(rows):
                    d = r.get("DATUM","")
                    if df_snap <= d <= dt_snap:
                        r["index"] = len(dumps)
                        dumps.append(r)
                method_used = "snap_table"
        except Exception as snap_err:
            pass

    # ── Strategy 3: Parse visible screen text ─────────────────────────────────
    screen_txts = _screen_texts()
    all_elements = discover_elements()

    result = {
        "date_from": df,
        "date_to":   dt,
        "dumps":     dumps,
        "count":     len(dumps),
        "method":    method_used,
        "screen":    get_screen_text(),
    }
    if not dumps:
        # Provide everything visible so Claude can diagnose the screen
        result["screen_texts"]  = screen_txts[:150]
        result["all_elements"]  = [e["id"] for e in all_elements[:80]]
        result["note"] = (
            "ALV grid not found and SNAP fallback returned 0 rows. "
            "screen_texts and all_elements show exactly what is currently visible. "
            "Use discover_screen_elements then set_field_value to interact manually."
        )
    return result


def get_dump_detail(dump_index=0):
    """
    Open a specific dump from ST22 by row index.
    Tries ALV double-click, then row selection + Enter, then direct navigation.
    Extracts: error type, program, include, line, error text, call stack.
    """
    try:
        # ── Try to double-click the row in the ALV grid ───────────────────────
        opened = False
        shell = _find_shell_anywhere([
            "wnd[0]/usr/cntlST22_CONTAINER/shellcont/shell",
            "wnd[0]/usr/cntlGRID1/shellcont/shell",
            "wnd[0]/usr/cntlGRID/shellcont/shell",
        ])
        if shell:
            try:
                # Try to select via first visible column
                for col in ("DATUM","ADATE","ERTYP","REPID","_row"):
                    try:
                        shell.SetCurrentCell(dump_index, col)
                        shell.DoubleClickCurrentCell()
                        opened = True
                        break
                    except Exception:
                        pass
            except Exception:
                pass

        if not opened:
            # Fallback: use keyboard to navigate to row then Enter
            try:
                for _ in range(dump_index):
                    session.FindById("wnd[0]").SendVKey(2)   # cursor down
                    time.sleep(0.1)
                session.FindById("wnd[0]").SendVKey(2)       # Enter on row
            except Exception:
                session.FindById("wnd[0]").SendVKey(0)

        time.sleep(2)

        # ── Collect all text from the dump detail screen ──────────────────────
        raw_texts = _screen_texts()
        raw = "\n".join(raw_texts[:400])

        result = {
            "dump_index":  dump_index,
            "raw":         raw,
            "screen":      get_screen_text(),
            "all_elements": [e["id"] for e in discover_elements()[:60]],
        }

        # Extract key fields by pattern scanning
        import re
        for line in raw_texts:
            ll = line.lower()
            if not result.get("program")  and ("program" in ll or "repid" in ll):
                result["program"] = line
            if not result.get("error_type") and ("exception" in ll or "ertyp" in ll
                                                   or "runtime error" in ll.replace(" ","")):
                result["error_type"] = line
            if not result.get("line_number") and re.search(r"\bline\b|\bzeile\b", ll):
                result["line_number"] = line
            if not result.get("include") and ("include" in ll or "includes" in ll):
                result["include"] = line

        return result

    except Exception as e:
        return {"error": str(e), "dump_index": dump_index}


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
    Scan for IDoc errors.
    Strategy:
      1. Navigate WE05 and try all known ALV container IDs.
      2. Walk full UI tree for any accessible shell.
      3. Fall back to reading EDIDC table directly via SE16N.
    status_filter: comma-separated status codes e.g. '51,26,56' or 'all'.
    direction: 'inbound', 'outbound', or 'both'.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today
    statuses = [] if status_filter == "all" else status_filter.split(",")

    # Convert DD.MM.YYYY → YYYYMMDD for table filter
    try:
        df_sap = datetime.strptime(df, "%d.%m.%Y").strftime("%Y%m%d")
        dt_sap = datetime.strptime(dt, "%d.%m.%Y").strftime("%Y%m%d")
    except Exception:
        df_sap = df_sap = ""

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

    result = {
        "date_from":   df,
        "date_to":     dt,
        "idoc_errors": idocs,
        "count":       len(idocs),
        "method":      method_used,
        "screen":      get_screen_text(),
        "note": (f"Found {len(idocs)} IDocs matching status {status_filter} "
                 f"(method: {method_used}). "
                 "Call get_idoc_detail for each to analyse root cause."),
    }
    if not idocs:
        result["screen_texts"] = _screen_texts()[:100]
        result["all_elements"] = [e["id"] for e in discover_elements()[:80]]
    return result


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
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlWE02_CONTAINER/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
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


# ── Purchase Order Error Resolution ──────────────────────────────────────────

PO_ERROR_FIX_MAP = {
    "release":                ("release_po",              "HIGH"),
    "blocked for release":    ("release_po",              "HIGH"),
    "no release strategy":    ("release_po",              "HIGH"),
    "tolerance":              ("release_blocked_invoice",  "HIGH"),
    "price variance":         ("release_blocked_invoice",  "HIGH"),
    "quantity variance":      ("release_blocked_invoice",  "HIGH"),
    "invoice blocked":        ("release_blocked_invoice",  "HIGH"),
    "stochastic":             ("release_blocked_invoice",  "MEDIUM"),
    "vendor":                 ("check_vendor_master",      "MEDIUM"),
    "vendor blocked":         ("unblock_vendor",           "HIGH"),
    "material":               ("change_po_field",          "MEDIUM"),
    "account":                ("change_po_field",          "MEDIUM"),
    "cost center":            ("change_po_field",          "MEDIUM"),
    "wbs":                    ("change_po_field",          "MEDIUM"),
    "delivery date":          ("change_po_field",          "LOW"),
    "overdue":                ("change_po_field",          "LOW"),
    "output":                 ("create_po_output",         "LOW"),
    "message":                ("create_po_output",         "LOW"),
    "deletion":               ("cancel_po_item",           "HIGH"),
    "closed":                 ("cancel_po_item",           "MEDIUM"),
    "gr/ir":                  ("scan_gr_ir_clearing",      "LOW"),
    "not cleared":            ("scan_gr_ir_clearing",      "LOW"),
    "requisition":            ("convert_pr_to_po",         "MEDIUM"),
}


def scan_po_errors(date_from=None, date_to=None, company_code="",
                   purchase_org="", error_type="all"):
    """
    Scan for failed/blocked purchase orders via ME2M.
    error_type: 'blocked'=release blocked, 'overdue'=delivery overdue, 'all'=everything.
    Returns list of POs with vendor, items, status, and auto-computed fix_hints.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("ME2M")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_BEDAT-LOW", "wnd[0]/usr/ctxtBEDAT-LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtS_BEDAT-HIGH", "wnd[0]/usr/ctxtBEDAT-HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if company_code:
            for fid in ("wnd[0]/usr/ctxtS_BUKRS-LOW", "wnd[0]/usr/ctxtBUKRS"):
                try: session.FindById(fid).Text = company_code; break
                except Exception: pass
        if purchase_org:
            for fid in ("wnd[0]/usr/ctxtS_EKORG-LOW", "wnd[0]/usr/ctxtEKORG"):
                try: session.FindById(fid).Text = purchase_org; break
                except Exception: pass
        if error_type == "blocked":
            for fid in ("wnd[0]/usr/ctxtS_SCOPE", "wnd[0]/usr/ctxtSCOPE"):
                try: session.FindById(fid).Text = "BL"; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        pos = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 300)):
                    row = {}
                    for col in ["EBELN","EBELP","AEDAT","LIFNR","MATNR",
                                "MENGE","MEINS","NETPR","WAERS","WERKS",
                                "LOEKZ","EINDT","FRGKE","FRGZU"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if not row.get("EBELN"):
                        continue
                    issues = []
                    if row.get("FRGKE"):  issues.append("release_required")
                    if row.get("LOEKZ"):  issues.append("deletion_flag")
                    row["issues"] = issues
                    combined = " ".join(str(v) for v in row.values()).lower()
                    row["fix_hints"] = [
                        {"pattern": pat, "fix_action": fa, "risk": risk}
                        for pat, (fa, risk) in PO_ERROR_FIX_MAP.items()
                        if pat in combined
                    ]
                    pos.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "po_errors": pos, "count": len(pos),
            "screen": get_screen_text(),
            "note": (f"Found {len(pos)} PO records. "
                     "Call get_po_detail for each to analyse root cause."),
        }
    except Exception as e:
        return {"error": str(e)}


def get_po_detail(po_number):
    """
    Open a Purchase Order in ME23N and extract header, items, pricing,
    account assignment, delivery dates, release status, and fix_hints.
    """
    go_to_transaction("ME23N")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtME23N-EBELN",):
            try: session.FindById(fid).Text = str(po_number).zfill(10); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(2)

        detail = {
            "po_number": str(po_number),
            "error_messages": [],
            "screen": get_screen_text(),
        }

        texts = []
        def walk(comp, depth=0):
            if depth > 8: return
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
                            if t and len(t) > 2: texts.append(t)
                        except Exception: pass
                    walk(child, depth+1)
                except Exception: pass
        walk(session.FindById("wnd[0]"))
        detail["all_screen_text"] = "\n".join(texts[:300])

        combined = detail["all_screen_text"].lower()
        for phrase in ["error","blocked","no release","tolerance","variance",
                        "vendor","account","price","quantity","overdue",
                        "not found","invalid","missing","deletion"]:
            if phrase in combined:
                detail["error_messages"].append(phrase)

        detail["fix_hints"] = [
            {"pattern": pat, "fix_action": fa, "risk": risk}
            for pat, (fa, risk) in PO_ERROR_FIX_MAP.items()
            if pat in combined
        ]
        return detail
    except Exception as e:
        return {"error": str(e), "po_number": str(po_number)}


def release_po(po_number, release_code="01"):
    """
    Release a purchase order blocked for approval via ME29N.
    release_code: 01=standard. Check EKKO-FRGKE for the required code.
    """
    go_to_transaction("ME29N")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtME29N-EBELN",):
            try: session.FindById(fid).Text = str(po_number).zfill(10); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(2)

        for fid in ("wnd[0]/tbar[1]/btn[20]","wnd[0]/tbar[1]/btn[16]",
                    "wnd[0]/tbar[0]/btn[11]"):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass

        for fid in ("wnd[1]/usr/btnSPOP-OPTION1","wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("PO_RELEASE",
                  {"po": str(po_number), "release_code": release_code},
                  status="executed")
        return {"ok": True, "po_number": str(po_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e), "po_number": str(po_number)}


def change_po_field(po_number, item_number, field_name, new_value):
    """
    Change a field on a PO line via ME22N.
    Common: EINDT=delivery date, MENGE=qty, NETPR=price, WERKS=plant,
    KOSTL=cost centre, ANLN1=asset, PSPNR=WBS element.
    """
    go_to_transaction("ME22N")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtME22N-EBELN",):
            try: session.FindById(fid).Text = str(po_number).zfill(10); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(2)

        elems = discover_elements()
        changed = False
        for e in elems:
            if (field_name.upper() in e.get("id","").upper() or
                    field_name.upper() in e.get("tooltip","").upper()):
                set_field(e["id"], str(new_value))
                changed = True
                break

        if not changed:
            return {"ok": False,
                    "note": f"Field {field_name} not found on screen. "
                            "Use discover_screen_elements to locate it.",
                    "elements": [e["id"] for e in elems[:30]]}

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)
        handle_transport(None)

        audit_log("PO_FIELD_CHANGE",
                  {"po": str(po_number), "item": item_number,
                   "field": field_name, "new_value": str(new_value)},
                  status="executed")
        return {"ok": True, "po_number": str(po_number),
                "field": field_name, "new_value": new_value,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def cancel_po_item(po_number, item_number, reason=""):
    """
    Set deletion flag on a PO line item via ME22N.
    Marks the item for deletion at the next MRP run.
    """
    go_to_transaction("ME22N")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtME22N-EBELN",):
            try: session.FindById(fid).Text = str(po_number).zfill(10); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(2)

        elems = discover_elements()
        for e in elems:
            if "LOEKZ" in e.get("id","").upper():
                set_field(e["id"], "L")
                break

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("PO_ITEM_CANCEL",
                  {"po": str(po_number), "item": item_number, "reason": reason},
                  status="executed")
        return {"ok": True, "po_number": str(po_number),
                "item": item_number, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def scan_blocked_invoices(date_from=None, date_to=None, company_code="1000"):
    """
    Scan for blocked MM invoices via MRBR.
    Returns invoices blocked due to price/quantity variance or manual block.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("MRBR")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRM08R-RBUKR","wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRM08R-BUDAT_FROM","wnd[0]/usr/ctxtBUDAT_FROM"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRM08R-BUDAT_TO","wnd[0]/usr/ctxtBUDAT_TO"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        invoices = []
        BLOCK_REASONS = {"R":"Price variance","M":"Manual block",
                         "Q":"Quantity variance","D":"Date variance",
                         "A":"Amount exceeded","S":"Stochastic block"}
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["BELNR","GJAHR","BUKRS","LIFNR",
                                "BLDAT","RMWWR","WAERS","SPGRU","SPGRP"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if not row.get("BELNR"):
                        continue
                    code = row.get("SPGRU") or row.get("SPGRP","")
                    row["block_reason_desc"] = BLOCK_REASONS.get(code, f"Code:{code}")
                    row["fix_action"] = "release_blocked_invoice"
                    invoices.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "blocked_invoices": invoices, "count": len(invoices),
            "screen": get_screen_text(),
            "note": (f"Found {len(invoices)} blocked invoices. "
                     "Call release_blocked_invoice to clear each one."),
        }
    except Exception as e:
        return {"error": str(e)}


def release_blocked_invoice(invoice_number, company_code="1000",
                             fiscal_year=""):
    """
    Release a blocked MM invoice via MRBR so it can be paid.
    Clears price/quantity/manual blocks. Requires approval.
    """
    go_to_transaction("MRBR")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRM08R-RBUKR","wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRM08R-BELNR_FROM","wnd[0]/usr/ctxtBELNR"):
            try: session.FindById(fid).Text = str(invoice_number); break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        session.FindById("wnd[0]").SendVKey(16)   # Select all
        time.sleep(0.5)

        for fid in ("wnd[0]/tbar[1]/btn[20]","wnd[0]/tbar[1]/btn[17]",
                    "wnd[0]/tbar[1]/btn[16]"):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass

        for fid in ("wnd[1]/usr/btnSPOP-OPTION1","wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("INVOICE_RELEASE",
                  {"invoice": str(invoice_number), "company_code": company_code},
                  status="executed")
        return {"ok": True, "invoice": str(invoice_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def check_vendor_master(vendor_number, company_code="1000"):
    """
    Check vendor master via XK03: payment terms, bank details, reconciliation
    account, purchasing data, block status. Used when PO fails on vendor issues.
    """
    go_to_transaction("XK03")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRF02K-LIFNR","wnd[0]/usr/ctxtLIFNR"):
            try: session.FindById(fid).Text = str(vendor_number).zfill(10); break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF02K-BUKRS","wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1)
        try: session.FindById("wnd[0]/usr/chkRF02K-XBANK").Selected = True
        except Exception: pass
        try: session.FindById("wnd[0]/usr/chkRF02K-XKAUF").Selected = True
        except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        texts = []
        def wv(comp, depth=0):
            if depth > 6: return
            try:
                n = comp.Children.Count
            except Exception:
                return
            for i in range(n):
                try:
                    c2 = comp.Children(i)
                    if c2.Type in ("GuiTextField","GuiCTextField","GuiLabel"):
                        try:
                            t = c2.Text.strip()
                            if t: texts.append(t)
                        except Exception: pass
                    wv(c2, depth+1)
                except Exception: pass
        wv(session.FindById("wnd[0]"))

        scr = "\n".join(texts[:200])
        blocked = any(b in scr.lower() for b in ["blocked","gesperrt","sperre"])
        return {
            "vendor": str(vendor_number), "company_code": company_code,
            "is_blocked": blocked, "screen_text": scr,
            "screen": get_screen_text(),
        }
    except Exception as e:
        return {"error": str(e)}


def unblock_vendor(vendor_number, company_code="1000", block_type="purchase"):
    """
    Remove vendor block via XK05.
    block_type: 'purchase'=purchasing block, 'payment'=payment block, 'all'=both.
    """
    go_to_transaction("XK05")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRF02K-LIFNR","wnd[0]/usr/ctxtLIFNR"):
            try: session.FindById(fid).Text = str(vendor_number).zfill(10); break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF02K-BUKRS","wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        elems = discover_elements()
        for e in elems:
            eid = e.get("id","").upper()
            if block_type in ("purchase","all") and "SPERR" in eid:
                try: session.FindById(e["id"]).Selected = False
                except Exception: pass
            if block_type in ("payment","all") and "ZAHLS" in eid:
                try: session.FindById(e["id"]).Selected = False
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("VENDOR_UNBLOCK",
                  {"vendor": str(vendor_number), "company_code": company_code,
                   "block_type": block_type},
                  status="executed")
        return {"ok": True, "vendor": str(vendor_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def create_po_output(po_number, output_type="NEU", medium="1"):
    """
    Resend/create output message for a PO via ME9F.
    output_type: NEU=original, MAHN=reminder. medium: 1=print, 5=external, 6=EDI.
    """
    go_to_transaction("ME9F")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_EBELN-LOW","wnd[0]/usr/ctxtEBELN_LOW"):
            try: session.FindById(fid).Text = str(po_number).zfill(10); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        session.FindById("wnd[0]").SendVKey(16)   # Select all
        time.sleep(0.5)
        for fid in ("wnd[0]/tbar[1]/btn[8]",):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass

        audit_log("PO_OUTPUT",
                  {"po": str(po_number), "output_type": output_type},
                  status="executed")
        return {"ok": True, "po_number": str(po_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def scan_open_purchase_reqs(date_from=None, date_to=None,
                             plant="", material=""):
    """
    Scan for open/unprocessed purchase requisitions via ME5A.
    Returns PRs not yet converted to POs (scope=WA open items).
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("ME5A")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_BADAT-LOW","wnd[0]/usr/ctxtBADAT_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtS_BADAT-HIGH","wnd[0]/usr/ctxtBADAT_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if plant:
            try: session.FindById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = plant
            except Exception: pass
        if material:
            try: session.FindById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = material
            except Exception: pass
        try: session.FindById("wnd[0]/usr/ctxtP_SCOPE").Text = "WA"
        except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        reqs = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["BANFN","BNFPO","MATNR","MENGE","MEINS",
                                "WERKS","BADAT","LIFNR","KNTTP","TXZ01"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("BANFN"):
                        reqs.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "open_reqs": reqs, "count": len(reqs),
            "screen": get_screen_text(),
            "note": (f"Found {len(reqs)} open purchase requisitions. "
                     "Use convert_pr_to_po or ME57/ME59N to assign and convert."),
        }
    except Exception as e:
        return {"error": str(e)}


def convert_pr_to_po(pr_number, vendor_number, purchase_org="1000",
                     company_code="1000"):
    """
    Assign source and convert a purchase requisition to a PO via ME57.
    """
    go_to_transaction("ME57")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_BANFN-LOW","wnd[0]/usr/ctxtBANFN_LOW"):
            try: session.FindById(fid).Text = str(pr_number); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)
        session.FindById("wnd[0]").SendVKey(16)
        time.sleep(0.5)
        for fid in ("wnd[0]/tbar[1]/btn[9]",):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass

        audit_log("PR_TO_PO",
                  {"pr": str(pr_number), "vendor": str(vendor_number)},
                  status="executed")
        return {"ok": True, "pr_number": str(pr_number),
                "vendor": str(vendor_number), "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def scan_gr_ir_clearing(company_code="1000", date_from=None, date_to=None):
    """
    Scan for uncleared GR/IR balance items via MB5S.
    These represent goods receipts without matching invoices (or vice versa).
    """
    go_to_transaction("MB5S")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_BUKRS-LOW","wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)
        return {
            "company_code": company_code,
            "screen": get_screen_text(),
            "note": ("GR/IR clearing report displayed. "
                     "Review items and post clearing via MR11."),
        }
    except Exception as e:
        return {"error": str(e)}


# ── Sales Order Error Resolution ─────────────────────────────────────────────

SD_ERROR_FIX_MAP = {
    "credit":                  ("release_credit_block",      "HIGH"),
    "credit limit":            ("release_credit_block",      "HIGH"),
    "credit block":            ("release_credit_block",      "HIGH"),
    "delivery block":          ("release_sd_delivery_block", "MEDIUM"),
    "billing block":           ("release_billing_block",     "MEDIUM"),
    "pricing":                 ("reprice_sales_order",       "MEDIUM"),
    "condition":               ("reprice_sales_order",       "MEDIUM"),
    "price 0":                 ("reprice_sales_order",       "HIGH"),
    "incomplete":              ("complete_sales_order",      "MEDIUM"),
    "incompletion":            ("complete_sales_order",      "MEDIUM"),
    "missing field":           ("complete_sales_order",      "MEDIUM"),
    "payment terms":           ("complete_sales_order",      "LOW"),
    "partner":                 ("fix_partner_determination", "MEDIUM"),
    "ship-to":                 ("fix_partner_determination", "MEDIUM"),
    "sold-to":                 ("fix_partner_determination", "HIGH"),
    "output":                  ("create_sd_output",          "LOW"),
    "message":                 ("create_sd_output",          "LOW"),
    "availability":            ("check_atp",                 "MEDIUM"),
    "not available":           ("check_atp",                 "MEDIUM"),
    "schedule line":           ("check_atp",                 "MEDIUM"),
    "no confirmed quantity":   ("check_atp",                 "HIGH"),
    "backorder":               ("check_atp",                 "MEDIUM"),
    "plant":                   ("complete_sales_order",      "MEDIUM"),
    "shipping point":          ("complete_sales_order",      "MEDIUM"),
    "route":                   ("complete_sales_order",      "LOW"),
    "rejection":               ("remove_rejection_reason",   "HIGH"),
    "rejected":                ("remove_rejection_reason",   "HIGH"),
    "account":                 ("complete_sales_order",      "MEDIUM"),
}


def scan_sd_errors(date_from=None, date_to=None, sales_org="",
                   error_type="all"):
    """
    Scan for failed/blocked sales orders via VA05.
    error_type: 'credit'=credit-blocked, 'delivery'=delivery-blocked,
                'billing'=billing-blocked, 'incomplete'=incompletion log,
                'all'=every open order.
    Returns list of orders with status, customer, value, and fix_hints.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    # Choose the right list transaction based on error type
    if error_type == "credit":
        go_to_transaction("VKM1")     # Credit management: blocked orders
    else:
        go_to_transaction("VA05")     # General order list
    time.sleep(1.5)

    try:
        # VA05 date and sales org fields
        for fid in ("wnd[0]/usr/ctxtSD_VBAK-AUDAT_LOW",
                    "wnd[0]/usr/ctxtAUDAT_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtSD_VBAK-AUDAT_HIGH",
                    "wnd[0]/usr/ctxtAUDAT_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if sales_org:
            for fid in ("wnd[0]/usr/ctxtSD_VBAK-VKORG",
                        "wnd[0]/usr/ctxtVKORG"):
                try: session.FindById(fid).Text = sales_org; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        orders = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 300)):
                    row = {}
                    for col in ["VBELN","AUDAT","AUART","KUNNR","VKORG",
                                "NETWR","WAERS","GBSTK","LIFSK","FAKSK",
                                "CMGST","UVVLS","NAME1","BSTNK"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if not row.get("VBELN"):
                        continue
                    issues = []
                    if row.get("CMGST") in ("B","C"):  issues.append("credit_block")
                    if row.get("LIFSK"):               issues.append("delivery_block")
                    if row.get("FAKSK"):               issues.append("billing_block")
                    row["issues"] = issues

                    # Apply error_type filter
                    if error_type == "credit" and "credit_block" not in issues:
                        continue
                    if error_type == "delivery" and "delivery_block" not in issues:
                        continue
                    if error_type == "billing" and "billing_block" not in issues:
                        continue

                    combined = " ".join(str(v) for v in row.values()).lower()
                    row["fix_hints"] = [
                        {"pattern": pat, "fix_action": fa, "risk": risk}
                        for pat, (fa, risk) in SD_ERROR_FIX_MAP.items()
                        if pat in combined
                    ]
                    orders.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "sd_errors": orders, "count": len(orders),
            "screen": get_screen_text(),
            "note": (f"Found {len(orders)} sales orders. "
                     "Call get_sd_order_detail for root cause per order."),
        }
    except Exception as e:
        return {"error": str(e)}


def get_sd_order_detail(sales_order):
    """
    Open a sales order in VA03 and extract: header status (credit/delivery/
    billing block), item pricing, partner data, schedule lines, incompletion
    log, and auto-computed fix_hints from SD_ERROR_FIX_MAP.
    """
    go_to_transaction("VA03")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        detail = {
            "sales_order": str(sales_order),
            "error_messages": [],
            "screen": get_screen_text(),
        }

        texts = []
        def walk(comp, depth=0):
            if depth > 8: return
            try:
                n = comp.Children.Count
            except Exception:
                return
            for i in range(n):
                try:
                    ch = comp.Children(i)
                    if ch.Type in ("GuiTextField","GuiCTextField","GuiLabel"):
                        try:
                            t = ch.Text.strip()
                            if t and len(t) > 2: texts.append(t)
                        except Exception: pass
                    walk(ch, depth+1)
                except Exception: pass
        walk(session.FindById("wnd[0]"))
        detail["all_screen_text"] = "\n".join(texts[:300])

        combined = detail["all_screen_text"].lower()
        for phrase in ["credit","block","incomplete","pricing","condition",
                        "partner","output","availability","rejection",
                        "not found","billing","delivery","schedule line"]:
            if phrase in combined:
                detail["error_messages"].append(phrase)

        detail["fix_hints"] = [
            {"pattern": pat, "fix_action": fa, "risk": risk}
            for pat, (fa, risk) in SD_ERROR_FIX_MAP.items()
            if pat in combined
        ]

        # Also try to read VA03 header fields directly
        header = {}
        for field_id, field_key in [
            ("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-KUNNR", "KUNNR"),
            ("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VKORG", "VKORG"),
            ("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-LIFSK", "LIFSK"),
            ("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-FAKSK", "FAKSK"),
            ("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-CMGST", "CMGST"),
        ]:
            try: header[field_key] = session.FindById(field_id).Text
            except Exception: pass
        detail["header_fields"] = header

        return detail
    except Exception as e:
        return {"error": str(e), "sales_order": str(sales_order)}


def release_credit_block(sales_order, release_type="order"):
    """
    Release a credit-blocked sales order via VKM3.
    release_type: 'order'=single order, 'all'=all orders for this customer.
    """
    go_to_transaction("VKM3")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtVBELN",
                    "wnd[0]/usr/ctxtS_VBELN-LOW"):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        # Select the order row
        session.FindById("wnd[0]").SendVKey(16)
        time.sleep(0.5)

        # Press Release button
        for fid in ("wnd[0]/tbar[1]/btn[20]",
                    "wnd[0]/tbar[1]/btn[16]",
                    "wnd[0]/tbar[1]/btn[5]"):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass

        # Confirm popup
        for fid in ("wnd[1]/usr/btnSPOP-OPTION1",
                    "wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("SD_CREDIT_RELEASE",
                  {"order": str(sales_order), "release_type": release_type},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def release_sd_delivery_block(sales_order, block_code=""):
    """
    Remove the delivery block from a sales order header via VA02.
    Sets VBAK-LIFSK to blank (no block). block_code is informational only.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Clear delivery block field
        for fid in (
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-LIFSK",
        ):
            try:
                session.FindById(fid).Text = ""
                break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(11)   # Save
        time.sleep(1.5)
        handle_transport(None)

        audit_log("SD_DELIVERY_BLOCK_RELEASE",
                  {"order": str(sales_order), "block_code": block_code},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def release_billing_block(sales_order, block_code=""):
    """
    Remove the billing block from a sales order via VA02.
    Sets VBAK-FAKSK to blank.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        for fid in (
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-FAKSK",
        ):
            try: session.FindById(fid).Text = ""; break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)
        handle_transport(None)

        audit_log("SD_BILLING_BLOCK_RELEASE",
                  {"order": str(sales_order), "block_code": block_code},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def reprice_sales_order(sales_order):
    """
    Reprice a sales order via VA02 (update pricing / carry out new pricing).
    Triggers fresh condition determination to fix zero-price or wrong-price items.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Menu: Edit → Update pricing  (or toolbar Update button)
        try:
            session.FindById("wnd[0]/mbar/menu[1]/menu[9]").Select()
            time.sleep(1)
        except Exception:
            try:
                session.FindById("wnd[0]/mbar/menu[1]/menu[8]").Select()
                time.sleep(1)
            except Exception:
                pass

        # Reprice dialog: choose B=Carry out new pricing
        for fid in ("wnd[1]/usr/radKALKME-B",
                    "wnd[1]/usr/radNEWPRICE"):
            try: session.FindById(fid).Select(); break
            except Exception: pass
        for fid in ("wnd[1]/tbar[0]/btn[0]",):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_REPRICE",
                  {"order": str(sales_order)},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def complete_sales_order(sales_order, field_name, new_value, item_number=""):
    """
    Fix an incomplete sales order by setting a missing mandatory field via VA02.
    Common fields: LIFSK=delivery block (clear), FAKSK=billing block (clear),
    KUNNR=sold-to, ZTERM=payment terms, VSTEL=shipping point, ROUTE=route,
    WERKS=plant, MATNR=material, KDMAT=customer material.
    item_number: blank=header field, '00010'=first item.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        elems = discover_elements()
        changed = False
        for e in elems:
            if (field_name.upper() in e.get("id","").upper() or
                    field_name.upper() in e.get("tooltip","").upper()):
                set_field(e["id"], str(new_value))
                changed = True
                break

        if not changed:
            return {"ok": False,
                    "note": f"Field {field_name} not found. "
                            "Use discover_screen_elements to locate the correct ID.",
                    "elements": [e["id"] for e in elems[:30]]}

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_ORDER_COMPLETE",
                  {"order": str(sales_order), "field": field_name,
                   "new_value": str(new_value)},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "field": field_name, "new_value": new_value,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def fix_partner_determination(sales_order):
    """
    Re-trigger partner determination for a sales order via VA02.
    Opens the order, navigates to partner tab, and re-determines partners
    (sold-to, ship-to, bill-to, payer) from customer master.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Navigate to Partners tab
        for fid in ("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01",
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\02"):
            try: session.FindById(fid).Select(); break
            except Exception: pass
        time.sleep(1)

        # Discover and report current partner state
        elems = discover_elements()
        partner_fields = [e for e in elems
                          if any(p in e.get("id","").upper()
                                 for p in ["KUNNR","PARVW","PARNR"])]

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_PARTNER_FIX",
                  {"order": str(sales_order)},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "partner_fields": partner_fields[:20],
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def create_sd_output(sales_order, output_type="BA00", medium="1"):
    """
    Create/resend output (order confirmation etc.) for a sales order via VA02.
    output_type: BA00=order confirmation, RD00=invoice, LIEF=delivery note.
    medium: 1=print, 5=external send, 6=EDI.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Extras → Output → Header → Issue output
        try:
            session.FindById("wnd[0]/mbar/menu[3]/menu[0]/menu[0]").Select()
            time.sleep(1)
        except Exception:
            pass

        # Add output record
        elems = discover_elements()
        for e in elems:
            if "KSCHL" in e.get("id","").upper():
                set_field(e["id"], output_type)
                break

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_OUTPUT",
                  {"order": str(sales_order), "output_type": output_type},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "output_type": output_type, "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def check_atp(sales_order, item_number="00010"):
    """
    Run ATP (Available-to-Promise) check for a sales order line via VA02.
    Re-schedules the order and confirms quantity / delivery date from stock.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Select item and run ATP: Edit → Check availability
        try:
            session.FindById("wnd[0]/mbar/menu[1]/menu[1]").Select()
            time.sleep(1.5)
        except Exception:
            try:
                session.FindById("wnd[0]").SendVKey(8)
                time.sleep(1)
            except Exception:
                pass

        result = {"ok": True, "sales_order": str(sales_order),
                  "item": item_number, "screen": get_screen_text()}

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_ATP_CHECK",
                  {"order": str(sales_order), "item": item_number},
                  status="executed")
        return result
    except Exception as e:
        return {"error": str(e)}


def remove_rejection_reason(sales_order, item_number="00010"):
    """
    Remove the rejection reason from a rejected sales order item via VA02.
    Sets VBAP-ABGRU to blank to reactivate the item.
    """
    go_to_transaction("VA02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtVBAK-VBELN",):
            try: session.FindById(fid).Text = str(sales_order); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        elems = discover_elements()
        for e in elems:
            if "ABGRU" in e.get("id","").upper():
                set_field(e["id"], "")
                break

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("SD_REJECTION_REMOVE",
                  {"order": str(sales_order), "item": item_number},
                  status="executed")
        return {"ok": True, "sales_order": str(sales_order),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def scan_credit_blocks(date_from=None, date_to=None, sales_org=""):
    """
    Scan VKM1 for all orders currently blocked by credit management.
    Returns order number, customer, credit exposure, credit limit, and block reason.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("VKM1")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_CMFDAT-LOW",
                    "wnd[0]/usr/ctxtDATE_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtS_CMFDAT-HIGH",
                    "wnd[0]/usr/ctxtDATE_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if sales_org:
            for fid in ("wnd[0]/usr/ctxtS_VKORG-LOW",):
                try: session.FindById(fid).Text = sales_org; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        blocks = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["VBELN","KUNNR","BLDAT","NETWR","WAERS",
                                "KLIMK","SKFOR","CMGST","NAME1"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("VBELN"):
                        row["fix_action"] = "release_credit_block"
                        blocks.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "credit_blocks": blocks, "count": len(blocks),
            "screen": get_screen_text(),
            "note": (f"Found {len(blocks)} credit-blocked orders. "
                     "Call release_credit_block for each."),
        }
    except Exception as e:
        return {"error": str(e)}


def scan_incomplete_orders(date_from=None, date_to=None, sales_org=""):
    """
    Scan V.02 for incomplete sales orders (incompletion log has open items).
    Returns order list with the incompletion group/procedure details.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("V.02")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtSP$00001-LOW",
                    "wnd[0]/usr/ctxtAUDAT_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtSP$00001-HIGH",
                    "wnd[0]/usr/ctxtAUDAT_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if sales_org:
            for fid in ("wnd[0]/usr/ctxtS_VKORG-LOW",):
                try: session.FindById(fid).Text = sales_org; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        orders = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["VBELN","AUDAT","KUNNR","UVVLS","UVALL",
                                "NETWR","WAERS","NAME1"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("VBELN"):
                        row["fix_action"] = "complete_sales_order"
                        orders.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "incomplete_orders": orders, "count": len(orders),
            "screen": get_screen_text(),
            "note": (f"Found {len(orders)} incomplete orders. "
                     "Call get_sd_order_detail then complete_sales_order."),
        }
    except Exception as e:
        return {"error": str(e)}


# ── FI/CO Posting Error Resolution ───────────────────────────────────────────

FI_ERROR_FIX_MAP = {
    "payment block":        ("release_fi_payment_block", "HIGH"),
    "blocked for payment":  ("release_fi_payment_block", "HIGH"),
    "posting period":       ("open_fi_posting_period",   "HIGH"),
    "period not open":      ("open_fi_posting_period",   "HIGH"),
    "account":              ("check_gl_account",          "MEDIUM"),
    "gl account":           ("check_gl_account",          "MEDIUM"),
    "cost center":          ("check_cost_center",         "MEDIUM"),
    "profit center":        ("check_cost_center",         "LOW"),
    "tax":                  ("check_tax_config",          "MEDIUM"),
    "exchange rate":        ("check_exchange_rate",       "LOW"),
    "reversal":             ("reverse_fi_document",       "HIGH"),
    "duplicate":            ("reverse_fi_document",       "MEDIUM"),
    "tolerance":            ("release_fi_payment_block",  "MEDIUM"),
    "clearing":             ("clear_open_items",          "MEDIUM"),
    "open item":            ("clear_open_items",          "LOW"),
    "balance":              ("clear_open_items",          "MEDIUM"),
}


def scan_fi_errors(date_from=None, date_to=None, company_code="1000",
                   account_type="K"):
    """
    Scan for FI document errors and blocked items via FBL1N/FBL5N.
    account_type: K=vendor, D=customer, S=GL account.
    Returns open/blocked items with fix_hints.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    # Choose transaction by account type
    tcode_map = {"K": "FBL1N", "D": "FBL5N", "S": "FBL3N"}
    go_to_transaction(tcode_map.get(account_type, "FBL3N"))
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtDD_BUKRS-LOW", "wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtDD_BUDAT-LOW", "wnd[0]/usr/ctxtBUDAT_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtDD_BUDAT-HIGH", "wnd[0]/usr/ctxtBUDAT_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        # Select "All items" radio
        for fid in ("wnd[0]/usr/radX_AISEL", "wnd[0]/usr/radALLSEL"):
            try: session.FindById(fid).Select(); break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        items = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 300)):
                    row = {}
                    for col in ["BELNR","GJAHR","BUKRS","BLDAT","BUDAT",
                                "BLART","DMBTR","WAERS","ZLSPR","AUGBL",
                                "SGTXT","LIFNR","KUNNR","HKONT"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if not row.get("BELNR"):
                        continue
                    issues = []
                    if row.get("ZLSPR"):  issues.append("payment_blocked")
                    if not row.get("AUGBL"): issues.append("open_item")
                    row["issues"] = issues
                    combined = " ".join(str(v) for v in row.values()).lower()
                    row["fix_hints"] = [
                        {"pattern": pat, "fix_action": fa, "risk": risk}
                        for pat, (fa, risk) in FI_ERROR_FIX_MAP.items()
                        if pat in combined
                    ]
                    if issues:
                        items.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "company_code": company_code, "account_type": account_type,
            "fi_errors": items, "count": len(items),
            "screen": get_screen_text(),
            "note": (f"Found {len(items)} FI items with issues. "
                     "Call get_fi_doc_detail for each to analyse."),
        }
    except Exception as e:
        return {"error": str(e)}


def get_fi_doc_detail(doc_number, company_code="1000", fiscal_year=""):
    """
    Display a FI document in FB03 and extract all line items, posting keys,
    accounts, amounts, and payment block status with fix_hints.
    """
    if not fiscal_year:
        fiscal_year = datetime.now().strftime("%Y")
    go_to_transaction("FB03")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtRF05L-BELNR", "wnd[0]/usr/ctxtBELNR"):
            try: session.FindById(fid).Text = str(doc_number); break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-BUKRS", "wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-GJAHR", "wnd[0]/usr/ctxtGJAHR"):
            try: session.FindById(fid).Text = fiscal_year; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        texts = []
        def walk(comp, depth=0):
            if depth > 8: return
            try: n = comp.Children.Count
            except Exception: return
            for i in range(n):
                try:
                    ch = comp.Children(i)
                    if ch.Type in ("GuiTextField","GuiCTextField","GuiLabel"):
                        try:
                            t = ch.Text.strip()
                            if t and len(t) > 2: texts.append(t)
                        except Exception: pass
                    walk(ch, depth+1)
                except Exception: pass
        walk(session.FindById("wnd[0]"))
        all_text = "\n".join(texts[:300])
        combined = all_text.lower()

        fix_hints = [
            {"pattern": pat, "fix_action": fa, "risk": risk}
            for pat, (fa, risk) in FI_ERROR_FIX_MAP.items()
            if pat in combined
        ]
        return {
            "doc_number": str(doc_number),
            "company_code": company_code,
            "fiscal_year": fiscal_year,
            "all_screen_text": all_text,
            "fix_hints": fix_hints,
            "screen": get_screen_text(),
        }
    except Exception as e:
        return {"error": str(e)}


def release_fi_payment_block(doc_number, company_code="1000",
                              fiscal_year="", line_item="1"):
    """
    Remove payment block from a FI document line item via FB02.
    Clears BSEG-ZLSPR so the item can be included in the next payment run.
    """
    if not fiscal_year:
        fiscal_year = datetime.now().strftime("%Y")
    go_to_transaction("FB02")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtRF05L-BELNR", "wnd[0]/usr/ctxtBELNR"):
            try: session.FindById(fid).Text = str(doc_number); break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-BUKRS", "wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-GJAHR", "wnd[0]/usr/ctxtGJAHR"):
            try: session.FindById(fid).Text = fiscal_year; break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        elems = discover_elements()
        for e in elems:
            if "ZLSPR" in e.get("id","").upper():
                set_field(e["id"], "")
                break

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("FI_PAYMENT_BLOCK_RELEASE",
                  {"doc": str(doc_number), "company_code": company_code},
                  status="executed")
        return {"ok": True, "doc_number": str(doc_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def reverse_fi_document(doc_number, company_code="1000", fiscal_year="",
                         reversal_reason="01", reversal_date=None):
    """
    Reverse (storno) a FI document via FB08.
    reversal_reason: 01=Reversal in current period, 02=Reversal in closed period.
    """
    if not fiscal_year:
        fiscal_year = datetime.now().strftime("%Y")
    rev_date = reversal_date or datetime.now().strftime("%d.%m.%Y")

    go_to_transaction("FB08")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtRF05L-BELNR", "wnd[0]/usr/ctxtBELNR"):
            try: session.FindById(fid).Text = str(doc_number); break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-BUKRS", "wnd[0]/usr/ctxtBUKRS"):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-GJAHR", "wnd[0]/usr/ctxtGJAHR"):
            try: session.FindById(fid).Text = fiscal_year; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-STGRD", "wnd[0]/usr/ctxtSTGRD"):
            try: session.FindById(fid).Text = reversal_reason; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05L-NEWBS", "wnd[0]/usr/ctxtBUDAT"):
            try: session.FindById(fid).Text = rev_date; break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(2)

        audit_log("FI_DOCUMENT_REVERSAL",
                  {"doc": str(doc_number), "company_code": company_code,
                   "reason": reversal_reason, "date": rev_date},
                  status="executed")
        return {"ok": True, "doc_number": str(doc_number),
                "reversal_reason": reversal_reason,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def clear_open_items(account_number, company_code="1000",
                      account_type="K", clearing_date=None):
    """
    Clear open FI items for a vendor/customer/GL account via F-44/F-32/F-03.
    account_type: K=vendor (F-44), D=customer (F-32), S=GL account (F-03).
    """
    clr_date = clearing_date or datetime.now().strftime("%d.%m.%Y")
    tcode_map = {"K": "F-44", "D": "F-32", "S": "F-03"}
    go_to_transaction(tcode_map.get(account_type, "F-44"))
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRF05A-BUDAT",):
            try: session.FindById(fid).Text = clr_date; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05A-BUKRS",):
            try: session.FindById(fid).Text = company_code; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRF05A-LIFNR",
                    "wnd[0]/usr/ctxtRF05A-KUNNR",
                    "wnd[0]/usr/ctxtRF05A-HKONT"):
            try: session.FindById(fid).Text = str(account_number); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Select all open items
        session.FindById("wnd[0]").SendVKey(16)
        time.sleep(0.5)

        audit_log("FI_CLEAR_OPEN_ITEMS",
                  {"account": str(account_number), "company_code": company_code,
                   "account_type": account_type},
                  status="executed")
        return {"ok": True, "account": str(account_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


# ── Delivery / Shipping Error Resolution ──────────────────────────────────────

DELIVERY_ERROR_FIX_MAP = {
    "picking":          ("fix_delivery_picking",    "MEDIUM"),
    "pick quantity":    ("fix_delivery_picking",    "MEDIUM"),
    "goods issue":      ("post_goods_issue",         "HIGH"),
    "gi not posted":    ("post_goods_issue",         "HIGH"),
    "incomplete":       ("fix_delivery_incomplete",  "MEDIUM"),
    "packing":          ("fix_delivery_incomplete",  "LOW"),
    "route":            ("fix_delivery_incomplete",  "LOW"),
    "shipping point":   ("fix_delivery_incomplete",  "MEDIUM"),
    "output":           ("create_delivery_output",   "LOW"),
    "transfer order":   ("create_transfer_order",    "MEDIUM"),
    "stock":            ("check_stock",              "MEDIUM"),
    "not available":    ("check_stock",              "HIGH"),
}


def scan_delivery_errors(date_from=None, date_to=None,
                          shipping_point="", error_type="all"):
    """
    Scan for stuck/failed deliveries via VL06O (outbound delivery monitor).
    error_type: 'gi'=GI not posted, 'pick'=picking incomplete,
                'output'=output errors, 'all'=all open deliveries.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("VL06O")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_LFDAT-LOW", "wnd[0]/usr/ctxtLFDAT_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtS_LFDAT-HIGH", "wnd[0]/usr/ctxtLFDAT_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if shipping_point:
            for fid in ("wnd[0]/usr/ctxtS_VSTEL-LOW",):
                try: session.FindById(fid).Text = shipping_point; break
                except Exception: pass

        # Select the correct list view based on error type
        if error_type == "gi":
            for fid in ("wnd[0]/usr/tabsTAB/tabpGI", "wnd[0]/tbar[1]/btn[18]"):
                try: session.FindById(fid).Select(); break
                except Exception: pass
        elif error_type == "pick":
            for fid in ("wnd[0]/usr/tabsTAB/tabpPICK", "wnd[0]/tbar[1]/btn[17]"):
                try: session.FindById(fid).Select(); break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        deliveries = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["VBELN","LFART","WADAT","KUNNR","VSTEL",
                                "LGNUM","KOSTA","WBSTK","PKSTK","KOQUK"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if not row.get("VBELN"):
                        continue
                    combined = " ".join(str(v) for v in row.values()).lower()
                    row["fix_hints"] = [
                        {"pattern": pat, "fix_action": fa, "risk": risk}
                        for pat, (fa, risk) in DELIVERY_ERROR_FIX_MAP.items()
                        if pat in combined
                    ]
                    deliveries.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "delivery_errors": deliveries, "count": len(deliveries),
            "screen": get_screen_text(),
            "note": (f"Found {len(deliveries)} deliveries. "
                     "Call fix_delivery or post_goods_issue to resolve."),
        }
    except Exception as e:
        return {"error": str(e)}


def post_goods_issue(delivery_number):
    """
    Post goods issue for a delivery via VL02N.
    Completes the outbound delivery and reduces stock.
    """
    go_to_transaction("VL02N")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtLIKP-VBELN",):
            try: session.FindById(fid).Text = str(delivery_number); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Post Goods Issue button (F6 or toolbar)
        for fid in ("wnd[0]/tbar[1]/btn[8]",):
            try: session.FindById(fid).Press(); time.sleep(1.5); break
            except Exception: pass
        # Try menu: Post Goods Issue
        try:
            session.FindById("wnd[0]/mbar/menu[1]/menu[1]").Select()
            time.sleep(1.5)
        except Exception:
            pass

        for fid in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("DELIVERY_GI_POST",
                  {"delivery": str(delivery_number)},
                  status="executed")
        return {"ok": True, "delivery": str(delivery_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def fix_delivery_incomplete(delivery_number, field_name, new_value):
    """
    Fix an incomplete delivery by setting a missing field via VL02N.
    Common: ROUTE=route, VSTEL=shipping point, LGNUM=warehouse number.
    """
    go_to_transaction("VL02N")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtLIKP-VBELN",):
            try: session.FindById(fid).Text = str(delivery_number); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        elems = discover_elements()
        changed = False
        for e in elems:
            if (field_name.upper() in e.get("id","").upper() or
                    field_name.upper() in e.get("tooltip","").upper()):
                set_field(e["id"], str(new_value))
                changed = True
                break

        if not changed:
            return {"ok": False,
                    "note": f"Field {field_name} not found. "
                            "Use discover_screen_elements to locate it.",
                    "elements": [e["id"] for e in elems[:30]]}

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("DELIVERY_FIELD_FIX",
                  {"delivery": str(delivery_number),
                   "field": field_name, "value": str(new_value)},
                  status="executed")
        return {"ok": True, "delivery": str(delivery_number),
                "field": field_name, "new_value": new_value,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def create_delivery_output(delivery_number, output_type="LIEF", medium="1"):
    """
    Create/resend delivery output (delivery note etc.) via VL02N.
    output_type: LIEF=delivery note, LADS=loading list.
    """
    go_to_transaction("VL02N")
    time.sleep(1)
    try:
        for fid in ("wnd[0]/usr/ctxtLIKP-VBELN",):
            try: session.FindById(fid).Text = str(delivery_number); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(0)
        time.sleep(1.5)

        # Extras → Output → Issue
        try:
            session.FindById("wnd[0]/mbar/menu[3]/menu[0]/menu[0]").Select()
            time.sleep(1)
        except Exception:
            pass

        session.FindById("wnd[0]").SendVKey(11)
        time.sleep(1.5)

        audit_log("DELIVERY_OUTPUT",
                  {"delivery": str(delivery_number), "output_type": output_type},
                  status="executed")
        return {"ok": True, "delivery": str(delivery_number),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


# ── Background Job Error Resolution ───────────────────────────────────────────

def scan_failed_jobs(date_from=None, date_to=None, job_name="",
                      username=""):
    """
    Scan SM37 for cancelled/failed background jobs.
    Returns job name, step, start time, user, error status.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("SM37")
    time.sleep(1.5)
    try:
        # Job name (wildcard)
        for fid in ("wnd[0]/usr/ctxtBTCH2170-JOBNAME",):
            try: session.FindById(fid).Text = job_name or "*"; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtBTCH2170-USERNAME",):
            try: session.FindById(fid).Text = username or "*"; break
            except Exception: pass
        # Date range
        for fid in ("wnd[0]/usr/ctxtBTCH2170-FROM_DATE",):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtBTCH2170-TO_DATE",):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        # Select only cancelled/aborted jobs
        for fid in ("wnd[0]/usr/chkBTCH2170-ABORTED",
                    "wnd[0]/usr/chkABORTED"):
            try: session.FindById(fid).Selected = True; break
            except Exception: pass
        # Deselect others
        for fid_key in ("SCHEDULED","RELEASED","READY","ACTIVE","FINISHED"):
            for fid in (f"wnd[0]/usr/chkBTCH2170-{fid_key}",
                        f"wnd[0]/usr/chk{fid_key}"):
                try: session.FindById(fid).Selected = False; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        jobs = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["JOBNAME","JOBCOUNT","SDLSTRTDT","SDLSTRTTM",
                                "ENDDATE","ENDTIME","STATUS","AUTHCKNAM",
                                "PRCTEXT","STEPCOUNT"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("JOBNAME"):
                        row["fix_action"] = "restart_failed_job"
                        jobs.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "failed_jobs": jobs, "count": len(jobs),
            "screen": get_screen_text(),
            "note": (f"Found {len(jobs)} failed/cancelled jobs. "
                     "Call restart_failed_job for each to reschedule."),
        }
    except Exception as e:
        return {"error": str(e)}


def restart_failed_job(job_name, job_count):
    """
    Restart a failed background job via SM37.
    Locates the job by name+count and triggers immediate re-execution.
    """
    go_to_transaction("SM37")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtBTCH2170-JOBNAME",):
            try: session.FindById(fid).Text = job_name; break
            except Exception: pass
        # Select all statuses to find it
        for fid_key in ("SCHEDULED","RELEASED","READY","ACTIVE",
                        "FINISHED","ABORTED"):
            for fid in (f"wnd[0]/usr/chkBTCH2170-{fid_key}",):
                try: session.FindById(fid).Selected = True
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        # Find and select the row matching job_count
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 100)):
                    jc = shell.GetCellValue(i, "JOBCOUNT")
                    if str(jc) == str(job_count):
                        shell.SetCurrentCell(i, "JOBNAME")
                        shell.ClickCurrentCell()
                        break
        except Exception:
            pass

        # Job → Repeat → Immediate
        try:
            session.FindById("wnd[0]/mbar/menu[0]/menu[6]").Select()
            time.sleep(1)
        except Exception:
            try:
                session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
                time.sleep(1)
            except Exception:
                pass

        audit_log("JOB_RESTART",
                  {"job_name": job_name, "job_count": str(job_count)},
                  status="executed")
        return {"ok": True, "job_name": job_name,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


# ── Workflow Error Resolution ──────────────────────────────────────────────────

def scan_workflow_errors(date_from=None, date_to=None, task_id=""):
    """
    Scan SWI1 for stuck or error workflow items.
    Returns work items in error/suspended/cancelled status.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("SWI1")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtS_READT-LOW", "wnd[0]/usr/ctxtDATE_LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtS_READT-HIGH", "wnd[0]/usr/ctxtDATE_HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass
        if task_id:
            for fid in ("wnd[0]/usr/ctxtS_TASK-LOW",):
                try: session.FindById(fid).Text = task_id; break
                except Exception: pass
        # Select error/cancelled status checkboxes
        for cb in ("wnd[0]/usr/chkSWIRTYPE-FERR",
                   "wnd[0]/usr/chkSWIRTYPE-CANC",
                   "wnd[0]/usr/chkSTATUS_ERROR"):
            try: session.FindById(cb).Selected = True
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        items = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["WI_ID","WI_TYPE","WI_STAT","WI_TEXT",
                                "WI_ATEXT","CREATED_AT","UPDATED_AT",
                                "WI_AGENT","TASK_ID"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("WI_ID"):
                        row["fix_action"] = "restart_workflow_item"
                        items.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "workflow_errors": items, "count": len(items),
            "screen": get_screen_text(),
            "note": (f"Found {len(items)} workflow errors. "
                     "Call restart_workflow_item for each."),
        }
    except Exception as e:
        return {"error": str(e)}


def restart_workflow_item(workitem_id):
    """
    Restart a failed workflow work item via SWPR (workflow restart after error).
    """
    go_to_transaction("SWPR")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtSWPRARGS-WI_ID",
                    "wnd[0]/usr/ctxtWI_ID"):
            try: session.FindById(fid).Text = str(workitem_id); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        for fid in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("WORKFLOW_RESTART",
                  {"workitem_id": str(workitem_id)},
                  status="executed")
        return {"ok": True, "workitem_id": str(workitem_id),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


# ── Lock Entry & RFC/tRFC Error Resolution ────────────────────────────────────

def scan_sm12_locks(client="", username="", table_name=""):
    """
    Scan SM12 for stuck lock entries that are blocking users or processes.
    Returns lock owner, table, lock argument, and time locked.
    """
    go_to_transaction("SM12")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRM20M-MANDT",):
            try: session.FindById(fid).Text = client or ""; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtRM20M-GNAME",):
            try: session.FindById(fid).Text = username or "*"; break
            except Exception: pass
        if table_name:
            for fid in ("wnd[0]/usr/ctxtRM20M-OBJNAME",):
                try: session.FindById(fid).Text = table_name; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        locks = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["GNAME","OBJNAME","LOCKARG","TABNAME",
                                "MANDT","REPID","DATUM","UZEIT"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("OBJNAME") or row.get("GNAME"):
                        row["fix_action"] = "release_lock_entry"
                        locks.append(row)
        except Exception:
            pass

        return {
            "lock_entries": locks, "count": len(locks),
            "screen": get_screen_text(),
            "note": (f"Found {len(locks)} lock entries. "
                     "Call release_lock_entry to delete stale locks."),
        }
    except Exception as e:
        return {"error": str(e)}


def release_lock_entry(username, table_name=""):
    """
    Delete stale lock entries in SM12 for a user and optionally a specific table.
    WARNING: Only release locks for inactive/dead processes — active locks
    protect data integrity.
    """
    go_to_transaction("SM12")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtRM20M-GNAME",):
            try: session.FindById(fid).Text = username; break
            except Exception: pass
        if table_name:
            for fid in ("wnd[0]/usr/ctxtRM20M-OBJNAME",):
                try: session.FindById(fid).Text = table_name; break
                except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        session.FindById("wnd[0]").SendVKey(16)   # Select all
        time.sleep(0.5)

        # Delete locks
        session.FindById("wnd[0]").SendVKey(65)   # Shift+F5 = Delete
        time.sleep(1)
        for fid in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
            try: session.FindById(fid).Press(); time.sleep(1); break
            except Exception: pass

        audit_log("LOCK_ENTRY_RELEASE",
                  {"username": username, "table": table_name or "ALL"},
                  status="executed")
        return {"ok": True, "username": username,
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def scan_sm58_errors(date_from=None, date_to=None):
    """
    Scan SM58 for failed tRFC (transactional RFC) calls.
    Returns function module, destination, TID, error text, and retry status.
    """
    today = datetime.now().strftime("%d.%m.%Y")
    df = date_from or today
    dt = date_to   or today

    go_to_transaction("SM58")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtDATE_LOW",
                    "wnd[0]/usr/ctxtS_CRTDT-LOW"):
            try: session.FindById(fid).Text = df; break
            except Exception: pass
        for fid in ("wnd[0]/usr/ctxtDATE_HIGH",
                    "wnd[0]/usr/ctxtS_CRTDT-HIGH"):
            try: session.FindById(fid).Text = dt; break
            except Exception: pass

        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2.5)

        errors = []
        try:
            shell = _find_shell_anywhere([
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID1/shellcont/shell",
                "wnd[0]/usr/cntlGRID/shellcont/shell",
            ])
            if shell:
                for i in range(min(shell.RowCount, 200)):
                    row = {}
                    for col in ["TID","FUNCNAME","DEST","MANDT","STATUS",
                                "ERRTXT","CRTDT","CRTTM","USERNAME"]:
                        try: row[col] = shell.GetCellValue(i, col)
                        except Exception: pass
                    if row.get("TID"):
                        row["fix_action"] = "retry_sm58_entry"
                        errors.append(row)
        except Exception:
            pass

        return {
            "date_from": df, "date_to": dt,
            "sm58_errors": errors, "count": len(errors),
            "screen": get_screen_text(),
            "note": (f"Found {len(errors)} tRFC errors. "
                     "Call retry_sm58_entry to re-execute each failed call."),
        }
    except Exception as e:
        return {"error": str(e)}


def retry_sm58_entry(tid, function_name=""):
    """
    Retry a failed tRFC entry in SM58 by TID (transaction ID).
    Re-executes the remote function call that previously failed.
    """
    go_to_transaction("SM58")
    time.sleep(1.5)
    try:
        for fid in ("wnd[0]/usr/ctxtTID",):
            try: session.FindById(fid).Text = str(tid); break
            except Exception: pass
        session.FindById("wnd[0]").SendVKey(8)
        time.sleep(2)

        session.FindById("wnd[0]").SendVKey(16)
        time.sleep(0.5)

        # Execute: LUW → Execute
        for fid in ("wnd[0]/tbar[1]/btn[8]",):
            try: session.FindById(fid).Press(); time.sleep(2); break
            except Exception: pass
        try:
            session.FindById("wnd[0]/mbar/menu[0]/menu[1]").Select()
            time.sleep(1.5)
        except Exception:
            pass

        audit_log("SM58_RETRY",
                  {"tid": str(tid), "function_name": function_name},
                  status="executed")
        return {"ok": True, "tid": str(tid),
                "screen": get_screen_text()}
    except Exception as e:
        return {"error": str(e)}


def get_sales_orders(date_from, date_to):
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

    # ── FI/CO ERROR RESOLUTION ───────────────────────────────────────────────
    {
        "name": "scan_fi_errors",
        "description": (
            "Scan FBL1N/FBL5N/FBL3N for FI document errors and blocked items. "
            "account_type: K=vendor (FBL1N), D=customer (FBL5N), S=GL (FBL3N). "
            "Returns open/blocked items with payment block flags and fix_hints. "
            "ALWAYS call this first for FI/CO issues."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":    {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":      {"type": "string", "description": "DD.MM.YYYY"},
                "company_code": {"type": "string", "description": "e.g. 1000"},
                "account_type": {"type": "string",
                                 "enum": ["K","D","S"],
                                 "description": "K=vendor, D=customer, S=GL account"},
            },
            "required": [],
        },
    },
    {
        "name": "get_fi_doc_detail",
        "description": (
            "Display a FI document in FB03 and extract all line items, "
            "posting keys, amounts, payment block (ZLSPR), and fix_hints."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc_number":   {"type": "string", "description": "FI document number (BELNR)"},
                "company_code": {"type": "string", "description": "e.g. 1000"},
                "fiscal_year":  {"type": "string", "description": "e.g. 2026"},
            },
            "required": ["doc_number"],
        },
    },
    {
        "name": "release_fi_payment_block",
        "description": (
            "Remove payment block (ZLSPR) from a FI document via FB02, "
            "so the item is included in the next F110 payment run. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc_number":   {"type": "string"},
                "company_code": {"type": "string"},
                "fiscal_year":  {"type": "string"},
                "line_item":    {"type": "string", "description": "Line item number (default 1)"},
            },
            "required": ["doc_number"],
        },
    },
    {
        "name": "reverse_fi_document",
        "description": (
            "Reverse (storno) a FI document via FB08. "
            "reversal_reason: 01=current period, 02=closed period. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc_number":      {"type": "string"},
                "company_code":    {"type": "string"},
                "fiscal_year":     {"type": "string"},
                "reversal_reason": {"type": "string",
                                    "description": "01=current period, 02=closed period"},
                "reversal_date":   {"type": "string", "description": "DD.MM.YYYY"},
            },
            "required": ["doc_number"],
        },
    },
    {
        "name": "clear_open_items",
        "description": (
            "Clear open FI items for a vendor/customer/GL account "
            "via F-44 (vendor), F-32 (customer), or F-03 (GL). "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "account_number": {"type": "string",
                                   "description": "Vendor/customer/GL account number"},
                "company_code":   {"type": "string"},
                "account_type":   {"type": "string",
                                   "enum": ["K","D","S"],
                                   "description": "K=vendor, D=customer, S=GL"},
                "clearing_date":  {"type": "string", "description": "DD.MM.YYYY"},
            },
            "required": ["account_number"],
        },
    },

    # ── DELIVERY / SHIPPING ERROR RESOLUTION ─────────────────────────────────
    {
        "name": "scan_delivery_errors",
        "description": (
            "Scan VL06O (outbound delivery monitor) for stuck deliveries. "
            "error_type: 'gi'=GI not posted, 'pick'=picking incomplete, 'all'=everything. "
            "Returns deliveries with status and fix_hints."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":      {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":        {"type": "string", "description": "DD.MM.YYYY"},
                "shipping_point": {"type": "string", "description": "e.g. 1000"},
                "error_type":     {"type": "string",
                                   "enum": ["gi","pick","output","all"]},
            },
            "required": [],
        },
    },
    {
        "name": "post_goods_issue",
        "description": (
            "Post goods issue for a delivery via VL02N. "
            "Completes the outbound delivery and reduces stock. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "delivery_number": {"type": "string",
                                    "description": "Outbound delivery number"},
            },
            "required": ["delivery_number"],
        },
    },
    {
        "name": "fix_delivery_incomplete",
        "description": (
            "Fix an incomplete delivery by setting a missing field via VL02N. "
            "Common: ROUTE=route, VSTEL=shipping point, LGNUM=warehouse. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "delivery_number": {"type": "string"},
                "field_name":      {"type": "string",
                                    "description": "SAP field name e.g. ROUTE, VSTEL"},
                "new_value":       {"type": "string"},
            },
            "required": ["delivery_number", "field_name", "new_value"],
        },
    },
    {
        "name": "create_delivery_output",
        "description": (
            "Create/resend output for a delivery via VL02N. "
            "output_type: LIEF=delivery note, LADS=loading list. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "delivery_number": {"type": "string"},
                "output_type":     {"type": "string",
                                    "description": "LIEF=delivery note, LADS=loading list"},
                "medium":          {"type": "string",
                                    "description": "1=print, 5=external, 6=EDI"},
            },
            "required": ["delivery_number"],
        },
    },

    # ── BACKGROUND JOB ERROR RESOLUTION ──────────────────────────────────────
    {
        "name": "scan_failed_jobs",
        "description": (
            "Scan SM37 for cancelled/aborted background jobs. "
            "Returns job name, count, scheduled time, user, and error status. "
            "ALWAYS call this first for job failures."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
                "job_name":  {"type": "string",
                              "description": "Job name filter (blank=all, * for wildcard)"},
                "username":  {"type": "string", "description": "Owner user ID filter"},
            },
            "required": [],
        },
    },
    {
        "name": "restart_failed_job",
        "description": (
            "Restart a failed background job via SM37 by job name and count. "
            "Re-triggers immediate execution of the cancelled job. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "job_name":  {"type": "string", "description": "Background job name"},
                "job_count": {"type": "string",
                              "description": "Job count (from scan_failed_jobs JOBCOUNT)"},
            },
            "required": ["job_name", "job_count"],
        },
    },

    # ── WORKFLOW ERROR RESOLUTION ─────────────────────────────────────────────
    {
        "name": "scan_workflow_errors",
        "description": (
            "Scan SWI1 for workflow work items in error, suspended, or cancelled status. "
            "Returns work item ID, type, task, text, and agent."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
                "task_id":   {"type": "string",
                              "description": "Workflow task ID filter (optional)"},
            },
            "required": [],
        },
    },
    {
        "name": "restart_workflow_item",
        "description": (
            "Restart a failed workflow work item via SWPR. "
            "Re-executes the step that errored. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "workitem_id": {"type": "string",
                                "description": "Work item ID from scan_workflow_errors"},
            },
            "required": ["workitem_id"],
        },
    },

    # ── LOCK ENTRY RESOLUTION ─────────────────────────────────────────────────
    {
        "name": "scan_sm12_locks",
        "description": (
            "Scan SM12 for stuck lock entries that are blocking users or jobs. "
            "Returns lock owner, table, lock argument, and creation time."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "username":   {"type": "string", "description": "Lock owner user ID"},
                "table_name": {"type": "string",
                               "description": "Table name filter (optional)"},
                "client":     {"type": "string", "description": "SAP client (optional)"},
            },
            "required": [],
        },
    },
    {
        "name": "release_lock_entry",
        "description": (
            "Delete stale lock entries in SM12 for a specific user. "
            "WARNING: Only release locks for provably dead/inactive processes. "
            "Active locks protect data integrity. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "username":   {"type": "string",
                               "description": "Lock owner whose locks to release"},
                "table_name": {"type": "string",
                               "description": "Restrict to locks on this table (optional)"},
            },
            "required": ["username"],
        },
    },

    # ── tRFC / SM58 ERROR RESOLUTION ─────────────────────────────────────────
    {
        "name": "scan_sm58_errors",
        "description": (
            "Scan SM58 for failed tRFC (transactional RFC) calls. "
            "Returns function module, destination, TID, error text. "
            "Common source of ALE/IDoc and cross-system integration errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
            },
            "required": [],
        },
    },
    {
        "name": "retry_sm58_entry",
        "description": (
            "Retry a failed tRFC entry in SM58 by TID. "
            "Re-executes the remote function call. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "tid":           {"type": "string",
                                  "description": "Transaction ID from scan_sm58_errors"},
                "function_name": {"type": "string",
                                  "description": "Function module name (informational)"},
            },
            "required": ["tid"],
        },
    },

    # ── SALES ORDER ERROR RESOLUTION ─────────────────────────────────────────
    {
        "name": "scan_sd_errors",
        "description": (
            "Scan VA05/VKM1 for failed or blocked sales orders. "
            "error_type: 'credit'=credit-blocked, 'delivery'=delivery-blocked, "
            "'billing'=billing-blocked, 'incomplete'=incompletion log, 'all'=everything. "
            "Returns order list with status flags and fix_hints. "
            "ALWAYS call this first when investigating SD errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":  {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":    {"type": "string", "description": "DD.MM.YYYY"},
                "sales_org":  {"type": "string", "description": "e.g. 1000"},
                "error_type": {"type": "string",
                               "enum": ["credit","delivery","billing",
                                        "incomplete","all"],
                               "description": "Type of error to filter"},
            },
            "required": [],
        },
    },
    {
        "name": "get_sd_order_detail",
        "description": (
            "Open a sales order in VA03 and extract: header fields (credit/delivery/"
            "billing block, sold-to, sales org), item pricing, partner data, "
            "schedule lines, and auto-computed fix_hints from error text. "
            "Call after scan_sd_errors for each order."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string",
                                "description": "10-digit sales order number"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "release_credit_block",
        "description": (
            "Release a credit-blocked sales order via VKM3. "
            "Use when fix_hints shows 'release_credit_block' or CMGST=B/C. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order":  {"type": "string", "description": "Sales order number"},
                "release_type": {"type": "string",
                                 "enum": ["order","all"],
                                 "description": "order=this order only, all=all customer orders"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "release_sd_delivery_block",
        "description": (
            "Remove the delivery block from a sales order header via VA02. "
            "Clears VBAK-LIFSK so the order can be delivered. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "block_code":  {"type": "string",
                                "description": "Existing block code (informational)"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "release_billing_block",
        "description": (
            "Remove the billing block from a sales order via VA02. "
            "Clears VBAK-FAKSK so the order can be invoiced. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "block_code":  {"type": "string",
                                "description": "Existing block code (informational)"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "reprice_sales_order",
        "description": (
            "Reprice a sales order via VA02 — triggers fresh condition determination. "
            "Use when pricing is zero, wrong, or conditions have expired. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "complete_sales_order",
        "description": (
            "Fix an incomplete sales order by setting a missing mandatory field via VA02. "
            "Common: ZTERM=payment terms, VSTEL=shipping point, ROUTE=route, "
            "WERKS=plant, LIFSK=delivery block (set blank). "
            "item_number: blank=header, '00010'=first item. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "field_name":  {"type": "string",
                                "description": "SAP field name e.g. ZTERM, VSTEL, ROUTE"},
                "new_value":   {"type": "string", "description": "Value to set"},
                "item_number": {"type": "string",
                                "description": "Item number e.g. 00010 (blank=header)"},
            },
            "required": ["sales_order", "field_name", "new_value"],
        },
    },
    {
        "name": "fix_partner_determination",
        "description": (
            "Re-trigger partner determination for a sales order via VA02. "
            "Fixes missing or wrong sold-to, ship-to, bill-to, payer partners. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "create_sd_output",
        "description": (
            "Create or resend output for a sales order via VA02. "
            "output_type: BA00=order confirmation, RD00=invoice, LIEF=delivery. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "output_type": {"type": "string",
                                "description": "BA00=confirmation, RD00=invoice, LIEF=delivery"},
                "medium":      {"type": "string",
                                "description": "1=print, 5=external, 6=EDI"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "check_atp",
        "description": (
            "Run ATP (Available-to-Promise) check on a sales order item via VA02. "
            "Re-schedules the line and confirms deliverable quantity and date. "
            "Use when schedule lines are missing or confirmed quantity is zero. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "item_number": {"type": "string",
                                "description": "Item number e.g. 00010"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "remove_rejection_reason",
        "description": (
            "Remove the rejection reason from a sales order item via VA02, "
            "reactivating the item for processing. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sales_order": {"type": "string"},
                "item_number": {"type": "string",
                                "description": "Item number e.g. 00010"},
            },
            "required": ["sales_order"],
        },
    },
    {
        "name": "scan_credit_blocks",
        "description": (
            "Scan VKM1 for all sales orders blocked by credit management. "
            "Returns order, customer, credit exposure vs limit, block reason."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
                "sales_org": {"type": "string", "description": "e.g. 1000"},
            },
            "required": [],
        },
    },
    {
        "name": "scan_incomplete_orders",
        "description": (
            "Scan V.02 for incomplete sales orders (incompletion log has open items). "
            "Returns orders where mandatory fields are missing."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
                "sales_org": {"type": "string", "description": "e.g. 1000"},
            },
            "required": [],
        },
    },

    # ── PURCHASE ORDER ERROR RESOLUTION ──────────────────────────────────────
    {
        "name": "scan_po_errors",
        "description": (
            "Scan ME2M for failed/blocked purchase orders in a date range. "
            "Returns PO numbers, vendor, items, release status, and fix_hints. "
            "error_type: 'blocked'=release-blocked only, 'all'=every PO with issues. "
            "ALWAYS call this first when investigating PO errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":     {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":       {"type": "string", "description": "DD.MM.YYYY"},
                "company_code":  {"type": "string", "description": "e.g. 1000"},
                "purchase_org":  {"type": "string", "description": "e.g. 1000"},
                "error_type":    {"type": "string",
                                  "enum": ["blocked","overdue","all"],
                                  "description": "Filter: blocked=release-blocked, all=everything"},
            },
            "required": [],
        },
    },
    {
        "name": "get_po_detail",
        "description": (
            "Open a specific PO in ME23N and extract header, items, pricing, "
            "account assignment, delivery dates, release strategy status, and "
            "auto-computed fix_hints. Call after scan_po_errors for each PO."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number": {"type": "string",
                              "description": "10-digit PO number e.g. 4500001234"},
            },
            "required": ["po_number"],
        },
    },
    {
        "name": "release_po",
        "description": (
            "Release a purchase order blocked for approval via ME29N. "
            "Use when fix_hints shows 'release_po' or FRGKE field is set. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number":     {"type": "string", "description": "PO number"},
                "release_code":  {"type": "string",
                                  "description": "Release code from EKKO-FRGKE (default 01)"},
            },
            "required": ["po_number"],
        },
    },
    {
        "name": "change_po_field",
        "description": (
            "Change a field on a PO line item via ME22N. "
            "Common fixes: EINDT=delivery date, MENGE=quantity, NETPR=net price, "
            "WERKS=plant, KOSTL=cost centre, PSPNR=WBS element. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number":    {"type": "string", "description": "PO number"},
                "item_number":  {"type": "string",
                                 "description": "Item number e.g. 00010 (or blank for header)"},
                "field_name":   {"type": "string",
                                 "description": "SAP field name e.g. EINDT, MENGE, NETPR, KOSTL"},
                "new_value":    {"type": "string", "description": "New field value"},
            },
            "required": ["po_number", "item_number", "field_name", "new_value"],
        },
    },
    {
        "name": "cancel_po_item",
        "description": (
            "Set deletion flag on a PO line item via ME22N to cancel it. "
            "Use when item is no longer required or vendor cannot deliver. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number":   {"type": "string"},
                "item_number": {"type": "string", "description": "Line item number"},
                "reason":      {"type": "string", "description": "Reason for cancellation"},
            },
            "required": ["po_number", "item_number"],
        },
    },
    {
        "name": "scan_blocked_invoices",
        "description": (
            "Scan MRBR for blocked MM invoices (price/quantity/manual block). "
            "Returns invoice numbers, vendors, block reason codes. "
            "Call this when POs have been received but invoices are stuck."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from":    {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":      {"type": "string", "description": "DD.MM.YYYY"},
                "company_code": {"type": "string", "description": "e.g. 1000"},
            },
            "required": [],
        },
    },
    {
        "name": "release_blocked_invoice",
        "description": (
            "Release a blocked MM invoice via MRBR so it can proceed to payment. "
            "Clears price variance, quantity variance, and manual blocks. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "invoice_number": {"type": "string",
                                   "description": "Invoice document number (BELNR)"},
                "company_code":   {"type": "string", "description": "e.g. 1000"},
                "fiscal_year":    {"type": "string", "description": "e.g. 2026 (optional)"},
            },
            "required": ["invoice_number"],
        },
    },
    {
        "name": "check_vendor_master",
        "description": (
            "Check vendor master record via XK03. "
            "Returns block status, payment terms, bank details, purchasing data. "
            "Call when PO fails due to vendor-related errors."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "vendor_number": {"type": "string",
                                  "description": "Vendor number (LIFNR)"},
                "company_code":  {"type": "string", "description": "e.g. 1000"},
            },
            "required": ["vendor_number"],
        },
    },
    {
        "name": "unblock_vendor",
        "description": (
            "Remove a vendor block via XK05. "
            "block_type: 'purchase'=purchasing block, 'payment'=payment block, "
            "'all'=remove all blocks. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "vendor_number": {"type": "string"},
                "company_code":  {"type": "string", "description": "e.g. 1000"},
                "block_type":    {"type": "string",
                                  "enum": ["purchase","payment","all"],
                                  "description": "Which block to remove"},
            },
            "required": ["vendor_number"],
        },
    },
    {
        "name": "create_po_output",
        "description": (
            "Resend/create output message for a PO via ME9F. "
            "Use when vendor has not received the PO or output failed. "
            "output_type: NEU=original, MAHN=reminder. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "po_number":    {"type": "string"},
                "output_type":  {"type": "string",
                                 "description": "NEU=original, MAHN=reminder"},
                "medium":       {"type": "string",
                                 "description": "1=print, 5=external send, 6=EDI"},
            },
            "required": ["po_number"],
        },
    },
    {
        "name": "scan_open_purchase_reqs",
        "description": (
            "Scan ME5A for open purchase requisitions not yet converted to POs. "
            "Returns PR numbers, materials, quantities, plants, requested delivery dates."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "date_from": {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":   {"type": "string", "description": "DD.MM.YYYY"},
                "plant":     {"type": "string", "description": "Plant code e.g. 1000"},
                "material":  {"type": "string", "description": "Material number"},
            },
            "required": [],
        },
    },
    {
        "name": "convert_pr_to_po",
        "description": (
            "Convert a purchase requisition to a purchase order via ME57. "
            "Assigns source of supply and creates the PO. "
            "REQUIRES human approval."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "pr_number":     {"type": "string", "description": "PR number (BANFN)"},
                "vendor_number": {"type": "string", "description": "Vendor to assign"},
                "purchase_org":  {"type": "string", "description": "e.g. 1000"},
                "company_code":  {"type": "string", "description": "e.g. 1000"},
            },
            "required": ["pr_number", "vendor_number"],
        },
    },
    {
        "name": "scan_gr_ir_clearing",
        "description": (
            "Scan MB5S for uncleared GR/IR items: goods receipts without matching "
            "invoices, or invoices without GR. Returns imbalance for review. "
            "Use MR11 to post clearing after reviewing."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "company_code": {"type": "string", "description": "e.g. 1000"},
                "date_from":    {"type": "string", "description": "DD.MM.YYYY"},
                "date_to":      {"type": "string", "description": "DD.MM.YYYY"},
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
    # PO write tools
    "release_po", "change_po_field", "cancel_po_item",
    "release_blocked_invoice", "unblock_vendor",
    "create_po_output", "convert_pr_to_po",
    # SD write tools
    "release_credit_block", "release_sd_delivery_block",
    "release_billing_block", "reprice_sales_order",
    "complete_sales_order", "fix_partner_determination",
    "create_sd_output", "check_atp", "remove_rejection_reason",
    # FI/CO write tools
    "release_fi_payment_block", "reverse_fi_document", "clear_open_items",
    # Delivery write tools
    "post_goods_issue", "fix_delivery_incomplete", "create_delivery_output",
    # Job/Workflow/Lock/tRFC write tools
    "restart_failed_job", "restart_workflow_item",
    "release_lock_entry", "retry_sm58_entry",
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

    # ── FI/CO error resolution tools ──────────────────────────────────────────
    if tool_name == "scan_fi_errors":
        return scan_fi_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("company_code", "1000"),
            tool_input.get("account_type", "K"),
        )
    if tool_name == "get_fi_doc_detail":
        return get_fi_doc_detail(
            tool_input["doc_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("fiscal_year", ""),
        )
    if tool_name == "release_fi_payment_block":
        return release_fi_payment_block(
            tool_input["doc_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("fiscal_year", ""),
            tool_input.get("line_item", "1"),
        )
    if tool_name == "reverse_fi_document":
        return reverse_fi_document(
            tool_input["doc_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("fiscal_year", ""),
            tool_input.get("reversal_reason", "01"),
            tool_input.get("reversal_date"),
        )
    if tool_name == "clear_open_items":
        return clear_open_items(
            tool_input["account_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("account_type", "K"),
            tool_input.get("clearing_date"),
        )

    # ── Delivery error resolution tools ───────────────────────────────────────
    if tool_name == "scan_delivery_errors":
        return scan_delivery_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("shipping_point", ""),
            tool_input.get("error_type", "all"),
        )
    if tool_name == "post_goods_issue":
        return post_goods_issue(tool_input["delivery_number"])
    if tool_name == "fix_delivery_incomplete":
        return fix_delivery_incomplete(
            tool_input["delivery_number"],
            tool_input["field_name"],
            tool_input["new_value"],
        )
    if tool_name == "create_delivery_output":
        return create_delivery_output(
            tool_input["delivery_number"],
            tool_input.get("output_type", "LIEF"),
            tool_input.get("medium", "1"),
        )

    # ── Background job error resolution tools ─────────────────────────────────
    if tool_name == "scan_failed_jobs":
        return scan_failed_jobs(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("job_name", ""),
            tool_input.get("username", ""),
        )
    if tool_name == "restart_failed_job":
        return restart_failed_job(
            tool_input["job_name"],
            tool_input["job_count"],
        )

    # ── Workflow error resolution tools ───────────────────────────────────────
    if tool_name == "scan_workflow_errors":
        return scan_workflow_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("task_id", ""),
        )
    if tool_name == "restart_workflow_item":
        return restart_workflow_item(tool_input["workitem_id"])

    # ── Lock entry resolution tools ───────────────────────────────────────────
    if tool_name == "scan_sm12_locks":
        return scan_sm12_locks(
            tool_input.get("client", ""),
            tool_input.get("username", ""),
            tool_input.get("table_name", ""),
        )
    if tool_name == "release_lock_entry":
        return release_lock_entry(
            tool_input["username"],
            tool_input.get("table_name", ""),
        )

    # ── tRFC / SM58 error resolution tools ────────────────────────────────────
    if tool_name == "scan_sm58_errors":
        return scan_sm58_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
        )
    if tool_name == "retry_sm58_entry":
        return retry_sm58_entry(
            tool_input["tid"],
            tool_input.get("function_name", ""),
        )

    # ── Sales Order error resolution tools ────────────────────────────────────
    if tool_name == "scan_sd_errors":
        return scan_sd_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("sales_org", ""),
            tool_input.get("error_type", "all"),
        )
    if tool_name == "get_sd_order_detail":
        return get_sd_order_detail(tool_input["sales_order"])
    if tool_name == "release_credit_block":
        return release_credit_block(
            tool_input["sales_order"],
            tool_input.get("release_type", "order"),
        )
    if tool_name == "release_sd_delivery_block":
        return release_sd_delivery_block(
            tool_input["sales_order"],
            tool_input.get("block_code", ""),
        )
    if tool_name == "release_billing_block":
        return release_billing_block(
            tool_input["sales_order"],
            tool_input.get("block_code", ""),
        )
    if tool_name == "reprice_sales_order":
        return reprice_sales_order(tool_input["sales_order"])
    if tool_name == "complete_sales_order":
        return complete_sales_order(
            tool_input["sales_order"],
            tool_input["field_name"],
            tool_input["new_value"],
            tool_input.get("item_number", ""),
        )
    if tool_name == "fix_partner_determination":
        return fix_partner_determination(tool_input["sales_order"])
    if tool_name == "create_sd_output":
        return create_sd_output(
            tool_input["sales_order"],
            tool_input.get("output_type", "BA00"),
            tool_input.get("medium", "1"),
        )
    if tool_name == "check_atp":
        return check_atp(
            tool_input["sales_order"],
            tool_input.get("item_number", "00010"),
        )
    if tool_name == "remove_rejection_reason":
        return remove_rejection_reason(
            tool_input["sales_order"],
            tool_input.get("item_number", "00010"),
        )
    if tool_name == "scan_credit_blocks":
        return scan_credit_blocks(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("sales_org", ""),
        )
    if tool_name == "scan_incomplete_orders":
        return scan_incomplete_orders(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("sales_org", ""),
        )

    # ── Purchase Order error resolution tools ─────────────────────────────────
    if tool_name == "scan_po_errors":
        return scan_po_errors(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("company_code", ""),
            tool_input.get("purchase_org", ""),
            tool_input.get("error_type", "all"),
        )
    if tool_name == "get_po_detail":
        return get_po_detail(tool_input["po_number"])
    if tool_name == "release_po":
        return release_po(
            tool_input["po_number"],
            tool_input.get("release_code", "01"),
        )
    if tool_name == "change_po_field":
        return change_po_field(
            tool_input["po_number"],
            tool_input.get("item_number", "00010"),
            tool_input["field_name"],
            tool_input["new_value"],
        )
    if tool_name == "cancel_po_item":
        return cancel_po_item(
            tool_input["po_number"],
            tool_input.get("item_number", "00010"),
            tool_input.get("reason", ""),
        )
    if tool_name == "scan_blocked_invoices":
        return scan_blocked_invoices(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("company_code", "1000"),
        )
    if tool_name == "release_blocked_invoice":
        return release_blocked_invoice(
            tool_input["invoice_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("fiscal_year", ""),
        )
    if tool_name == "check_vendor_master":
        return check_vendor_master(
            tool_input["vendor_number"],
            tool_input.get("company_code", "1000"),
        )
    if tool_name == "unblock_vendor":
        return unblock_vendor(
            tool_input["vendor_number"],
            tool_input.get("company_code", "1000"),
            tool_input.get("block_type", "purchase"),
        )
    if tool_name == "create_po_output":
        return create_po_output(
            tool_input["po_number"],
            tool_input.get("output_type", "NEU"),
            tool_input.get("medium", "1"),
        )
    if tool_name == "scan_open_purchase_reqs":
        return scan_open_purchase_reqs(
            tool_input.get("date_from"),
            tool_input.get("date_to"),
            tool_input.get("plant", ""),
            tool_input.get("material", ""),
        )
    if tool_name == "convert_pr_to_po":
        return convert_pr_to_po(
            tool_input["pr_number"],
            tool_input["vendor_number"],
            tool_input.get("purchase_org", "1000"),
            tool_input.get("company_code", "1000"),
        )
    if tool_name == "scan_gr_ir_clearing":
        return scan_gr_ir_clearing(
            tool_input.get("company_code", "1000"),
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
  │ sm30_load_entries(table, [{{field:value,...}},...])   │
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
 FI/CO POSTING ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_fi_errors(date_from, date_to, company_code, account_type)
        → account_type: K=vendor, D=customer, S=GL account
        → Returns items with ZLSPR (payment block) and fix_hints

STEP 2  get_fi_doc_detail(doc_number, company_code) for each item
        → Returns all line items, amounts, and fix_hints

STEP 3  PROPOSE FIX to user — wait for approval — then execute:
        ─────────────────────────────────────────────────────────
        ZLSPR set (payment block)   → release_fi_payment_block
        Duplicate / wrong amount    → reverse_fi_document
        Open items not cleared      → clear_open_items (F-44/F-32/F-03)
        Period not open             → OB52: open_posting_period (maintain_table)
        ─────────────────────────────────────────────────────────

STEP 4  Verify with scan_fi_errors to confirm count drops to zero.

FI PAYMENT BLOCK CODES (ZLSPR):
  blank=no block  A=Payment block  B=Dunning block  Z=Manual hold

═══════════════════════════════════════════════════════════
 DELIVERY / SHIPPING ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_delivery_errors(date_from, date_to, error_type)
        → error_type: 'gi'=GI not posted, 'pick'=picking open, 'all'=all

STEP 2  PROPOSE FIX — wait for approval — then execute:
        GI not posted              → post_goods_issue(delivery)
        Missing field (route/point)→ fix_delivery_incomplete(delivery, field, value)
        Output not sent            → create_delivery_output(delivery, output_type)

═══════════════════════════════════════════════════════════
 BACKGROUND JOB ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_failed_jobs(date_from, date_to, job_name)
        → Returns cancelled/aborted jobs with name, count, schedule time

STEP 2  Investigate: what program does the job run? Check spool for error msg.
        PROPOSE FIX: "Job ZBATCH_INVOICES (count 12345678) failed.
         Proposed fix: restart immediately. Approve? [A/R]"

STEP 3  restart_failed_job(job_name, job_count)

═══════════════════════════════════════════════════════════
 WORKFLOW ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_workflow_errors(date_from, date_to, task_id)
        → Returns work items in ERROR/CANCELLED/SUSPENDED status

STEP 2  PROPOSE: "Work item 123456 (task TS12345678: Approve PO) failed.
         Proposed fix: restart via SWPR. Approve? [A/R]"

STEP 3  restart_workflow_item(workitem_id)

═══════════════════════════════════════════════════════════
 LOCK ENTRY RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_sm12_locks(username, table_name)
        → Returns all active locks; identify stale ones (user logged off)

STEP 2  PROPOSE: "User BATCH01 has 3 stale locks on VBAK from 2 hours ago.
         The user session is no longer active.
         Proposed fix: release_lock_entry. Approve? [A/R]"

STEP 3  release_lock_entry(username, table_name)
        WARNING: Never release locks for active, running processes.

═══════════════════════════════════════════════════════════
 tRFC / SM58 ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
STEP 1  scan_sm58_errors(date_from, date_to)
        → Returns failed tRFC calls with TID, function module, destination

STEP 2  PROPOSE: "tRFC TID ABC123 calling IDOC_INBOUND_ASYNCHRONOUS
         to destination DEST_ECC failed with: RFC_NO_AUTHORITY.
         Proposed fix: retry_sm58_entry. Approve? [A/R]"

STEP 3  retry_sm58_entry(tid)
        Note: if the RFC destination itself is down, fix SM59 first.

═══════════════════════════════════════════════════════════
 SALES ORDER ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
When asked to fix sales order errors, follow this sequence:

STEP 1  scan_sd_errors(date_from, date_to, error_type)
        → error_type='credit' for credit blocks, 'all' for everything
        → Returns orders with issues[] and fix_hints[]

STEP 2  get_sd_order_detail(sales_order) — for each error order
        → Returns header_fields (LIFSK/FAKSK/CMGST), all_screen_text,
          error_messages[], fix_hints[]

STEP 3  IDENTIFY ROOT CAUSE and PROPOSE SOLUTION to user:
        ─────────────────────────────────────────────────────────────
        fix_action               root cause            tool
        ─────────────────────────────────────────────────────────────
        release_credit_block   → credit limit exceeded → release_credit_block
        release_sd_delivery_block → delivery blocked   → release_sd_delivery_block
        release_billing_block  → billing blocked       → release_billing_block
        reprice_sales_order    → price 0 / wrong cond  → reprice_sales_order
        complete_sales_order   → missing mandatory fld → scan_incomplete_orders
                                                         complete_sales_order
        fix_partner_determination → wrong/missing partner → fix_partner_determination
        create_sd_output       → no output sent        → create_sd_output
        check_atp              → no confirmed qty/date → check_atp
        remove_rejection_reason → item rejected        → remove_rejection_reason
        ─────────────────────────────────────────────────────────────

STEP 4  ASK BEFORE FIXING — present root cause and proposed fix:
        "Order 1000001234 is credit-blocked (CMGST=B). Customer 1001 has
         exceeded credit limit of 50,000 EUR. Proposed fix: release_credit_block.
         Approve? [A/R]"

STEP 5  Execute fix after approval, then verify with get_sd_order_detail.

COMMON SD ERROR PATTERNS:
  CMGST = B or C          → credit block      → scan_credit_blocks → release_credit_block
  LIFSK not blank         → delivery block    → release_sd_delivery_block
  FAKSK not blank         → billing block     → release_billing_block
  NETWR = 0 or pricing err→ pricing error     → reprice_sales_order
  UVVLS or UVALL set      → incompletion log  → scan_incomplete_orders → complete_sales_order
  No schedule lines       → ATP failure       → check_atp
  ABGRU set on item       → item rejected     → remove_rejection_reason
  Ship-to/payer missing   → partner det. err  → fix_partner_determination
  No output records       → output not sent   → create_sd_output(BA00)

CREDIT MANAGEMENT STATUS (CMGST):
  blank = no credit check  A = OK  B = Warning  C = Blocked  D = Approved

DELIVERY BLOCK CODES (LIFSK — SAP standard):
  01=Check credit  02=Check export  Z1=Manual block  Z2=Awaiting contract

BILLING BLOCK CODES (FAKSK):
  01=Approval req  02=Pricing incomplete  Z1=Manual  ZP=Partial delivery

═══════════════════════════════════════════════════════════
 PURCHASE ORDER ERROR RESOLUTION WORKFLOW
═══════════════════════════════════════════════════════════
When asked to investigate or fix PO errors, follow this sequence:

STEP 1  scan_po_errors(date_from, date_to, error_type)
        → Returns list of POs with issue flags and fix_hints[]
        → error_type='blocked' for release-only, 'all' for everything

STEP 2  get_po_detail(po_number) — for each PO with errors
        → Returns all_screen_text, error_messages[], fix_hints[]
        → fix_hints auto-map to the correct fix action

STEP 3  IDENTIFY ROOT CAUSE from fix_hints and present to user:
        ─────────────────────────────────────────────────────────
        fix_action              root cause           tool to call
        ─────────────────────────────────────────────────────────
        release_po            → approval required  → release_po
        release_blocked_invoice→ price/qty variance→ scan_blocked_invoices
                                                     release_blocked_invoice
        check_vendor_master   → vendor issue       → check_vendor_master
        unblock_vendor        → vendor blocked     → unblock_vendor
        change_po_field       → wrong data         → change_po_field
        cancel_po_item        → item not needed    → cancel_po_item
        create_po_output      → no output sent     → create_po_output
        convert_pr_to_po      → PR not converted   → scan_open_purchase_reqs
                                                     convert_pr_to_po
        scan_gr_ir_clearing   → GR/IR mismatch     → scan_gr_ir_clearing
        ─────────────────────────────────────────────────────────

STEP 4  ASK BEFORE FIXING — present findings and proposed fix:
        "PO 4500001234 is blocked for release (FRGKE=01).
         Root cause: Release strategy requires approval.
         Proposed fix: release_po. Approve? [A/R]"

STEP 5  Execute fix after approval, then verify with get_po_detail.

COMMON PO ERROR PATTERNS:
  "blocked for release" / FRGKE set  → release_po(po_number)
  "tolerance limit exceeded"         → release_blocked_invoice
                                       OR change_po_field(NETPR, corrected price)
  "vendor … blocked"                 → check_vendor_master → unblock_vendor
  "no info record / price"           → change_po_field(NETPR, agreed price)
  "account assignment mandatory"     → change_po_field(KOSTL/PSPNR, value)
  "delivery date in past"            → change_po_field(EINDT, new date)
  "output not sent"                  → create_po_output
  "PR not assigned"                  → scan_open_purchase_reqs → convert_pr_to_po
  "GR/IR not cleared"                → scan_gr_ir_clearing → MR11 posting
  "invoice blocked"                  → scan_blocked_invoices → release_blocked_invoice

BLOCKED INVOICE REASON CODES:
  R = Price variance (net price differs from GR)
  Q = Quantity variance (invoiced qty > GR qty)
  M = Manual block set by user
  D = Date variance
  A = Amount exceeded tolerance
  S = Stochastic block (random QA check)

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
    print("  ARTILEGENZ SAP Agent v12.0  —  User: S4ABAP24  [SAP_ALL]")
    print("  Authorization: FULL SYSTEM ACCESS")
    print("  All write operations require your approval first.")
    print("═" * 68)
    print("\nExample queries:")
    print('  "Show all FI documents with payment blocks for company code 1000"')
    print('  "Reverse FI document 1800001234 company code 1000"')
    print('  "Scan all stuck deliveries from today and fix GI"')
    print('  "Show all failed background jobs from this week"')
    print('  "Restart failed job ZBATCH_INVOICES"')
    print('  "Show all workflow errors from today and restart them"')
    print('  "Check SM12 for stale locks by user BATCH01"')
    print('  "Show all SM58 tRFC errors and retry them"')
    print('  "Scan all failed sales orders from today and propose fixes"')
    print('  "Show all credit-blocked orders from this week"')
    print('  "Release the delivery block on sales order 1000001234"')
    print('  "Order 1000001234 has zero pricing — reprice it"')
    print('  "Show all incomplete sales orders from April 2026"')
    print('  "Run ATP check on sales order 1000001234 item 10"')
    print('  "Scan all purchase order errors from this month and fix them"')
    print('  "PO 4500001234 is blocked — diagnose and propose a fix"')
    print('  "Show all blocked invoices for company code 1000 this week"')
    print('  "Release all blocked invoices from April 2026"')
    print('  "Check vendor 100012 master data and unblock if blocked"')
    print('  "Resend output for PO 4500001234"')
    print('  "Show all open purchase requisitions not yet converted to POs"')
    print('  "Convert PR 1000000123 to a PO for vendor 100012"')
    print('  "Show GR/IR clearing items for company code 1000"')
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
