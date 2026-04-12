"""
ARTILEGENZ SAP Claude Agent v5.0
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
    print("  ARTILEGENZ SAP Agent v7.0  —  User: S4ABAP24  [SAP_ALL]")
    print("  Authorization: FULL SYSTEM ACCESS")
    print("  All write operations require your approval first.")
    print("═" * 68)
    print("\nExample queries:")
    print('  "Create an ABAP report that lists all open sales orders with ALV"')
    print('  "Create a function module to validate customer credit limit"')
    print('  "Create a custom Z-table to log all AI changes with timestamp"')
    print('  "Write an ABAP program to reprocess all failed IDocs from today"')
    print('  "Create a background job program that emails daily error summary"')
    print('  "Scan all ABAP dumps from today and fix them"')
    print('  "Fix the ABAP dump in program SAPMV45A"')
    print('  "Create the full IDES org structure with transports"')
    print('  "Show all sales orders from January 2025"')

    while True:
        query = input("\nQuery (or exit): ").strip()
        if query.lower() in ("exit", "quit", "q", ""):
            break
        run_agent(query, API_KEY)
