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
]

# ── Write operations that need approval ───────────────────────────────────────
WRITE_TOOLS = {
    "set_field_value", "press_button", "send_vkey",
    "handle_popup", "handle_transport_request",
    "maintain_table", "execute_abap_program",
    "create_transport_request", "release_transport_request",
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
    print("  ARTILEGENZ SAP Agent v5.0  —  User: S4ABAP24  [SAP_ALL]")
    print("  Authorization: FULL SYSTEM ACCESS")
    print("  All write operations require your approval first.")
    print("═" * 68)
    print("\nExample queries:")
    print('  "Read table T001 and show all existing company codes"')
    print('  "Create the full IDES org structure with transports"')
    print('  "Create company code 1000 IDES AG Germany EUR"')
    print('  "Create plant 1000 Hamburg Germany"')
    print('  "Create sales org 1000 Deutschland assign to company code 1000"')
    print('  "Create a transport for all org structure changes"')
    print('  "Show all sales orders from January 2025"')
    print('  "Fix all failed IDocs from today"')
    print('  "Show ABAP dumps from ST22 and propose fixes"')
    print('  "Maintain table T001W and add plant 2000 Berlin"')

    while True:
        query = input("\nQuery (or exit): ").strip()
        if query.lower() in ("exit", "quit", "q", ""):
            break
        run_agent(query, API_KEY)
