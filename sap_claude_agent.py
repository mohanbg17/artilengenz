"""
ARTILEGENZ SAP Claude Agent v3.0
Supports READ + WRITE operations with human approval guardrails.
Handles transport requests, popups, field discovery.

Requirements:
    pip install anthropic pywin32
Run from WinPython Command Prompt with SAP GUI open and logged in.
"""

import win32com.client
import anthropic
import time
import json
import os
import getpass
from datetime import datetime

# ── SAP Connection ────────────────────────────────────────────────────────────
def connect_sap():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    app = SapGuiAuto.GetScriptingEngine
    conn = app.Children(0)
    sess = conn.Children(0)
    print(f"Connected to SAP: {conn.Description}")
    return sess

session = connect_sap()

CURRENT_USER = getpass.getuser()
LOG_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "SAP_Agent_Logs")
os.makedirs(LOG_DIR, exist_ok=True)

# ── Audit Log ─────────────────────────────────────────────────────────────────
audit_entries = []

def audit_log(action, details, status="completed", approved_by=None):
    entry = {
        "timestamp": datetime.now().isoformat(),
        "user": approved_by or CURRENT_USER,
        "action": action,
        "details": details,
        "status": status,
    }
    audit_entries.append(entry)
    log_path = os.path.join(LOG_DIR, f"audit_{datetime.now().strftime('%Y%m%d')}.json")
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(audit_entries, f, indent=2, ensure_ascii=False)
    return entry

# ── Human Approval ────────────────────────────────────────────────────────────
def ask_human_approval(action_description, details, risk="medium"):
    """
    Always called before any write operation.
    Returns True (approved) or False (rejected) + optional reason.
    """
    risk_colors = {"low": "[ LOW ]", "medium": "[MEDIUM]", "high": "[ HIGH ]"}
    print("\n" + "═" * 65)
    print(f"  APPROVAL REQUIRED  {risk_colors.get(risk, '[?????]')}")
    print("═" * 65)
    print(f"  Action : {action_description}")
    print(f"  Details: {json.dumps(details, indent=4)}")
    print("═" * 65)
    print("  A = Approve    R = Reject")
    print("─" * 65)

    while True:
        choice = input("  Your decision [A/R]: ").strip().upper()
        if choice == "A":
            audit_log(action_description, details, status="approved", approved_by=CURRENT_USER)
            return True
        elif choice == "R":
            reason = input("  Rejection reason (optional): ").strip()
            audit_log(action_description, {**details, "rejection_reason": reason},
                      status="rejected", approved_by=CURRENT_USER)
            return False

# ── SAP Screen Helpers ────────────────────────────────────────────────────────
def go_to_transaction(tcode):
    session.StartTransaction(tcode)
    time.sleep(1.5)

def get_screen_text():
    """Read status bar + window title from current SAP screen."""
    try:
        wnd = session.FindById("wnd[0]")
        title = wnd.Text
        try:
            sbar = session.FindById("wnd[0]/sbar").Text
        except Exception:
            sbar = ""
        return f"Screen: {title}\nStatus: {sbar}"
    except Exception as e:
        return f"Error reading screen: {e}"

def discover_elements():
    """
    Walk the current SAP screen and return every interactable element
    with its ID, type, and current value.  Claude uses this to learn
    which field IDs exist before trying to fill them.
    """
    results = []

    def walk(component, depth=0):
        if depth > 6:
            return
        try:
            count = component.Children.Count
        except Exception:
            return
        for i in range(count):
            try:
                child = component.Children(i)
                try:
                    elem_id   = child.Id
                    elem_type = child.Type
                    try:
                        elem_text = child.Text
                    except Exception:
                        elem_text = ""
                    try:
                        tooltip = child.Tooltip
                    except Exception:
                        tooltip = ""
                    if elem_type in (
                        "GuiTextField", "GuiCTextField", "GuiComboBox",
                        "GuiRadioButton", "GuiCheckBox", "GuiButton",
                        "GuiTab", "GuiMenubar", "GuiStatusbar",
                    ):
                        results.append({
                            "id": elem_id,
                            "type": elem_type,
                            "value": elem_text,
                            "tooltip": tooltip,
                        })
                except Exception:
                    pass
                walk(child, depth + 1)
            except Exception:
                pass

    try:
        walk(session.FindById("wnd[0]"))
    except Exception as e:
        results.append({"error": str(e)})

    return results

def set_field_value(element_id, value):
    """Set the text value of a SAP field by its element ID."""
    try:
        field = session.FindById(element_id)
        field.Text = str(value)
        field.SetFocus()
        return {"success": True, "element": element_id, "value": value}
    except Exception as e:
        return {"success": False, "error": str(e), "element": element_id}

def press_button(element_id):
    """Press a SAP button or toolbar button by element ID."""
    try:
        btn = session.FindById(element_id)
        btn.Press()
        time.sleep(1)
        return {"success": True, "pressed": element_id, "screen": get_screen_text()}
    except Exception as e:
        return {"success": False, "error": str(e), "element": element_id}

def send_vkey(key_code):
    """
    Send a virtual key to the active SAP window.
    Common codes: 0=Enter, 3=F3/Back, 8=F8/Execute, 11=Save(F11),
                  70=Ctrl+S, 16=F16, 31=Ctrl+P
    """
    try:
        session.FindById("wnd[0]").SendVKey(int(key_code))
        time.sleep(1)
        return {"success": True, "vkey": key_code, "screen": get_screen_text()}
    except Exception as e:
        return {"success": False, "error": str(e)}

def handle_popup(action="confirm", text_input=None):
    """
    Handle SAP popup dialogs.
    action: 'confirm'  → press Enter / OK (VKey 0)
            'cancel'   → press Cancel    (VKey 12)
            'yes'      → press Yes button
            'no'       → press No button
            'input'    → enter text_input then confirm
    """
    try:
        wnd1 = session.FindById("wnd[1]", False)
    except Exception:
        wnd1 = None

    if wnd1 is None:
        return {"success": False, "error": "No popup found"}

    try:
        popup_text = wnd1.Text
    except Exception:
        popup_text = ""

    try:
        if action == "input" and text_input is not None:
            try:
                field = wnd1.FindById("usr/ctxtDY_FILENAME", False) or \
                        wnd1.FindById("usr/txtDY_FILENAME", False)
                if field:
                    field.Text = str(text_input)
            except Exception:
                pass
            wnd1.SendVKey(0)
        elif action in ("confirm", "yes"):
            wnd1.SendVKey(0)
        elif action in ("cancel", "no"):
            wnd1.SendVKey(12)
        time.sleep(0.8)
        return {"success": True, "popup_title": popup_text, "action": action,
                "screen": get_screen_text()}
    except Exception as e:
        return {"success": False, "error": str(e)}

def handle_transport_popup(transport_number=None):
    """
    Respond to SAP's 'Workbench/Customising Request' popup.
    If transport_number is given, assigns to that request.
    Otherwise creates a new transport.
    Returns the transport number used.
    """
    try:
        wnd1 = session.FindById("wnd[1]", False)
        if wnd1 is None:
            # Try pressing Save first to trigger the popup
            session.FindById("wnd[0]").SendVKey(11)
            time.sleep(1.5)
            wnd1 = session.FindById("wnd[1]", False)

        if wnd1 is None:
            return {"success": True, "note": "No transport popup appeared (may be local)"}

        popup_title = wnd1.Text

        if transport_number:
            # Assign to existing transport
            try:
                field = wnd1.FindById("usr/ctxtKORR-TRKORR", False)
                if field:
                    field.Text = transport_number
            except Exception:
                pass
            wnd1.SendVKey(0)   # Enter / OK
            return {"success": True, "transport": transport_number,
                    "action": "assigned_existing", "popup": popup_title}
        else:
            # Accept default (creates new transport or uses workbench default)
            wnd1.SendVKey(0)
            time.sleep(1)
            # Try to read the transport number from status bar
            try:
                status = session.FindById("wnd[0]/sbar").Text
            except Exception:
                status = ""
            return {"success": True, "transport": "NEW",
                    "action": "created_new", "status": status}

    except Exception as e:
        return {"success": False, "error": str(e)}

# ── SAP Read Functions (existing, unchanged) ──────────────────────────────────
def run_va05_and_extract(date_from="01.01.2025", date_to="31.12.2026"):
    go_to_transaction("VA05")
    try:
        session.FindById("wnd[0]/usr/ctxtSD_VBAK-AUDAT_LOW").Text  = date_from
        session.FindById("wnd[0]/usr/ctxtSD_VBAK-AUDAT_HIGH").Text = date_to
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        time.sleep(2)
    except Exception as e:
        return f"VA05 error: {e}"
    return get_screen_text()

def run_mb51_and_extract():
    go_to_transaction("MB51")
    try:
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        time.sleep(2)
    except Exception as e:
        return f"MB51 error: {e}"
    return get_screen_text()

# ── Claude Tool Definitions ───────────────────────────────────────────────────
TOOLS = [
    {
        "name": "go_to_transaction",
        "description": "Navigate to any SAP transaction code (read or write).",
        "input_schema": {
            "type": "object",
            "properties": {
                "tcode": {"type": "string", "description": "e.g. VA02, OX02, SPRO, SE09"},
            },
            "required": ["tcode"],
        },
    },
    {
        "name": "discover_screen_elements",
        "description": (
            "Scan the current SAP screen and return all interactable elements "
            "with their IDs, types and current values. Call this after navigating "
            "to a transaction so you know which field IDs to use."
        ),
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "set_field_value",
        "description": (
            "Set the value of a SAP input field by its exact element ID. "
            "Always call discover_screen_elements first to get valid IDs."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "element_id": {
                    "type": "string",
                    "description": "Full SAP element path e.g. wnd[0]/usr/ctxtBUKRS",
                },
                "value": {"type": "string", "description": "Value to enter"},
            },
            "required": ["element_id", "value"],
        },
    },
    {
        "name": "press_button",
        "description": "Press a SAP button or toolbar button by its element ID.",
        "input_schema": {
            "type": "object",
            "properties": {
                "element_id": {
                    "type": "string",
                    "description": "e.g. wnd[0]/tbar[0]/btn[0] or wnd[0]/tbar[1]/btn[8]",
                },
            },
            "required": ["element_id"],
        },
    },
    {
        "name": "send_vkey",
        "description": (
            "Send a virtual key to SAP. "
            "0=Enter, 3=Back, 8=Execute(F8), 11=Save(F11), 12=Cancel"
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "key_code": {
                    "type": "integer",
                    "description": "Virtual key number",
                },
            },
            "required": ["key_code"],
        },
    },
    {
        "name": "handle_popup",
        "description": (
            "Handle a SAP popup/dialog window. "
            "Use action='confirm' for OK/Yes, 'cancel' for Cancel/No, "
            "'input' to type text_input then confirm."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["confirm", "cancel", "yes", "no", "input"],
                },
                "text_input": {
                    "type": "string",
                    "description": "Text to enter in popup (only for action=input)",
                },
            },
            "required": ["action"],
        },
    },
    {
        "name": "handle_transport_request",
        "description": (
            "Handle the SAP transport/customising request popup that appears "
            "when saving configuration changes. Assigns save to a transport."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "transport_number": {
                    "type": "string",
                    "description": "Existing transport number to assign to (leave blank to create new)",
                },
            },
            "required": [],
        },
    },
    {
        "name": "get_sales_orders",
        "description": "Run VA05 and extract the list of sales orders.",
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
        "name": "get_material_movements",
        "description": "Run MB51 and extract material movement data.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "read_screen",
        "description": "Return the current SAP screen title and status bar text.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
]

# ── Tool Dispatcher ───────────────────────────────────────────────────────────
WRITE_TOOLS = {"set_field_value", "press_button", "send_vkey",
               "handle_popup", "handle_transport_request"}

def handle_tool(tool_name, tool_input):
    """Route Claude's tool call to the correct Python function."""

    # ── WRITE OPERATIONS need human approval ──────────────────────────────────
    if tool_name in WRITE_TOOLS:
        approved = ask_human_approval(
            action_description=f"SAP Write: {tool_name}",
            details=tool_input,
            risk="high" if tool_name == "handle_transport_request" else "medium",
        )
        if not approved:
            return {"status": "rejected", "message": "User rejected this action."}

    # ── Dispatch ──────────────────────────────────────────────────────────────
    if tool_name == "go_to_transaction":
        go_to_transaction(tool_input["tcode"])
        return {"navigated_to": tool_input["tcode"], "screen": get_screen_text()}

    elif tool_name == "discover_screen_elements":
        elements = discover_elements()
        return {"element_count": len(elements), "elements": elements[:80]}  # cap at 80

    elif tool_name == "set_field_value":
        return set_field_value(tool_input["element_id"], tool_input["value"])

    elif tool_name == "press_button":
        return press_button(tool_input["element_id"])

    elif tool_name == "send_vkey":
        return send_vkey(tool_input["key_code"])

    elif tool_name == "handle_popup":
        return handle_popup(
            action=tool_input.get("action", "confirm"),
            text_input=tool_input.get("text_input"),
        )

    elif tool_name == "handle_transport_request":
        return handle_transport_popup(
            transport_number=tool_input.get("transport_number")
        )

    elif tool_name == "get_sales_orders":
        return run_va05_and_extract(
            tool_input.get("date_from", "01.01.2025"),
            tool_input.get("date_to",   "31.12.2026"),
        )

    elif tool_name == "get_material_movements":
        return run_mb51_and_extract()

    elif tool_name == "read_screen":
        return {"screen": get_screen_text()}

    return {"error": f"Unknown tool: {tool_name}"}

# ── Agent Loop ────────────────────────────────────────────────────────────────
def run_agent(user_query, api_key):
    client = anthropic.Anthropic(api_key=api_key)

    system_prompt = f"""You are ARTILEGENZ — an expert SAP S/4 HANA consultant with direct GUI access.

CURRENT USER : {CURRENT_USER}
CURRENT TIME : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

CAPABILITIES:
- Navigate to any SAP transaction
- Discover screen elements (always do this before filling fields)
- Fill input fields by element ID
- Press buttons and send virtual keys
- Handle SAP popups and transport request dialogs
- Read screen content after each action

STRICT RULES:
1. ALWAYS call discover_screen_elements after navigating to a transaction
   before attempting to set any field values.
2. For SPRO / customising activities: after saving always call
   handle_transport_request to assign changes to a transport.
3. Confirm each major step by calling read_screen after it.
4. If a field ID from discover_screen_elements does not match what you
   expect, try the closest match or discover again.
5. For organisation structure (company codes, plants, sales orgs etc.):
   use the individual IMG transaction codes (OX02, OX10, OVX5, etc.)
   rather than navigating SPRO's tree — it is faster and more reliable.

COMMON VKEYS: 0=Enter  3=Back  8=Execute  11=Save  12=Cancel
"""

    messages = [{"role": "user", "content": user_query}]
    print("\nARTILEGENZ agent started...\n")

    while True:
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=8192,
            system=system_prompt,
            tools=TOOLS,
            messages=messages,
        )

        # Print any narrative Claude produces
        for block in response.content:
            if hasattr(block, "text") and block.text.strip():
                print(f"\n[Claude] {block.text}")

        if response.stop_reason == "end_turn":
            break

        if response.stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []

            for block in response.content:
                if block.type != "tool_use":
                    continue

                print(f"\n>> Tool: {block.name}  Input: {json.dumps(block.input)}")

                try:
                    result = handle_tool(block.name, block.input)
                except Exception as exc:
                    result = {"error": str(exc)}

                print(f"   Result: {json.dumps(result)[:300]}")

                tool_results.append({
                    "type":        "tool_result",
                    "tool_use_id": block.id,
                    "content":     json.dumps(result, ensure_ascii=False),
                })

            messages.append({"role": "user", "content": tool_results})
        else:
            break

    # Save report
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    topic = user_query[:40].replace(" ", "_").replace("/", "-")
    report_path = os.path.join(LOG_DIR, f"SAP_Session_{topic}_{ts}.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"Query  : {user_query}\n")
        f.write(f"Time   : {datetime.now()}\n")
        f.write(f"User   : {CURRENT_USER}\n\n")
        for m in messages:
            role = m["role"].upper()
            content = m["content"]
            if isinstance(content, list):
                for c in content:
                    if isinstance(c, dict) and c.get("type") == "text":
                        f.write(f"[{role}] {c['text']}\n")
                    elif hasattr(c, "text"):
                        f.write(f"[{role}] {c.text}\n")
            elif isinstance(content, str):
                f.write(f"[{role}] {content}\n")

    print(f"\n✓ Session saved → {report_path}")
    print(f"✓ Audit log    → {LOG_DIR}")

# ── Entry Point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
    if not API_KEY:
        API_KEY = input("Enter Anthropic API key: ").strip()

    print("\n" + "═" * 65)
    print("  ARTILEGENZ SAP Agent v3.0")
    print("  All write operations require your approval.")
    print("═" * 65)

    while True:
        query = input("\nWhat do you want to do in SAP? (or 'exit'): ").strip()
        if query.lower() in ("exit", "quit", "q"):
            break
        if not query:
            continue
        run_agent(query, API_KEY)
