"""
SAP MCP Server — exposes SAP GUI Scripting as Claude tools via stdio.

Usage (on the Windows Remote Desktop):
    python server.py

Configure in Claude Code (~/.claude.json or .mcp.json):
    {
      "mcpServers": {
        "sap": {
          "command": "python",
          "args": ["C:\\path\\to\\artilengenz\\server.py"]
        }
      }
    }
"""

import json
from mcp.server.fastmcp import FastMCP
from sap_client import SAPClient, SAPError

mcp = FastMCP("SAP MCP Server")
sap = SAPClient()


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _ok(data) -> str:
    if isinstance(data, (dict, list)):
        return json.dumps(data, indent=2, ensure_ascii=False)
    return str(data)


def _run(fn, *args, **kwargs) -> str:
    try:
        return _ok(fn(*args, **kwargs))
    except SAPError as e:
        return f"ERROR: {e}"
    except Exception as e:
        return f"ERROR (unexpected): {e}"


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def connect_sap(connection_index: int = 0, session_index: int = 0) -> str:
    """
    Attach to an already-open SAP GUI session on this machine.
    connection_index: which SAP connection (0 = first)
    session_index:    which SAP window/session (0 = first)
    Returns session info (system, user, current transaction).
    """
    return _run(sap.connect, connection_index, session_index)


@mcp.tool()
def get_session_info() -> str:
    """
    Return current SAP session details:
    system, client, user, language, active transaction, program, screen number, window title.
    """
    return _run(sap.get_session_info)


@mcp.tool()
def start_transaction(tcode: str) -> str:
    """
    Navigate to a SAP transaction code (e.g. SE16, ME21N, FB50, MM03, VA01).
    Equivalent to typing /n<tcode> in the command field.
    Returns updated session info.
    """
    return _run(sap.start_transaction, tcode)


@mcp.tool()
def get_screen_elements() -> str:
    """
    Get all interactive elements on the current SAP screen.
    Returns a list of {id, type, text, tooltip} objects.
    Use the 'id' values with set_field_value, press_button, get_field_value.
    """
    return _run(sap.get_screen_elements)


@mcp.tool()
def find_element_by_text(text: str) -> str:
    """
    Find screen elements whose label or tooltip contains the given text.
    Useful for locating the ID of a field when you know its label.
    """
    return _run(sap.find_element_by_text, text)


@mcp.tool()
def set_field_value(element_id: str, value: str) -> str:
    """
    Set a text field or combo box value.
    element_id: SAP element path, e.g. 'wnd[0]/usr/ctxtMATNR-LOW'
    value: the text to enter
    """
    return _run(sap.set_field_value, element_id, value)


@mcp.tool()
def get_field_value(element_id: str) -> str:
    """
    Read the current value of a field.
    element_id: SAP element path, e.g. 'wnd[0]/usr/ctxtMATNR-LOW'
    """
    return _run(sap.get_field_value, element_id)


@mcp.tool()
def press_button(element_id: str) -> str:
    """
    Click a button or toolbar button.
    element_id: SAP element path, e.g. 'wnd[0]/tbar[0]/btn[0]' (Execute)
    Common toolbar buttons:
      btn[0] = Enter/Execute   btn[3] = Save
      btn[15] = Back           btn[17] = Exit (F3)
    """
    return _run(sap.press_button, element_id)


@mcp.tool()
def press_key(key: str) -> str:
    """
    Send a keyboard shortcut to the SAP window.
    Supported values: Enter, Escape, Back, F1–F12, PgUp, PgDn
    Examples: 'F8' (Execute), 'F3' (Back), 'Enter' (confirm)
    """
    return _run(sap.press_key, key)


@mcp.tool()
def get_table_data(table_id: str) -> str:
    """
    Extract all rows from an ALV grid or classic table control.
    table_id: SAP element path of the table, e.g. 'wnd[0]/usr/cntlGRID/shellcont/shell'
    Returns a JSON array of row objects (column name → cell value).
    Tip: use get_screen_elements() to find the table's element ID first.
    """
    return _run(sap.get_table_data, table_id)


@mcp.tool()
def get_status_bar() -> str:
    """
    Read the SAP status bar message (message type S/W/E/A + text).
    Use after actions to check for success or error messages.
    """
    return _run(sap.get_status_bar)


@mcp.tool()
def take_screenshot() -> str:
    """
    Capture the current SAP window as a base64-encoded PNG.
    Use when you need to visually inspect the current screen state.
    """
    return _run(sap.take_screenshot)


@mcp.tool()
def create_sales_order(
    order_type: str,
    sales_org: str,
    dist_channel: str,
    division: str,
    sold_to: str,
    items: str,
    po_number: str = "",
    po_date: str = "",
    delivery_date: str = "",
) -> str:
    """
    Create a sales order in SAP via transaction VA01.

    Args:
        order_type:     SAP order type, e.g. 'OR' (standard order)
        sales_org:      Sales organization, e.g. '1000'
        dist_channel:   Distribution channel, e.g. '10'
        division:       Division, e.g. '00'
        sold_to:        Customer (sold-to party) number, e.g. '100023'
        items:          JSON array of line items, each with:
                          - material  (required): material number
                          - quantity  (required): order quantity
                          - plant     (optional): delivering plant
                        Example: '[{"material":"MAT001","quantity":"10","plant":"1000"}]'
        po_number:      Customer PO number (optional)
        po_date:        Customer PO date, format DD.MM.YYYY (optional)
        delivery_date:  Requested delivery date, format DD.MM.YYYY (optional)

    Returns JSON with: order_number, status ('success'/'error'), message.
    """
    try:
        parsed_items = json.loads(items)
    except json.JSONDecodeError as e:
        return f"ERROR: 'items' is not valid JSON — {e}"

    return _run(
        sap.create_sales_order,
        order_type,
        sales_org,
        dist_channel,
        division,
        sold_to,
        po_number,
        po_date,
        delivery_date,
        parsed_items,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
