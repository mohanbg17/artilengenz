"""
SAP GUI Scripting client for Windows.

Prerequisites on the Remote Desktop:
1. SAP GUI must be open and logged in.
2. SAP GUI Scripting must be enabled:
   - In SAP: Tools > Options > Accessibility & Scripting > Scripting
   - Check "Enable Scripting" and uncheck "Notify when a script attaches"
   - Or ask your SAP Basis admin to set profile parameter sapgui/scripting=1
"""

import base64
import json
from io import BytesIO

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


class SAPError(Exception):
    pass


class SAPClient:
    def __init__(self):
        self._session = None

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self, connection_index: int = 0, session_index: int = 0) -> dict:
        """
        Attach to an already-open SAP GUI session.
        connection_index: which connection (0 = first)
        session_index:    which session/window (0 = first)
        """
        if not WIN32_AVAILABLE:
            raise SAPError("pywin32 is not installed. Run: pip install pywin32")

        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
        except Exception as e:
            raise SAPError(
                f"Cannot find SAP GUI. Make sure SAP Logon is open and you are logged in. ({e})"
            )

        try:
            engine = sap_gui.GetScriptingEngine
            connection = engine.Children(connection_index)
            self._session = connection.Children(session_index)
        except Exception as e:
            raise SAPError(f"Cannot attach to session [{connection_index}][{session_index}]: {e}")

        return self._build_session_info()

    def _require_session(self):
        if self._session is None:
            raise SAPError("Not connected. Call connect_sap() first.")
        return self._session

    # ------------------------------------------------------------------
    # Session info
    # ------------------------------------------------------------------

    def _build_session_info(self) -> dict:
        s = self._require_session()
        info = s.Info
        try:
            title = s.FindById("wnd[0]").Text
        except Exception:
            title = ""
        return {
            "system": info.SystemName,
            "client": info.Client,
            "user": info.User,
            "language": info.Language,
            "transaction": info.Transaction,
            "program": info.Program,
            "screen_number": info.ScreenNumber,
            "window_title": title,
        }

    def get_session_info(self) -> dict:
        return self._build_session_info()

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------

    def start_transaction(self, tcode: str) -> dict:
        s = self._require_session()
        s.StartTransaction(tcode.strip().upper())
        return self._build_session_info()

    def press_key(self, key: str) -> str:
        """
        Send a virtual key to the main window.
        Supported: Enter, F1–F12, PgUp, PgDn, Escape, Back
        """
        KEY_MAP = {
            "Enter": 0, "Escape": 12, "Back": 3,
            "F1": 112, "F2": 113, "F3": 114, "F4": 115,
            "F5": 116, "F6": 117, "F7": 118, "F8": 119,
            "F9": 120, "F10": 121, "F11": 122, "F12": 123,
            "PgUp": 82, "PgDn": 83,
        }
        key_norm = key.strip()
        if key_norm not in KEY_MAP:
            raise SAPError(
                f"Unknown key '{key_norm}'. Supported: {', '.join(KEY_MAP.keys())}"
            )
        wnd = self._require_session().FindById("wnd[0]")
        wnd.SendVKey(KEY_MAP[key_norm])
        return f"Sent key: {key_norm}"

    # ------------------------------------------------------------------
    # Field / element interaction
    # ------------------------------------------------------------------

    def set_field_value(self, element_id: str, value: str) -> str:
        s = self._require_session()
        try:
            elem = s.FindById(element_id)
        except Exception:
            raise SAPError(f"Element not found: {element_id}")
        try:
            elem.Text = value
        except Exception as e:
            raise SAPError(f"Cannot set field '{element_id}': {e}")
        return f"Set '{element_id}' = '{value}'"

    def get_field_value(self, element_id: str) -> str:
        s = self._require_session()
        try:
            elem = s.FindById(element_id)
            return str(elem.Text)
        except Exception as e:
            raise SAPError(f"Cannot read field '{element_id}': {e}")

    def press_button(self, element_id: str) -> str:
        s = self._require_session()
        try:
            elem = s.FindById(element_id)
            elem.Press()
        except Exception as e:
            raise SAPError(f"Cannot press button '{element_id}': {e}")
        return f"Pressed '{element_id}'"

    def select_menu(self, menu_path: str) -> str:
        """Select a menu item by path, e.g. 'wnd[0]/mbar/menu[0]/menu[1]'"""
        s = self._require_session()
        try:
            s.FindById(menu_path).Select()
        except Exception as e:
            raise SAPError(f"Cannot select menu '{menu_path}': {e}")
        return f"Selected menu '{menu_path}'"

    # ------------------------------------------------------------------
    # Screen element discovery
    # ------------------------------------------------------------------

    def get_screen_elements(self, max_depth: int = 6) -> list:
        """
        Return a flat list of all interactive elements on the current screen.
        Each entry: {id, type, text, tooltip, enabled, changeable}
        """
        wnd = self._require_session().FindById("wnd[0]")
        elements = []
        self._walk_elements(wnd, elements, max_depth, 0)
        return elements

    # Types we skip (pure containers / layout)
    _SKIP_TYPES = {
        "GuiShell", "GuiUserArea", "GuiContainerShell",
        "GuiSplitterContainer", "GuiCustomControl",
    }

    def _walk_elements(self, component, result: list, max_depth: int, depth: int):
        if depth > max_depth:
            return
        try:
            elem_type = component.Type
            elem_id = component.Id

            if elem_type not in self._SKIP_TYPES:
                entry = {
                    "id": elem_id,
                    "type": elem_type,
                    "text": self._safe_attr(component, "Text"),
                    "tooltip": self._safe_attr(component, "Tooltip"),
                }
                # Extra info for editable fields
                if "TextField" in elem_type or "ComboBox" in elem_type:
                    entry["changeable"] = self._safe_attr(component, "Changeable")
                result.append(entry)

            if hasattr(component, "Children"):
                for i in range(component.Children.Count):
                    self._walk_elements(component.Children(i), result, max_depth, depth + 1)
        except Exception:
            pass

    @staticmethod
    def _safe_attr(obj, attr: str, default=""):
        try:
            return str(getattr(obj, attr))
        except Exception:
            return default

    def find_element_by_text(self, text: str) -> list:
        """Find elements whose text or tooltip contains the given string (case-insensitive)."""
        needle = text.lower()
        return [
            e for e in self.get_screen_elements()
            if needle in e.get("text", "").lower()
            or needle in e.get("tooltip", "").lower()
        ]

    # ------------------------------------------------------------------
    # Table / grid data
    # ------------------------------------------------------------------

    def get_table_data(self, table_id: str) -> list:
        """
        Extract rows from an ALV GridView or classic TableControl.
        Returns a list of dicts (column name → cell value).
        """
        s = self._require_session()
        try:
            table = s.FindById(table_id)
        except Exception:
            raise SAPError(f"Table element not found: {table_id}")

        rows = []
        t_type = table.Type

        if t_type == "GuiGridView":
            col_count = table.ColumnCount
            col_names = []
            for i in range(col_count):
                try:
                    col_names.append(table.ColumnOrder(i))
                except Exception:
                    col_names.append(str(i))

            row_count = table.RowCount
            for r in range(row_count):
                row_data = {}
                for col in col_names:
                    try:
                        row_data[col] = table.GetCellValue(r, col)
                    except Exception:
                        row_data[col] = ""
                rows.append(row_data)

        elif t_type == "GuiTableControl":
            col_names = [c.Name for c in table.Columns]
            for r in range(table.Rows.Count):
                row_data = {}
                for col in col_names:
                    try:
                        row_data[col] = table.GetCell(r, col).Text
                    except Exception:
                        row_data[col] = ""
                rows.append(row_data)

        else:
            raise SAPError(f"Unsupported table type '{t_type}' for element '{table_id}'")

        return rows

    # ------------------------------------------------------------------
    # Screenshot
    # ------------------------------------------------------------------

    def take_screenshot(self) -> str:
        """
        Capture the SAP main window and return a base64-encoded PNG string.
        Requires Pillow: pip install Pillow
        """
        try:
            from PIL import ImageGrab
        except ImportError:
            raise SAPError("Pillow is not installed. Run: pip install Pillow")

        wnd = self._require_session().FindById("wnd[0]")
        left = wnd.ScreenLeft
        top = wnd.ScreenTop
        right = left + wnd.Width
        bottom = top + wnd.Height

        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        buf = BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def get_status_bar(self) -> dict:
        """Read the SAP status bar (message type + text)."""
        s = self._require_session()
        try:
            bar = s.FindById("wnd[0]/sbar")
            return {
                "message_type": self._safe_attr(bar, "MessageType"),
                "text": self._safe_attr(bar, "Text"),
            }
        except Exception as e:
            raise SAPError(f"Cannot read status bar: {e}")

    # ------------------------------------------------------------------
    # VA01 — Create Sales Order
    # ------------------------------------------------------------------

    # Element IDs for standard SAP ECC / S/4HANA VA01
    _VA01 = {
        # Initial screen
        "order_type":    "wnd[0]/usr/ctxtVBAK-AUART",
        "sales_org":     "wnd[0]/usr/ctxtVBAK-VKORG",
        "dist_channel":  "wnd[0]/usr/ctxtVBAK-VTWEG",
        "division":      "wnd[0]/usr/ctxtVBAK-SPART",
        # Overview / header (after Enter on initial screen)
        "sold_to":       "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subSUBSCREEN_HEADER2:SAPMV45A:4007/ctxtKUNAG-KUNNR",
        "po_number":     "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subSUBSCREEN_HEADER2:SAPMV45A:4007/ctxtVBAK-BSTNK",
        "po_date":       "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subSUBSCREEN_HEADER2:SAPMV45A:4007/ctxtVBAK-BSTDK",
        "delivery_date": "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subSUBSCREEN_HEADER2:SAPMV45A:4007/ctxtRV45A-KETDAT",
        # Line item table columns (row appended at runtime)
        "item_material": "wnd[0]/usr/subSUBSCREEN_BODY:SAPMV45A:4900/subSUBSCREEN_TC:SAPMV45A:4050/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[0,{row}]",
        "item_quantity":  "wnd[0]/usr/subSUBSCREEN_BODY:SAPMV45A:4900/subSUBSCREEN_TC:SAPMV45A:4050/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[1,{row}]",
        "item_plant":    "wnd[0]/usr/subSUBSCREEN_BODY:SAPMV45A:4900/subSUBSCREEN_TC:SAPMV45A:4050/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[2,{row}]",
        # Buttons
        "save":          "wnd[0]/tbar[0]/btn[11]",
    }

    def _dismiss_popup(self):
        """Press Enter to dismiss any confirmation popup that may appear."""
        try:
            popup = self._session.FindById("wnd[1]")
            popup.SendVKey(0)  # Enter
        except Exception:
            pass  # No popup — fine

    def create_sales_order(
        self,
        order_type: str,
        sales_org: str,
        dist_channel: str,
        division: str,
        sold_to: str,
        po_number: str,
        po_date: str,
        delivery_date: str,
        items: list,
    ) -> dict:
        """
        Create a sales order via VA01.

        items: list of dicts with keys 'material', 'quantity', and optionally 'plant'
        dates: DD.MM.YYYY format

        Returns: {"order_number": "...", "status": "...", "message": "..."}
        """
        s = self._require_session()
        ids = self._VA01

        # 1. Navigate to VA01
        s.StartTransaction("VA01")

        # 2. Fill initial screen
        s.FindById(ids["order_type"]).Text   = order_type.strip().upper()
        s.FindById(ids["sales_org"]).Text    = sales_org.strip()
        s.FindById(ids["dist_channel"]).Text = dist_channel.strip()
        s.FindById(ids["division"]).Text     = division.strip()

        # 3. Confirm initial screen → go to overview
        s.FindById("wnd[0]").SendVKey(0)  # Enter
        self._dismiss_popup()

        # 4. Fill header fields
        s.FindById(ids["sold_to"]).Text       = sold_to.strip()
        s.FindById("wnd[0]").SendVKey(0)       # Enter to resolve customer name
        self._dismiss_popup()

        if po_number:
            s.FindById(ids["po_number"]).Text  = po_number.strip()
        if po_date:
            s.FindById(ids["po_date"]).Text    = po_date.strip()
        if delivery_date:
            s.FindById(ids["delivery_date"]).Text = delivery_date.strip()

        # 5. Fill line items
        for row, item in enumerate(items):
            material = item.get("material", "").strip()
            quantity = str(item.get("quantity", "")).strip()
            plant    = item.get("plant", "").strip()

            if not material:
                continue

            s.FindById(ids["item_material"].format(row=row)).Text = material
            if quantity:
                s.FindById(ids["item_quantity"].format(row=row)).Text = quantity
            if plant:
                s.FindById(ids["item_plant"].format(row=row)).Text = plant

        # 6. Save
        s.FindById(ids["save"]).Press()
        self._dismiss_popup()

        # 7. Read result from status bar
        bar = s.FindById("wnd[0]/sbar")
        msg_type = self._safe_attr(bar, "MessageType")
        msg_text = self._safe_attr(bar, "Text")

        # Extract order number from message like "Standard Order 1234567 has been saved"
        order_number = ""
        import re
        match = re.search(r"\b(\d{7,10})\b", msg_text)
        if match:
            order_number = match.group(1)

        return {
            "order_number": order_number,
            "status": "success" if msg_type == "S" else "error",
            "message": msg_text,
        }
