"""Query Builder — translates a ParsedIntent into a SAPQueryPlan.

Applies SAP system-type logic:
  - S4/HANA: prefer OData APIs → ACDOCA (universal journal)
  - ECC: prefer RFC/BAPI → BKPF/BSEG, FAGLFLEXT
  - Both: use BAPI when available and specific
"""

from __future__ import annotations

import logging
from datetime import date, datetime, timedelta

from nlp_sap.config import SAPSystemType, Settings, get_settings
from nlp_sap.connectors.base import SAPQueryRequest
from nlp_sap.nlp.models import ParsedIntent, SAPQueryPlan
from nlp_sap.schema.registry import SchemaRegistry, get_registry

logger = logging.getLogger(__name__)


class QueryBuilder:
    """Converts a ParsedIntent into a SAPQueryRequest ready for a connector."""

    def __init__(
        self,
        settings: Settings | None = None,
        registry: SchemaRegistry | None = None,
    ) -> None:
        self._settings = settings or get_settings()
        self._registry = registry or get_registry()

    # ── Public API ────────────────────────────────────────────────────────

    def build_plan(self, intent: ParsedIntent) -> SAPQueryPlan:
        """Generate a SAPQueryPlan for the intent."""
        handler = getattr(self, f"_plan_{intent.intent_name}", self._plan_generic)
        return handler(intent)

    def plan_to_request(self, plan: SAPQueryPlan) -> SAPQueryRequest:
        """Convert a SAPQueryPlan to a SAPQueryRequest for a connector."""
        req = SAPQueryRequest(
            table_or_function=plan.bapi_name or plan.primary_table,
            query_type=plan.query_type,
            filters=plan.filters,
            fields=plan.fields,
            max_rows=plan.max_rows,
            join_table=plan.join_table,
            join_on=plan.join_on,
            odata_service=plan.odata_service,
            odata_entity_set=plan.odata_entity_set,
            bapi_import_params=plan.bapi_import_params,
            bapi_table_params=plan.bapi_output_tables,
        )
        # Attach intent metadata (used by mock connector for fixture routing)
        req.__dict__["metadata"] = {"intent": plan.intent.intent_name}
        return req

    # ── Intent-specific plan builders ─────────────────────────────────────

    def _plan_gl_balance_inquiry(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        if intent.gl_account:
            filters["RACCT"] = intent.gl_account
        if intent.cost_center:
            filters["RCNTR"] = intent.cost_center
        if intent.profit_center:
            filters["PRCTR"] = intent.profit_center

        if self._is_s4():
            return SAPQueryPlan(
                intent=intent,
                primary_table="ACDOCA",
                query_type="odata",
                odata_service="API_GLACCOUNTLINEITEM_SRV",
                odata_entity_set="A_GLAccountLineItem",
                filters=self._map_fi_odata_filters(filters),
                fields=["CompanyCode", "FiscalYear", "GLAccount", "PostingDate",
                        "AmountInCompanyCodeCurrency", "CompanyCodeCurrency",
                        "CostCenter", "ProfitCenter"],
                aggregations=["sum_amount"],
                max_rows=intent.max_rows,
            )
        else:
            return SAPQueryPlan(
                intent=intent,
                primary_table="FAGLFLEXT",
                query_type="table",
                filters=filters,
                fields=["RBUKRS", "RYEAR", "RACCT", "PRCTR",
                        "HSL01", "HSL02", "HSL03", "HSL04",
                        "HSL05", "HSL06", "HSL07", "HSL08",
                        "HSL09", "HSL10", "HSL11", "HSL12"],
                aggregations=["sum_amount"],
                max_rows=intent.max_rows,
            )

    def _plan_ar_open_items(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        if intent.customer:
            filters["KUNNR"] = intent.customer
        if intent.overdue_only:
            filters["FAEDT"] = {"lte": date.today().isoformat()}

        return SAPQueryPlan(
            intent=intent,
            primary_table="BSID",
            query_type="bapi",
            bapi_name="BAPI_AR_ACC_GETOPENITEMS",
            bapi_import_params={
                "COMPANYCODE": intent.company_code or "",
                "CUSTOMER": intent.customer or "",
                "KEYDATE": date.today().isoformat(),
            },
            bapi_output_tables=["LINEITEMS"],
            filters=filters,
            aggregations=["sum_open_amount", "aging_buckets"],
            max_rows=intent.max_rows,
        )

    def _plan_ap_open_items(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        if intent.vendor:
            filters["LIFNR"] = intent.vendor
        if intent.overdue_only:
            filters["FAEDT"] = {"lte": date.today().isoformat()}

        return SAPQueryPlan(
            intent=intent,
            primary_table="BSIK",
            query_type="bapi",
            bapi_name="BAPI_AP_ACC_GETOPENITEMS",
            bapi_import_params={
                "COMPANYCODE": intent.company_code or "",
                "VENDOR": intent.vendor or "",
                "KEYDATE": date.today().isoformat(),
            },
            bapi_output_tables=["LINEITEMS"],
            filters=filters,
            aggregations=["sum_open_amount"],
            max_rows=intent.max_rows,
        )

    def _plan_document_search(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        if intent.document_number:
            filters["BELNR"] = intent.document_number
        if intent.document_type:
            filters["BLART"] = intent.document_type

        return SAPQueryPlan(
            intent=intent,
            primary_table="BKPF",
            join_table="BSEG",
            join_on={"BELNR": "BELNR", "GJAHR": "GJAHR", "BUKRS": "BUKRS"},
            query_type="table",
            filters=filters,
            fields=["BUKRS", "BELNR", "GJAHR", "BLART", "BLDAT", "BUDAT",
                    "WAERS", "BKTXT", "USNAM", "XBLNR"],
            max_rows=intent.max_rows,
        )

    def _plan_cost_center_report(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.company_code:
            filters["KOKRS"] = intent.company_code  # Often same as company code
        if intent.fiscal_year:
            filters["GJAHR"] = intent.fiscal_year
        if intent.cost_center:
            filters["KOSTL"] = intent.cost_center
        if intent.fiscal_period:
            filters["WRTTP"] = "04"   # Actual value type

        return SAPQueryPlan(
            intent=intent,
            primary_table="COSP",
            query_type="table",
            filters=filters,
            fields=["OBJNR", "GJAHR", "KSTAR", "WRTTP", "VERSN",
                    "WKG001", "WKG002", "WKG003", "WKG004",
                    "WKG005", "WKG006", "WKG007", "WKG008",
                    "WKG009", "WKG010", "WKG011", "WKG012"],
            aggregations=["sum_actual", "sum_plan", "variance"],
            max_rows=intent.max_rows,
        )

    def _plan_profit_center_report(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        if intent.profit_center:
            filters["PRCTR"] = intent.profit_center
        return self._plan_gl_balance_inquiry(intent)  # Same data source

    def _plan_purchase_order_status(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.company_code:
            filters["BUKRS"] = intent.company_code
        if intent.vendor:
            filters["LIFNR"] = intent.vendor
        if intent.plant:
            filters["WERKS"] = intent.plant
        if intent.po_number:
            filters["EBELN"] = intent.po_number
        if intent.date_from:
            filters["BEDAT"] = {"gte": intent.date_from, "lte": intent.date_to or date.today().isoformat()}

        if self._is_s4():
            return SAPQueryPlan(
                intent=intent,
                primary_table="EKKO",
                query_type="odata",
                odata_service="API_PURCHASEORDER_PROCESS_SRV",
                odata_entity_set="A_PurchaseOrder",
                odata_expand=["to_PurchaseOrderItem"],
                filters=self._map_po_odata_filters(filters),
                fields=["PurchaseOrder", "Supplier", "CompanyCode",
                        "PurchaseOrderDate", "DocumentCurrency", "NetPaymentAmount"],
                max_rows=intent.max_rows,
            )
        else:
            return SAPQueryPlan(
                intent=intent,
                primary_table="EKKO",
                join_table="EKPO",
                join_on={"EBELN": "EBELN"},
                query_type="table",
                filters=filters,
                fields=["EBELN", "LIFNR", "BUKRS", "EKORG", "BSART",
                        "BEDAT", "WAERS", "NETWR"],
                max_rows=intent.max_rows,
            )

    def _plan_goods_movement(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.plant:
            filters["WERKS"] = intent.plant
        if intent.material:
            filters["MATNR"] = intent.material
        if intent.po_number:
            filters["EBELN"] = intent.po_number
        if intent.document_type:
            filters["BWART"] = intent.document_type
        if intent.date_from:
            filters["BUDAT"] = {"gte": intent.date_from, "lte": intent.date_to or date.today().isoformat()}

        return SAPQueryPlan(
            intent=intent,
            primary_table="MSEG",
            query_type="table",
            filters=filters,
            fields=["MBLNR", "MJAHR", "ZEILE", "MATNR", "WERKS",
                    "LGORT", "BWART", "MENGE", "MEINS", "EBELN"],
            max_rows=intent.max_rows,
        )

    def _plan_inventory_stock(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.plant:
            filters["WERKS"] = intent.plant
        if intent.material:
            filters["MATNR"] = intent.material
        if intent.storage_location:
            filters["LGORT"] = intent.storage_location

        return SAPQueryPlan(
            intent=intent,
            primary_table="MCHB",
            query_type="table",
            filters=filters,
            fields=["MATNR", "WERKS", "LGORT", "CHARG", "CLABS", "CEINM"],
            max_rows=intent.max_rows,
        )

    def _plan_sales_order_status(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.customer:
            filters["KUNNR"] = intent.customer
        if intent.sales_org:
            filters["VKORG"] = intent.sales_org
        if intent.date_from:
            filters["ERDAT"] = {"gte": intent.date_from, "lte": intent.date_to or date.today().isoformat()}
        if intent.document_number:
            filters["VBELN"] = intent.document_number

        if self._is_s4():
            return SAPQueryPlan(
                intent=intent,
                primary_table="VBAK",
                query_type="odata",
                odata_service="API_SALES_ORDER_SRV",
                odata_entity_set="A_SalesOrder",
                odata_expand=["to_Item"],
                filters=self._map_so_odata_filters(filters),
                fields=["SalesOrder", "SoldToParty", "SalesOrganization",
                        "SalesOrderDate", "TransactionCurrency", "TotalNetAmount"],
                max_rows=intent.max_rows,
            )
        else:
            return SAPQueryPlan(
                intent=intent,
                primary_table="VBAK",
                join_table="VBAP",
                join_on={"VBELN": "VBELN"},
                query_type="table",
                filters=filters,
                fields=["VBELN", "KUNNR", "AUART", "VKORG", "ERDAT", "NETWR", "WAERK"],
                max_rows=intent.max_rows,
            )

    def _plan_revenue_report(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters: dict = {}
        if intent.customer:
            filters["KUNAG"] = intent.customer
        if intent.date_from:
            filters["FKDAT"] = {"gte": intent.date_from, "lte": intent.date_to or date.today().isoformat()}

        return SAPQueryPlan(
            intent=intent,
            primary_table="VBRK",
            query_type="table",
            filters=filters,
            fields=["VBELN", "FKART", "KUNAG", "FKDAT", "NETWR", "WAERK"],
            aggregations=["sum_revenue", "monthly_breakdown"],
            max_rows=intent.max_rows,
        )

    def _plan_cash_flow(self, intent: ParsedIntent) -> SAPQueryPlan:
        filters = self._base_fi_filters(intent)
        filters["KOART"] = "B"  # Bank GL accounts
        return SAPQueryPlan(
            intent=intent,
            primary_table="BSEG",
            query_type="table",
            filters=filters,
            fields=["BUKRS", "BELNR", "GJAHR", "BUZEI", "HKONT",
                    "DMBTR", "SHKZG", "BUDAT"],
            aggregations=["net_cashflow", "daily_balance"],
            max_rows=intent.max_rows,
        )

    def _plan_generic(self, intent: ParsedIntent) -> SAPQueryPlan:
        """Fallback generic plan for unknown intents."""
        logger.warning("No specific builder for intent '%s'; using generic plan", intent.intent_name)
        intent_def = self._registry.get_intent(intent.intent_name) or {}
        table = intent_def.get("primary_table", "BKPF")
        filters = self._base_fi_filters(intent)
        return SAPQueryPlan(
            intent=intent,
            primary_table=table,
            query_type="table",
            filters=filters,
            max_rows=intent.max_rows,
        )

    # ── Filter helpers ────────────────────────────────────────────────────

    def _base_fi_filters(self, intent: ParsedIntent) -> dict:
        """Common FI filter fields populated from intent."""
        filters: dict = {}
        if intent.company_code:
            filters["BUKRS"] = intent.company_code
        if intent.fiscal_year:
            filters["GJAHR"] = intent.fiscal_year
        if intent.fiscal_periods:
            filters["MONAT"] = intent.fiscal_periods
        elif intent.fiscal_period:
            filters["MONAT"] = intent.fiscal_period
        if intent.date_from:
            filters["BUDAT"] = {"gte": intent.date_from, "lte": intent.date_to or date.today().isoformat()}
        return filters

    def _map_fi_odata_filters(self, raw: dict) -> dict:
        """Map internal field names to OData API field names for FI."""
        mapping = {
            "BUKRS": "CompanyCode",
            "GJAHR": "FiscalYear",
            "MONAT": "FiscalPeriod",
            "RACCT": "GLAccount",
            "RCNTR": "CostCenter",
            "PRCTR": "ProfitCenter",
            "BUDAT": "PostingDate",
        }
        return {mapping.get(k, k): v for k, v in raw.items()}

    def _map_po_odata_filters(self, raw: dict) -> dict:
        mapping = {
            "BUKRS": "CompanyCode",
            "LIFNR": "Supplier",
            "EBELN": "PurchaseOrder",
            "WERKS": "Plant",
            "BEDAT": "PurchaseOrderDate",
        }
        return {mapping.get(k, k): v for k, v in raw.items()}

    def _map_so_odata_filters(self, raw: dict) -> dict:
        mapping = {
            "KUNNR": "SoldToParty",
            "VKORG": "SalesOrganization",
            "VBELN": "SalesOrder",
            "ERDAT": "SalesOrderDate",
        }
        return {mapping.get(k, k): v for k, v in raw.items()}

    def _is_s4(self) -> bool:
        return self._settings.sap_system_type == SAPSystemType.S4HANA
