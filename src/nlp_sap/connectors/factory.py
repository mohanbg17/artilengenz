"""Connector factory — picks the right connector based on config."""

from __future__ import annotations

from nlp_sap.config import Settings, get_settings
from nlp_sap.connectors.base import BaseSAPConnector


def build_connector(settings: Settings | None = None) -> BaseSAPConnector:
    """Return the appropriate connector for the current environment.

    Selection priority:
    1. MOCK_SAP=true  → MockSAPConnector (no SAP needed)
    2. SAP_RFC_ENABLED=true → RFCConnector (ECC legacy / BAPI preferred)
    3. Default → ODataConnector (S4/HANA preferred)
    """
    if settings is None:
        settings = get_settings()

    if settings.mock_sap:
        from nlp_sap.connectors.mock import MockSAPConnector

        return MockSAPConnector()

    if settings.sap_rfc_enabled:
        from nlp_sap.connectors.rfc import RFCConnector

        return RFCConnector(settings)

    from nlp_sap.connectors.odata import ODataConnector

    return ODataConnector(settings)
