"""Centralised configuration using Pydantic Settings."""

from __future__ import annotations

from enum import Enum
from functools import lru_cache
from typing import Literal

from pydantic import Field, SecretStr, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class SAPSystemType(str, Enum):
    S4HANA = "S4HANA"
    ECC = "ECC"


class AuthType(str, Enum):
    BASIC = "basic"
    OAUTH2 = "oauth2"
    SAML = "saml"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # ── LLM ─────────────────────────────────────────────────────────────────
    anthropic_api_key: SecretStr = Field(default="", alias="ANTHROPIC_API_KEY")
    llm_model: str = Field(default="claude-sonnet-4-6", alias="LLM_MODEL")
    llm_max_tokens: int = Field(default=4096, alias="LLM_MAX_TOKENS")
    llm_temperature: float = Field(default=0.0, alias="LLM_TEMPERATURE")

    # ── SAP HTTP / OData ─────────────────────────────────────────────────────
    sap_host: str = Field(default="localhost", alias="SAP_HOST")
    sap_http_port: int = Field(default=443, alias="SAP_HTTP_PORT")
    sap_https: bool = Field(default=True, alias="SAP_HTTPS")
    sap_client: str = Field(default="100", alias="SAP_CLIENT")
    sap_odata_base_path: str = Field(
        default="/sap/opu/odata/sap", alias="SAP_ODATA_BASE_PATH"
    )
    sap_auth_type: AuthType = Field(default=AuthType.BASIC, alias="SAP_AUTH_TYPE")
    sap_username: str = Field(default="", alias="SAP_USERNAME")
    sap_password: SecretStr = Field(default="", alias="SAP_PASSWORD")

    # OAuth2
    sap_token_url: str = Field(default="", alias="SAP_TOKEN_URL")
    sap_client_id: str = Field(default="", alias="SAP_CLIENT_ID")
    sap_client_secret: SecretStr = Field(default="", alias="SAP_CLIENT_SECRET")

    # ── SAP RFC ──────────────────────────────────────────────────────────────
    sap_rfc_host: str = Field(default="", alias="SAP_RFC_HOST")
    sap_rfc_sysnr: str = Field(default="00", alias="SAP_RFC_SYSNR")
    sap_rfc_logon_group: str = Field(default="", alias="SAP_RFC_LOGON_GROUP")
    sap_rfc_enabled: bool = Field(default=False, alias="SAP_RFC_ENABLED")

    # ── SAP Network / SSL ────────────────────────────────────────────────────
    # Proxy URL e.g. http://proxy.corp.com:8080 — leave blank to auto-detect
    # from system (HTTPS_PROXY env var or Windows registry proxy settings)
    sap_proxy: str = Field(default="", alias="SAP_PROXY")
    # Set false for self-signed SAP dev certificates (common on sandbox systems)
    sap_verify_ssl: bool = Field(default=False, alias="SAP_VERIFY_SSL")

    # ── SAP OData Service Overrides ───────────────────────────────────────────
    # Override the entity set name for FAC_FINANCIAL_DOCUMENT_SRV_01.
    # Try "HeaderSet" first; if probe returns 404 change to "HeaderCollection".
    sap_fac_entity_set: str = Field(default="HeaderSet", alias="SAP_FAC_ENTITY_SET")

    # ── SAP System ───────────────────────────────────────────────────────────
    sap_system_type: SAPSystemType = Field(
        default=SAPSystemType.S4HANA, alias="SAP_SYSTEM_TYPE"
    )
    sap_language: str = Field(default="EN", alias="SAP_LANGUAGE")

    # ── Application ──────────────────────────────────────────────────────────
    app_host: str = Field(default="0.0.0.0", alias="APP_HOST")
    app_port: int = Field(default=8080, alias="APP_PORT")
    app_env: Literal["development", "production"] = Field(
        default="development", alias="APP_ENV"
    )
    log_level: str = Field(default="INFO", alias="LOG_LEVEL")
    mock_sap: bool = Field(default=True, alias="MOCK_SAP")

    @field_validator("llm_temperature")
    @classmethod
    def validate_temperature(cls, v: float) -> float:
        if not 0.0 <= v <= 1.0:
            raise ValueError("LLM_TEMPERATURE must be between 0.0 and 1.0")
        return v

    @property
    def sap_base_url(self) -> str:
        scheme = "https" if self.sap_https else "http"
        return f"{scheme}://{self.sap_host}:{self.sap_http_port}"

    @property
    def rfc_params(self) -> dict:
        """Build pyrfc connection parameters dict."""
        return {
            "ashost": self.sap_rfc_host,
            "sysnr": self.sap_rfc_sysnr,
            "client": self.sap_client,
            "user": self.sap_username,
            "passwd": self.sap_password.get_secret_value(),
            "lang": self.sap_language,
        }


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
