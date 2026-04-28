"""Centralized settings for sap-error-intelligence."""
from __future__ import annotations

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    # Anthropic (only required for classifier)
    anthropic_api_key: str = ""
    claude_model: str = "claude-opus-4-7"
    claude_critique_model: str = "claude-opus-4-7"

    # Voyage (only required for embeddings)
    voyage_api_key: str = ""
    voyage_model: str = "voyage-3"
    voyage_dim: int = 1024

    # Pinecone (only required for vector search)
    pinecone_api_key: str = ""
    pinecone_index: str = "sap-errors"
    pinecone_namespace: str = "corpus_v1"
    pinecone_cloud: str = "aws"
    pinecone_region: str = "us-east-1"

    # Snowflake (only required when DB_ENGINE=snowflake)
    sf_account: str = ""
    sf_user: str = ""
    sf_private_key_path: str = ""
    sf_private_key_passphrase: str = ""
    sf_warehouse: str = "ARTILEGENZ_WH"
    sf_database: str = "SAP_ERRORS"
    sf_raw_schema: str = "RAW"
    sf_intel_schema: str = "INTEL"
    sf_role: str = "ARTILEGENZ_INTEL"

    # Postgres (only required when DB_ENGINE=postgres — default)
    pg_host: str = "localhost"
    pg_port: int = 5432
    pg_user: str = "artilegenz"
    pg_password: str = "artilegenz_local_dev"
    pg_database: str = "sap_errors"

    # DB engine switch — "postgres" (default) or "snowflake"
    db_engine: str = "postgres"

    # Classifier
    confidence_threshold: float = 0.70
    top_k_default: int = 8
    top_k_retry: int = 20
    max_retries: int = 1

    # API
    api_host: str = "0.0.0.0"
    api_port: int = 8000
    log_level: str = "INFO"


settings = Settings()  # type: ignore[call-arg]
