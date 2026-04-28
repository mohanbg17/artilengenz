"""Shared pytest fixtures for IDoc extractor tests."""
from __future__ import annotations

from pathlib import Path

import pytest

from extractors.idoc import IDoc, MockIDocClient


FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    return FIXTURES_DIR


@pytest.fixture
def mock_client(fixtures_dir: Path) -> MockIDocClient:
    return MockIDocClient(fixtures_dir=fixtures_dir)


@pytest.fixture
def all_fixture_idocs(mock_client: MockIDocClient) -> list[IDoc]:
    """All IDocs from all fixture files, no status filter."""
    return mock_client.list_failed_idocs(statuses=None, max_results=999)
