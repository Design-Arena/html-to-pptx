from pathlib import Path

import pytest
import pytest_asyncio

from html_to_pptx import extract_measurements

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def minimal_html() -> Path:
    return FIXTURES_DIR / "minimal.html"


@pytest.fixture
def tmp_pptx(tmp_path: Path) -> Path:
    return tmp_path / "output.pptx"


@pytest_asyncio.fixture(scope="session")
async def minimal_measurements() -> list[dict]:
    """Extract measurements once per test session — avoids repeated browser launches."""
    return await extract_measurements(str(FIXTURES_DIR / "minimal.html"))


@pytest_asyncio.fixture(scope="session")
async def fidelity_measurements() -> list[dict]:
    return await extract_measurements(str(FIXTURES_DIR / "fidelity.html"))
