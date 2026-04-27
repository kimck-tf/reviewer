from pathlib import Path
import pytest


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_text_only(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_text_only.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p} — run tests/fixtures/_make_fixtures.py")
    return p


@pytest.fixture
def sample_with_table(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_table.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p}")
    return p


@pytest.fixture
def sample_with_image(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_image.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p}")
    return p


@pytest.fixture
def sample_with_embedded(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_embedded.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p} — manually create with embedded OLE object")
    return p


@pytest.fixture
def tmp_work_dir(tmp_path: Path) -> Path:
    return tmp_path / "review_ws"
