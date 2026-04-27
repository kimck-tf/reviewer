from pathlib import Path
import sys
from unittest.mock import MagicMock
from PIL import Image as _PILImage
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
        pytest.skip(f"Fixture missing: {p} — run tests/fixtures/_make_fixtures.py")
    return p


@pytest.fixture
def sample_with_image(fixtures_dir: Path) -> Path:
    p = fixtures_dir / "sample_with_image.pptx"
    if not p.exists():
        pytest.skip(f"Fixture missing: {p} — run tests/fixtures/_make_fixtures.py")
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


@pytest.fixture
def mock_powerpoint_com(monkeypatch):
    """`win32com.client.Dispatch`와 `pythoncom`을 mock하여 실제 PowerPoint 호출 없이 테스트.

    slide.Export(path, fmt, w, h) 호출 시 실제로 PIL 더미 JPG를 path에 생성.
    """
    fake_win32 = MagicMock()
    fake_pythoncom = MagicMock()

    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32.client)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    def make_app(slide_count: int = 3, embedded_progids_per_slide: dict | None = None):
        app = MagicMock(name="PowerPointApplication")
        app.Visible = 0
        pres = MagicMock(name="Presentation")
        slides = []
        for i in range(1, slide_count + 1):
            sl = MagicMock(name=f"Slide{i}")
            sl.SlideIndex = i

            embedded = (embedded_progids_per_slide or {}).get(i, [])
            shapes = []
            for j, progid in enumerate(embedded, start=1):
                shape = MagicMock(name=f"Slide{i}Shape{j}")
                shape.Type = 7  # msoEmbeddedOLEObject
                shape.OLEFormat.ProgID = progid
                shapes.append(shape)
            sl.Shapes = shapes

            def _make_export(idx):
                def _export(path, fmt, w, h):
                    img = _PILImage.new("RGB", (w, h), color=(220, 230, 240))
                    img.save(path)
                return _export
            sl.Export.side_effect = _make_export(i)
            slides.append(sl)
        pres.Slides = slides
        app.Presentations.Open.return_value = pres
        fake_win32.client.Dispatch.return_value = app
        return app, pres, slides

    return make_app
