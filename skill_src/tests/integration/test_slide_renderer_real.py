"""실제 PowerPoint COM을 호출하는 통합 테스트.

수동 실행:
  .venv/Scripts/pytest tests/integration/test_slide_renderer_real.py -v -m real_com

CI 환경에서 자동 실행하지 않으려면 pyproject.toml의 addopts 설정으로 기본 제외됨.
"""
import sys
import pytest
from pathlib import Path


pytestmark = [
    pytest.mark.real_com,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows + MS PowerPoint 필요"),
]


def test_real_mode1_text_only(sample_text_only: Path, tmp_path: Path):
    from src.slide_renderer import convert_pptx_to_images_without_embedding

    out_dir = tmp_path / "out"
    result = convert_pptx_to_images_without_embedding(sample_text_only, out_dir)

    assert set(result.keys()) == {1, 2, 3, 4, 5}  # fixture는 5장
    for idx, p in result.items():
        assert p.exists()
        assert p.stat().st_size > 0


def test_real_render_with_thumbnails(sample_text_only: Path, tmp_path: Path):
    from src.extractor import extract
    from src.slide_renderer import render

    extracted = extract(sample_text_only)
    out_dir = tmp_path / "out"
    result = render(sample_text_only, out_dir, extracted)

    assert len(result) == 5
    thumb_dir = out_dir / "thumbnails"
    assert len(list(thumb_dir.glob("slide_*.jpg"))) == 5


def test_real_render_with_embedded(sample_with_embedded: Path, tmp_path: Path):
    """sample_with_embedded.pptx 부재 시 conftest의 fixture가 자동 skip."""
    from src.extractor import extract
    from src.slide_renderer import render

    extracted = extract(sample_with_embedded)
    out_dir = tmp_path / "out"
    result = render(sample_with_embedded, out_dir, extracted)

    assert len(result) > 0
    for paths in result.values():
        assert len(paths) >= 1
