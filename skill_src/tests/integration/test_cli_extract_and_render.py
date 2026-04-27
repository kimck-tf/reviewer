"""실제 PowerPoint COM이 필요한 통합 테스트. 사내 PC에서 수동 실행."""
import json
import sys
import subprocess
from pathlib import Path
import pytest

pytestmark = pytest.mark.real_com  # 기본 pytest 실행 시 제외됨


def test_cli_creates_extracted_json_and_images(sample_text_only: Path, tmp_path: Path):
    work_dir = tmp_path / "ws"
    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "extract_and_render.py"

    result = subprocess.run(
        [sys.executable, str(cli), str(sample_text_only), "--out", str(work_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode == 0, f"stderr: {result.stderr}"

    extracted_path = work_dir / "extracted.json"
    assert extracted_path.exists()
    data = json.loads(extracted_path.read_text(encoding="utf-8"))
    assert data["metadata"]["slide_count"] == 5
    for s in data["slides"]:
        assert s["image_path"] is not None
        assert Path(s["image_path"]).exists()
        assert s["thumbnail_path"] is not None
        assert Path(s["thumbnail_path"]).exists()


def test_cli_creates_slide_input_jsons(sample_text_only: Path, tmp_path: Path):
    """1단계 SA가 읽을 슬라이드별 입력 파일도 생성."""
    work_dir = tmp_path / "ws"
    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "extract_and_render.py"

    subprocess.run(
        [sys.executable, str(cli), str(sample_text_only), "--out", str(work_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )

    inputs_dir = work_dir / "slide_inputs"
    assert inputs_dir.exists()
    inputs = sorted(inputs_dir.glob("slide_*.json"))
    assert len(inputs) == 5
    sample = json.loads(inputs[0].read_text(encoding="utf-8"))
    assert "index" in sample
    assert "shapes" in sample
    assert "image_path" in sample
