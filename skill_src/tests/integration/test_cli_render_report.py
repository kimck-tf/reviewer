"""render_report.py CLI 통합 테스트 (PowerPoint 불필요, 자동 실행)."""
import json
import sys
import subprocess
from pathlib import Path
import pytest


def test_render_report_creates_md_and_html(tmp_path):
    """findings.json + extracted.json → review.md + review.html."""
    work_dir = tmp_path / "ws"
    work_dir.mkdir()
    out_dir = tmp_path / "out"

    extracted = {"metadata": {"title": "T", "slide_count": 1}, "slides": [{"index": 1, "title": "표지", "thumbnail_path": None}]}
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    (work_dir / "extracted.json").write_text(json.dumps(extracted, ensure_ascii=False), encoding="utf-8")
    (work_dir / "findings.json").write_text(json.dumps(findings, ensure_ascii=False), encoding="utf-8")

    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "render_report.py"
    result = subprocess.run(
        [sys.executable, str(cli), str(work_dir), "--out", str(out_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode == 0, f"stderr: {result.stderr}"
    assert (out_dir / "review.md").exists()
    assert (out_dir / "review.html").exists()
    assert (out_dir / "assets" / "style.css").exists()


def test_render_report_missing_findings_errors(tmp_path):
    """findings.json 없으면 에러."""
    work_dir = tmp_path / "ws"
    work_dir.mkdir()
    (work_dir / "extracted.json").write_text("{}", encoding="utf-8")
    out_dir = tmp_path / "out"

    skill_dir = Path(__file__).parent.parent.parent
    cli = skill_dir / "render_report.py"
    result = subprocess.run(
        [sys.executable, str(cli), str(work_dir), "--out", str(out_dir)],
        capture_output=True, text=True, cwd=str(skill_dir),
    )
    assert result.returncode != 0
    assert "findings.json" in result.stderr
