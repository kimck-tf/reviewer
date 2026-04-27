from pathlib import Path
from src.reporter_html import render


def test_empty_findings_creates_html(tmp_path):
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    extracted = {"metadata": {"title": "T", "slide_count": 0}, "slides": []}
    out_dir = tmp_path / "out"
    result = render(findings, extracted, out_dir)

    assert result.exists()
    assert result.suffix == ".html"
    assert result.name == "review.html"
    text = result.read_text(encoding="utf-8")
    assert "<!DOCTYPE html>" in text
    assert "T" in text
    assert "이슈 없음" in text or "no issues" in text.lower()
    # CSS 복사 확인
    assert (out_dir / "assets" / "style.css").exists()
