from pathlib import Path
from src.reporter_md import render


def test_empty_findings(tmp_path):
    findings = {"summary": {"total_issues": 0, "by_severity": {}, "by_category": {}}, "findings": []}
    extracted = {"metadata": {"title": "테스트 보고서", "slide_count": 3}, "slides": []}
    out_path = tmp_path / "review.md"
    result = render(findings, extracted, out_path)

    assert result == out_path
    assert out_path.exists()
    text = out_path.read_text(encoding="utf-8")
    assert "이슈 없음" in text or "발견된 이슈가 없습니다" in text
    assert "테스트 보고서" in text
