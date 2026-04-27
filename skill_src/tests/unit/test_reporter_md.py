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


def test_single_finding_includes_all_fields(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"critical": 1}, "by_category": {"data": 1}},
        "findings": [
            {
                "id": "F001",
                "category": "data",
                "severity": "critical",
                "slide_index": 5,
                "shape_id": "s5_sh3",
                "position_hint": "슬라이드 5 우측 상단 텍스트 박스 (좌측 70%, 상단 15%)",
                "quoted_text": "최대 응력 250 MPa",
                "issue": "표 5의 셀에는 240 MPa로 기재되어 있어 본문 인용과 불일치",
                "suggestion": "본문을 240 MPa로 수정",
                "evidence": "slide_5 표(s5_sh4) 행 3·열 2 = '240'",
            }
        ],
    }
    extracted = {"metadata": {"title": "테스트", "slide_count": 5}, "slides": []}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)

    text = out_path.read_text(encoding="utf-8")
    assert "F001" in text
    assert "data" in text or "데이터" in text
    assert "critical" in text or "Critical" in text or "심각" in text
    assert "슬라이드 5" in text
    assert "최대 응력 250 MPa" in text
    assert "본문 인용과 불일치" in text
    assert "240 MPa로 수정" in text
    assert "행 3·열 2" in text
