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


def test_findings_grouped_by_slide(tmp_path):
    findings = {
        "summary": {"total_issues": 3, "by_severity": {"critical": 1, "warning": 2}, "by_category": {"typo": 2, "data": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "슬라이드 1", "quoted_text": "오타1", "issue": "...", "suggestion": "..."},
            {"id": "F002", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "오타2", "issue": "...", "suggestion": "..."},
            {"id": "F003", "category": "data", "severity": "critical", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1", "quoted_text": "값", "issue": "...", "suggestion": "..."},
        ],
    }
    extracted = {"metadata": {"title": "테스트", "slide_count": 3}, "slides": [
        {"index": 1, "title": "표지"}, {"index": 2, "title": "본문"}, {"index": 3, "title": "결론"}
    ]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    assert "## 슬라이드별 이슈" in text
    assert "### 슬라이드 1" in text or "### 슬라이드 1: 표지" in text
    assert "### 슬라이드 2" in text or "### 슬라이드 2: 본문" in text
    s1_section_start = text.find("### 슬라이드 1")
    s2_section_start = text.find("### 슬라이드 2")
    s1_block = text[s1_section_start:s2_section_start]
    assert "F001" in s1_block
    assert "F003" in s1_block


def test_summary_header_and_category_grouping(tmp_path):
    findings = {
        "summary": {
            "total_issues": 3,
            "by_severity": {"critical": 1, "warning": 2, "info": 0},
            "by_category": {"typo": 2, "data": 1, "terminology": 0, "conclusion": 0, "improvement": 0, "logic": 0},
        },
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "슬라이드 1", "quoted_text": "x", "issue": "i1", "suggestion": "s1"},
            {"id": "F002", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "y", "issue": "i2", "suggestion": "s2"},
            {"id": "F003", "category": "data", "severity": "critical", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1", "quoted_text": "z", "issue": "i3", "suggestion": "s3"},
        ],
    }
    extracted = {"metadata": {"title": "T", "slide_count": 3}, "slides": [{"index": i} for i in [1, 2, 3]]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    assert "Critical: 1" in text or "🔴" in text
    assert "Warning: 2" in text or "🟠" in text
    assert "## 카테고리별 이슈" in text
    assert "### 오타" in text or "### 오타 (2건)" in text
    assert "### 데이터" in text or "### 데이터 (1건)" in text
