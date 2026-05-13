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
            "by_severity": {"critical": 1, "warning": 2, "minor": 0},
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


def test_document_review_section_rendered(tmp_path):
    """document_review 객체가 있으면 '문서 전체 평가' 섹션이 슬라이드 이슈보다 앞에 렌더된다."""
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {"id": "T001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "S1", "quoted_text": "x", "issue": "i", "suggestion": "s"},
        ],
        "document_review": {
            "thesis_question": "X 부품 내구 강도가 목표를 만족하는가?",
            "thesis_answered": "partial",
            "thesis_answer_summary": "결과 슬라이드는 있으나 결론 슬라이드 부재",
            "story_flow_severity": "warning",
            "story_flow_assessment": "분석 8장 대비 결론 0장으로 비대칭",
            "decision_information_severity": "critical",
            "decision_information_assessment": "권고 사항 누락",
            "audience_fit_severity": "ok",
            "audience_fit_assessment": "엔지니어 청중에 적절",
            "cross_slide_concerns": [
                {"slide_indexes": [3, 7], "severity": "critical",
                 "issue": "최대 응력 수치 불일치", "suggestion": "표 데이터 통일"},
            ],
            "overall_grade": "needs_work",
            "overall_assessment": "결론·권고 보강이 필요",
        },
    }
    extracted = {"metadata": {"title": "T", "slide_count": 8},
                 "slides": [{"index": 1, "title": "표지"}]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    # 섹션 헤더와 핵심 필드 노출
    assert "## 문서 전체 평가" in text
    assert "X 부품 내구 강도가 목표를 만족하는가?" in text
    assert "부분 답변" in text or "partial" in text
    assert "Needs Work" in text or "needs_work" in text
    assert "결론·권고 보강이 필요" in text
    # 5개 축 노출
    assert "스토리라인 흐름" in text
    assert "결정 정보 충분성" in text
    assert "청중 적합성" in text
    # cross-slide 모순 표기 (슬라이드 인덱스 + 제안)
    assert "슬라이드 3" in text and "슬라이드 7" in text
    assert "표 데이터 통일" in text
    # 위치 순서: 문서 전체 평가가 슬라이드별 이슈보다 먼저 와야 함
    assert text.find("## 문서 전체 평가") < text.find("## 슬라이드별 이슈")


def test_document_review_absent_keeps_legacy_layout(tmp_path):
    """document_review 키가 없으면 기존 출력과 동일하게 동작한다 (회귀 방지)."""
    findings = {
        "summary": {"total_issues": 0, "by_severity": {}, "by_category": {}},
        "findings": [],
    }
    extracted = {"metadata": {"title": "T", "slide_count": 1}, "slides": []}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")
    assert "## 문서 전체 평가" not in text


def test_merged_finding_displays_multiple_categories_and_source_ids(tmp_path):
    """Merger SA가 만든 통합 finding: categories[] 복수 라벨 + (원본 ...) 접미사."""
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"critical": 1}, "by_category": {}},
        "findings": [
            {
                "id": "F001",
                "categories": ["conclusion", "improvement", "logic"],
                "severity": "critical",
                "slide_index": 9, "shape_id": "s9_sh1",
                "position_hint": "슬라이드 9",
                "quoted_text": "환경변수 하나로 LLM을 무중단 전환할 수 있다",
                "issue": "종합 문제 설명",
                "suggestion": "종합 제안",
                "evidence": "종합 근거",
                "source_finding_ids": ["C002", "I005", "L001"],
            },
        ],
    }
    extracted = {"metadata": {"title": "T", "slide_count": 9}, "slides": [{"index": 9}]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")

    # 헤더에 원본 ID 표시
    assert "(원본 C002·I005·L001)" in text
    # 복수 카테고리 라벨 (한국어 슬래시 결합)
    assert "결론 검증 / 개선 제안 / 논리·강도" in text
    # 카테고리별 그룹: 세 카테고리 모두에 동일 finding 노출 (A안)
    assert text.count("F001") >= 3  # 발견된 이슈 + 슬라이드별 + 카테고리별 3개 그룹 = 5회 이상


def test_legacy_single_category_field_still_renders(tmp_path):
    """단수 `category` 필드만 있는 finding도 기존처럼 렌더 (호환성)."""
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {
                "id": "T001", "category": "typo", "severity": "warning",
                "slide_index": 1, "shape_id": "s1_sh1",
                "position_hint": "슬라이드 1", "quoted_text": "x",
                "issue": "오타", "suggestion": "수정", "evidence": "",
            },
        ],
    }
    extracted = {"metadata": {"title": "T", "slide_count": 1}, "slides": [{"index": 1}]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    text = out_path.read_text(encoding="utf-8")
    assert "오타" in text  # 카테고리 라벨
    assert "(원본" not in text  # 단일 카테고리는 원본 접미사 없음


def test_render_matches_golden(tmp_path, fixtures_dir):
    findings = {
        "summary": {"total_issues": 2, "by_severity": {"critical": 1, "warning": 1, "minor": 0},
                    "by_category": {"typo": 1, "data": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2 본문 영역", "quoted_text": "안응력",
             "issue": "오타: 안응력 → 인장응력으로 보임", "suggestion": "인장응력으로 수정", "evidence": ""},
            {"id": "F002", "category": "data", "severity": "critical", "slide_index": 5,
             "shape_id": "s5_sh3", "position_hint": "슬라이드 5 우측 상단 텍스트 박스", "quoted_text": "최대 응력 250 MPa",
             "issue": "표 5의 셀에는 240 MPa로 기재됨", "suggestion": "본문 240 MPa 또는 표 250 MPa로 통일",
             "evidence": "표 5(s5_sh4) 행 3·열 2 = 240"},
        ],
    }
    extracted = {"metadata": {"title": "테스트 보고서", "slide_count": 5},
                 "slides": [{"index": i, "title": f"슬라이드{i}"} for i in range(1, 6)]}
    out_path = tmp_path / "review.md"
    render(findings, extracted, out_path)
    actual = out_path.read_text(encoding="utf-8")
    expected = (fixtures_dir / "golden" / "reporter_md" / "sample_review.md").read_text(encoding="utf-8")
    assert actual == expected
