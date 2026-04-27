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


def test_html_includes_slide_cards_with_thumbnails(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "슬라이드 2", "quoted_text": "오타",
             "issue": "오타입니다", "suggestion": "수정 필요", "evidence": ""}
        ],
    }
    thumb = tmp_path / "ws" / "thumbnails" / "slide_002.jpg"
    thumb.parent.mkdir(parents=True, exist_ok=True)
    thumb.write_bytes(b"fake jpg")
    extracted = {
        "metadata": {"title": "T", "slide_count": 3},
        "slides": [
            {"index": 1, "title": "표지", "thumbnail_path": None},
            {"index": 2, "title": "본문", "thumbnail_path": str(thumb)},
            {"index": 3, "title": "결론", "thumbnail_path": None},
        ],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    assert "슬라이드 2" in text
    assert "F001" in text
    assert (out_dir / "assets" / "thumbnails" / "slide_002.jpg").exists()
    assert "assets/thumbnails/slide_002.jpg" in text


def test_html_includes_position_boxes(tmp_path):
    findings = {
        "summary": {"total_issues": 1, "by_severity": {"warning": 1}, "by_category": {"typo": 1}},
        "findings": [
            {"id": "F001", "category": "typo", "severity": "warning", "slide_index": 1,
             "shape_id": "s1_sh3", "position_hint": "슬라이드 1 우측 상단",
             "position_pct": {"left": 0.7, "top": 0.15, "width": 0.25, "height": 0.10},
             "quoted_text": "x", "issue": "i", "suggestion": "s", "evidence": ""}
        ],
    }
    thumb = tmp_path / "ws" / "thumbnails" / "slide_001.jpg"
    thumb.parent.mkdir(parents=True, exist_ok=True)
    thumb.write_bytes(b"fake")
    extracted = {
        "metadata": {"title": "T", "slide_count": 1},
        "slides": [{"index": 1, "title": "표지", "thumbnail_path": str(thumb)}],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    assert "position-box" in text
    assert "left:70" in text or "left: 70" in text
    assert "top:15" in text or "top: 15" in text
    assert "width:25" in text or "width: 25" in text


def test_html_full_integration(tmp_path):
    """모든 카테고리·심각도가 섞인 findings로 HTML 렌더 + DOM 구조 검증."""
    thumb1 = tmp_path / "ws" / "thumbnails" / "slide_001.jpg"
    thumb2 = tmp_path / "ws" / "thumbnails" / "slide_002.jpg"
    for t in (thumb1, thumb2):
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_bytes(b"fake")

    findings = {
        "summary": {
            "total_issues": 3,
            "by_severity": {"critical": 1, "warning": 1, "info": 1},
            "by_category": {"typo": 1, "data": 1, "logic": 1},
        },
        "findings": [
            {"id": "F001", "category": "typo", "severity": "info", "slide_index": 1,
             "shape_id": "s1_sh1", "position_hint": "S1",
             "position_pct": {"left": 0.1, "top": 0.1, "width": 0.3, "height": 0.1},
             "quoted_text": "오타", "issue": "i1", "suggestion": "s1", "evidence": ""},
            {"id": "F002", "category": "data", "severity": "critical", "slide_index": 2,
             "shape_id": "s2_sh1", "position_hint": "S2",
             "position_pct": {"left": 0.5, "top": 0.5, "width": 0.4, "height": 0.2},
             "quoted_text": "데이터", "issue": "i2", "suggestion": "s2", "evidence": "표 ..."},
            {"id": "F003", "category": "logic", "severity": "warning", "slide_index": 2,
             "shape_id": "s2_sh2", "position_hint": "S2-2",
             "position_pct": {"left": 0.0, "top": 0.8, "width": 1.0, "height": 0.2},
             "quoted_text": "결론", "issue": "i3", "suggestion": "s3", "evidence": ""},
        ],
    }
    extracted = {
        "metadata": {"title": "통합 테스트", "slide_count": 3},
        "slides": [
            {"index": 1, "title": "표지", "thumbnail_path": str(thumb1)},
            {"index": 2, "title": "본문", "thumbnail_path": str(thumb2)},
            {"index": 3, "title": "결론", "thumbnail_path": None},
        ],
    }
    out_dir = tmp_path / "out"
    render(findings, extracted, out_dir)

    text = (out_dir / "review.html").read_text(encoding="utf-8")
    for fid in ("F001", "F002", "F003"):
        assert fid in text
    assert "Critical: 1" in text
    assert "Warning: 1" in text
    assert "Info: 1" in text
    assert text.count('<div class="slide-card">') == 2  # 슬라이드 1, 2만
    assert text.count("position-box") == 3  # 각 finding당 1개
    assert (out_dir / "assets" / "style.css").exists()
    assert (out_dir / "assets" / "thumbnails" / "slide_001.jpg").exists()
    assert (out_dir / "assets" / "thumbnails" / "slide_002.jpg").exists()
