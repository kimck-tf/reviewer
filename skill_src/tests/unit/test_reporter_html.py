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
