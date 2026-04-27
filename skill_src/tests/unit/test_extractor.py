import json
from pathlib import Path
import pytest
from src.extractor import extract


def test_extract_returns_metadata_and_slides(sample_text_only: Path):
    result = extract(sample_text_only)
    assert "metadata" in result
    assert "slides" in result
    assert result["metadata"]["slide_count"] == 5  # 표지+목차+3장
    assert len(result["slides"]) == 5


def test_extract_first_slide_has_title(sample_text_only: Path):
    result = extract(sample_text_only)
    slide1 = result["slides"][0]
    assert slide1["index"] == 1
    titles = [s["text"] for s in slide1["shapes"] if s["type"] == "Title"]
    assert any("프론트 서스펜션 강도 해석" in t for t in titles)


def test_extract_includes_file_path(sample_text_only: Path):
    result = extract(sample_text_only)
    assert result["metadata"]["file_path"] == str(sample_text_only)


def test_extract_table_shape(sample_with_table: Path):
    result = extract(sample_with_table)
    slide = result["slides"][0]
    tables = [s for s in slide["shapes"] if s["type"] == "Table"]
    assert len(tables) == 1
    tbl = tables[0]["table"]
    assert tbl is not None
    assert tbl["rows"] == 4
    assert tbl["cols"] == 3
    # 헤더 행
    assert tbl["cells"][0] == ["항목", "값", "단위"]
    # 데이터 행
    assert tbl["cells"][1] == ["최대 응력", "240", "MPa"]


def test_extract_position_pct(sample_text_only: Path):
    result = extract(sample_text_only)
    slide1 = result["slides"][0]
    for shape in slide1["shapes"]:
        pct = shape["position_pct"]
        assert "left" in pct and 0.0 <= pct["left"] <= 1.0
        assert "top" in pct and 0.0 <= pct["top"] <= 1.0
        assert "width" in pct and 0.0 <= pct["width"] <= 1.0
        assert "height" in pct and 0.0 <= pct["height"] <= 1.0


def test_extract_speaker_notes(sample_text_only: Path):
    result = extract(sample_text_only)
    # _make_fixtures.py에서 슬라이드 3~5에 노트 추가됨
    slide3 = result["slides"][2]
    assert "발표 노트" in slide3["notes"]


def test_extract_no_embedded_for_text_only(sample_text_only: Path):
    result = extract(sample_text_only)
    for slide in result["slides"]:
        assert slide["has_embedded"] is False


def test_extract_detects_embedded(sample_with_embedded: Path):
    result = extract(sample_with_embedded)
    has_any_embedded = any(s["has_embedded"] for s in result["slides"])
    assert has_any_embedded is True
    for slide in result["slides"]:
        if slide["has_embedded"]:
            ole_shapes = [s for s in slide["shapes"] if s["type"] == "EmbeddedOLE"]
            assert len(ole_shapes) >= 1
            assert ole_shapes[0]["embedded_progid"]  # non-empty


def test_extract_text_only_matches_golden(sample_text_only: Path, fixtures_dir: Path):
    result = extract(sample_text_only)
    # 변동 가능 필드 정규화 (골든과 동일하게)
    result["metadata"]["file_path"] = "<FIXTURE>"
    result["metadata"]["created"] = "<NORMALIZED>"
    result["metadata"]["modified"] = "<NORMALIZED>"
    golden_path = fixtures_dir / "golden" / "extractor_outputs" / "sample_text_only.json"
    expected = json.loads(golden_path.read_text(encoding="utf-8"))
    assert result == expected
