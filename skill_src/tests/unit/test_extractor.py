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
