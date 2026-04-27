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
