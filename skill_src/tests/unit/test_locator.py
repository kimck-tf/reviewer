from src.locator import to_hint, format_location


def test_top_left():
    pct = {"left": 0.05, "top": 0.05, "width": 0.3, "height": 0.1}
    hint = to_hint(pct)
    assert "좌측 상단" in hint or "좌상단" in hint


def test_top_right():
    pct = {"left": 0.70, "top": 0.10, "width": 0.25, "height": 0.10}
    hint = to_hint(pct)
    assert "우측 상단" in hint or "우상단" in hint


def test_center():
    pct = {"left": 0.40, "top": 0.40, "width": 0.20, "height": 0.20}
    hint = to_hint(pct)
    assert "중앙" in hint or "가운데" in hint


def test_bottom_full_width():
    pct = {"left": 0.00, "top": 0.85, "width": 1.00, "height": 0.10}
    hint = to_hint(pct)
    assert "하단" in hint


def test_includes_pct_numbers():
    pct = {"left": 0.70, "top": 0.15, "width": 0.25, "height": 0.10}
    hint = to_hint(pct)
    assert "70%" in hint
    assert "15%" in hint


def test_empty_returns_unknown():
    hint = to_hint({})
    assert "위치 미상" in hint or "unknown" in hint.lower()


def test_format_location_full():
    """슬라이드 번호 + 슬라이드 제목 + 도형 hint 통합 형식."""
    loc = format_location(
        slide_index=5,
        slide_title="결과",
        shape_type="TextBox",
        position_pct={"left": 0.7, "top": 0.15, "width": 0.25, "height": 0.1},
    )
    assert "슬라이드 5" in loc
    assert "결과" in loc
    assert "TextBox" in loc or "텍스트" in loc
    assert "우측 상단" in loc or "우상단" in loc


def test_format_location_no_title():
    loc = format_location(
        slide_index=3,
        slide_title="",
        shape_type="Table",
        position_pct={"left": 0.1, "top": 0.1, "width": 0.3, "height": 0.3},
    )
    assert "슬라이드 3" in loc


def test_format_location_matches_spec_example():
    """Spec §5.5.3 예시: '슬라이드 5 우측 상단 텍스트 박스 (좌측 70%, 상단 15%)' 정합성."""
    loc = format_location(
        slide_index=5,
        slide_title="",
        shape_type="TextBox",
        position_pct={"left": 0.70, "top": 0.15, "width": 0.25, "height": 0.10},
    )
    assert "슬라이드 5" in loc
    assert "우측 상단" in loc
    assert "텍스트 박스" in loc
    assert "70%" in loc
    assert "15%" in loc
