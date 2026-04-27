from src.locator import to_hint


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
