from __future__ import annotations
from typing import Mapping


def to_hint(position_pct: Mapping[str, float]) -> str:
    """슬라이드 내 정규화 좌표(0~1)를 한국어 위치 hint 문자열로 변환.

    예: {"left":0.7,"top":0.15,"width":0.25,"height":0.1}
        → "우측 상단 (좌측 70%, 상단 15%)"
    """
    if not position_pct or "left" not in position_pct or "top" not in position_pct:
        return "위치 미상"

    left = position_pct["left"]
    top = position_pct["top"]
    width = position_pct.get("width", 0.0)

    # 가로 영역
    center_x = left + width / 2
    if center_x < 0.33:
        h_zone = "좌측"
    elif center_x < 0.67:
        h_zone = "중앙"
    else:
        h_zone = "우측"

    # 세로 영역
    center_y = top + position_pct.get("height", 0.0) / 2
    if center_y < 0.33:
        v_zone = "상단"
    elif center_y < 0.67:
        v_zone = "중단"
    else:
        v_zone = "하단"

    # 중앙·중단이면 "가운데"
    if h_zone == "중앙" and v_zone == "중단":
        zone = "가운데"
    else:
        zone = f"{h_zone} {v_zone}"

    return f"{zone} (좌측 {int(left*100)}%, 상단 {int(top*100)}%)"
