from __future__ import annotations
from pathlib import Path
from typing import Any


def render(findings: dict[str, Any], extracted: dict[str, Any], out_path: Path) -> Path:
    """findings.json + extracted.json → 마크다운 리포트 파일 생성. 출력 경로 반환."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    title = extracted.get("metadata", {}).get("title", "보고서")
    summary = findings.get("summary", {})
    total = summary.get("total_issues", 0)

    lines: list[str] = []
    lines.append(f"# 보고서 검토 결과: {title}")
    lines.append("")
    lines.append(f"슬라이드 수: {extracted.get('metadata', {}).get('slide_count', 0)}")
    lines.append(f"총 이슈: {total}개")
    lines.append("")

    if total == 0:
        lines.append("## 검토 결과 이슈 없음")
        lines.append("")
        lines.append("발견된 이슈가 없습니다. 보고서를 그대로 제출 가능합니다.")
    else:
        lines.append("(상세 내용은 후속 task에서 추가)")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path
