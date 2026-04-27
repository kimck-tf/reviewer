from __future__ import annotations
from pathlib import Path
from typing import Any


_SEVERITY_KO = {"critical": "🔴 Critical", "warning": "🟠 Warning", "info": "🔵 Info"}
_CATEGORY_KO = {
    "typo": "오타",
    "terminology": "용어 통일",
    "data": "데이터",
    "conclusion": "결론 검증",
    "improvement": "개선 제안",
    "logic": "논리·강도",
}


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
        lines.append("## 발견된 이슈")
        lines.append("")
        for f in findings.get("findings", []):
            lines.extend(_format_finding(f))
            lines.append("")

        # 슬라이드별 그룹
        lines.append("## 슬라이드별 이슈")
        lines.append("")
        slides_meta = {s["index"]: s.get("title", "") for s in extracted.get("slides", [])}
        by_slide: dict[int, list[dict[str, Any]]] = {}
        for f in findings.get("findings", []):
            by_slide.setdefault(f.get("slide_index", 0), []).append(f)
        for slide_idx in sorted(by_slide.keys()):
            title_val = slides_meta.get(slide_idx, "")
            heading = f"### 슬라이드 {slide_idx}"
            if title_val:
                heading += f": {title_val}"
            lines.append(heading)
            lines.append("")
            for f in by_slide[slide_idx]:
                sev = _SEVERITY_KO.get(f.get("severity", ""), "")
                cat = _CATEGORY_KO.get(f.get("category", ""), "")
                lines.append(f"- [{f.get('id', '?')}] {sev} · {cat} · {f.get('issue', '')}")
            lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


def _format_finding(f: dict[str, Any]) -> list[str]:
    """단일 finding을 마크다운 블록(라인 리스트)으로."""
    sev = _SEVERITY_KO.get(f.get("severity", "info"), f.get("severity", "info"))
    cat = _CATEGORY_KO.get(f.get("category", ""), f.get("category", ""))
    block = [
        f"### [{f.get('id', '?')}] {sev} · {cat} · {f.get('position_hint', '')}",
        "",
        f"**원문 인용**: \"{f.get('quoted_text', '')}\"",
        "",
        f"**문제**: {f.get('issue', '')}",
        "",
        f"**개선 제안**: {f.get('suggestion', '')}",
    ]
    evidence = f.get("evidence")
    if evidence:
        block.extend(["", f"**근거**: {evidence}"])
    return block
