from __future__ import annotations
import shutil
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape


_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"

_SEVERITY_LABEL = {"critical": "Critical", "warning": "Warning", "info": "Info"}
_CATEGORY_LABEL = {
    "typo": "오타", "terminology": "용어 통일", "data": "데이터",
    "conclusion": "결론 검증", "improvement": "개선 제안", "logic": "논리·강도",
}


def render(findings: dict[str, Any], extracted: dict[str, Any], out_dir: Path) -> Path:
    """findings.json + extracted.json → HTML 리포트 파일 생성. 출력 HTML 경로 반환."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = out_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    # CSS 복사
    css_src = _TEMPLATES_DIR / "style.css"
    css_dst = assets_dir / "style.css"
    shutil.copy(css_src, css_dst)

    env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=select_autoescape(["html"]))
    template = env.get_template("report.html.j2")

    title = extracted.get("metadata", {}).get("title", "보고서")
    slide_count = extracted.get("metadata", {}).get("slide_count", 0)
    summary = findings.get("summary", {})
    total = summary.get("total_issues", 0)

    severity_counts = []
    by_sev = summary.get("by_severity", {})
    for sk in ("critical", "warning", "info"):
        if by_sev.get(sk, 0) > 0:
            severity_counts.append((sk, _SEVERITY_LABEL[sk], by_sev[sk]))

    # 이슈가 있는 슬라이드만 카드로 구성
    slides_meta = {s["index"]: s for s in extracted.get("slides", [])}
    by_slide: dict[int, list[dict]] = {}
    for f in findings.get("findings", []):
        by_slide.setdefault(f.get("slide_index", 0), []).append(f)

    thumb_out_dir = assets_dir / "thumbnails"
    thumb_out_dir.mkdir(parents=True, exist_ok=True)

    slides_with_findings = []
    for slide_idx in sorted(by_slide.keys()):
        meta = slides_meta.get(slide_idx, {})
        thumb_rel = None
        thumb_src = meta.get("thumbnail_path")
        if thumb_src and Path(thumb_src).exists():
            thumb_dst = thumb_out_dir / f"slide_{slide_idx:03d}.jpg"
            shutil.copy(thumb_src, thumb_dst)
            thumb_rel = f"assets/thumbnails/slide_{slide_idx:03d}.jpg"

        formatted_findings = []
        for f in by_slide[slide_idx]:
            formatted_findings.append({
                "id": f.get("id", "?"),
                "severity": f.get("severity", "info"),
                "severity_label": _SEVERITY_LABEL.get(f.get("severity", "info"), ""),
                "category_label": _CATEGORY_LABEL.get(f.get("category", ""), f.get("category", "")),
                "quoted_text": f.get("quoted_text", ""),
                "issue": f.get("issue", ""),
                "suggestion": f.get("suggestion", ""),
                "evidence": f.get("evidence", ""),
            })

        boxes = []
        for f in by_slide[slide_idx]:
            pct = f.get("position_pct") or {}
            if "left" in pct and "top" in pct:
                boxes.append({
                    "left": int(pct["left"] * 100),
                    "top": int(pct["top"] * 100),
                    "width": int(pct.get("width", 0) * 100),
                    "height": int(pct.get("height", 0) * 100),
                })

        slides_with_findings.append({
            "index": slide_idx,
            "title": meta.get("title", ""),
            "thumbnail_rel": thumb_rel,
            "boxes": boxes,
            "findings": formatted_findings,
        })

    rendered = template.render(
        title=title,
        slide_count=slide_count,
        total_issues=total,
        severity_counts=severity_counts,
        slides_with_findings=slides_with_findings,
    )
    out_path = out_dir / "review.html"
    out_path.write_text(rendered, encoding="utf-8")
    return out_path
