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

    rendered = template.render(
        title=title,
        slide_count=slide_count,
        total_issues=total,
        severity_counts=severity_counts,
        slides_with_findings=[],  # Task 5.2에서 채움
    )
    out_path = out_dir / "review.html"
    out_path.write_text(rendered, encoding="utf-8")
    return out_path
