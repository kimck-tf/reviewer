from __future__ import annotations
import shutil
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape


_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"

_SEVERITY_LABEL = {"critical": "Critical", "warning": "Warning", "minor": "Minor", "ok": "OK"}
_CATEGORY_LABEL = {
    "typo": "오타", "terminology": "용어 통일", "data": "데이터",
    "conclusion": "결론 검증", "improvement": "개선 제안", "logic": "논리·강도",
}
_THESIS_ANSWERED_LABEL = {"yes": "답변됨", "partial": "부분 답변", "no": "미답변"}
_GRADE_LABEL = {
    "excellent": "Excellent", "good": "Good", "fair": "Fair", "needs_work": "Needs Work",
}
_DOC_AXES = [
    ("story_flow", "스토리라인 흐름 / 구성 균형"),
    ("decision_information", "결정 정보 충분성"),
    ("audience_fit", "청중 적합성"),
]


def _finding_categories(f: dict) -> list[str]:
    """finding의 카테고리 리스트. 복수 `categories[]` 우선, 없으면 단수 `category` 1개 리스트."""
    cats = f.get("categories")
    if cats:
        return list(cats)
    cat = f.get("category")
    return [cat] if cat else []


def _category_label(cats: list[str]) -> str:
    """카테고리 리스트 → ' / '로 결합된 한국어 라벨."""
    return " / ".join(_CATEGORY_LABEL.get(c, c) for c in cats)


def _source_ids_suffix(f: dict) -> str:
    """source_finding_ids가 2개 이상일 때 '(원본 A·B·C)' 형태 접미사. 그 외 빈 문자열."""
    src = f.get("source_finding_ids") or []
    if len(src) >= 2:
        return f"(원본 {'·'.join(src)})"
    return ""


def _build_document_review_ctx(doc: dict | None) -> dict | None:
    """document_review 객체 → 템플릿용 컨텍스트. 없거나 비어 있으면 None."""
    if not doc:
        return None
    axes_ctx = []
    for key, label in _DOC_AXES:
        sev = doc.get(f"{key}_severity")
        assessment = doc.get(f"{key}_assessment")
        if not (sev or assessment):
            continue
        axes_ctx.append({
            "label": label,
            "severity": sev or "",
            "severity_label": _SEVERITY_LABEL.get(sev, sev or ""),
            "assessment": assessment or "",
        })
    concerns_ctx = []
    for c in doc.get("cross_slide_concerns") or []:
        idxs = c.get("slide_indexes") or []
        concerns_ctx.append({
            "slide_indexes": idxs,
            "slide_indexes_label": ", ".join(f"슬라이드 {i}" for i in idxs),
            "severity": c.get("severity", "info"),
            "severity_label": _SEVERITY_LABEL.get(c.get("severity", "info"), ""),
            "issue": c.get("issue", ""),
            "suggestion": c.get("suggestion", ""),
        })
    grade = doc.get("overall_grade")
    return {
        "overall_grade": grade or "",
        "overall_grade_label": _GRADE_LABEL.get(grade, grade or ""),
        "overall_assessment": doc.get("overall_assessment", ""),
        "thesis_question": doc.get("thesis_question", ""),
        "thesis_answered": doc.get("thesis_answered", ""),
        "thesis_answered_label": _THESIS_ANSWERED_LABEL.get(doc.get("thesis_answered"), doc.get("thesis_answered", "")),
        "thesis_answer_summary": doc.get("thesis_answer_summary", ""),
        "axes": axes_ctx,
        "cross_slide_concerns": concerns_ctx,
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
    for sk in ("critical", "warning", "minor"):
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
                "category_label": _category_label(_finding_categories(f)),
                "source_ids_suffix": _source_ids_suffix(f),
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

    document_review_ctx = _build_document_review_ctx(findings.get("document_review"))

    # 카테고리별 그룹 — 통합 finding은 자신의 모든 카테고리 그룹에 중복 표시(A안)
    by_cat: dict[str, list[dict]] = {}
    for f in findings.get("findings", []):
        for c in _finding_categories(f):
            by_cat.setdefault(c, []).append(f)
    categories_with_findings = []
    for cat in sorted(by_cat.keys()):
        cat_items = []
        for f in by_cat[cat]:
            cat_items.append({
                "id": f.get("id", "?"),
                "source_ids_suffix": _source_ids_suffix(f),
                "severity": f.get("severity", "info"),
                "severity_label": _SEVERITY_LABEL.get(f.get("severity", "info"), ""),
                "slide_index": f.get("slide_index"),
                "issue": f.get("issue", ""),
            })
        categories_with_findings.append({
            "key": cat,
            "label": _CATEGORY_LABEL.get(cat, cat),
            "count": len(by_cat[cat]),
            "findings": cat_items,
        })

    rendered = template.render(
        title=title,
        slide_count=slide_count,
        total_issues=total,
        severity_counts=severity_counts,
        categories_with_findings=categories_with_findings,
        slides_with_findings=slides_with_findings,
        document_review=document_review_ctx,
    )
    out_path = out_dir / "review.html"
    out_path.write_text(rendered, encoding="utf-8")
    return out_path
