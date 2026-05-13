from __future__ import annotations
from pathlib import Path
from typing import Any


_SEVERITY_KO = {"critical": "🔴 Critical", "warning": "🟠 Warning", "minor": "🔵 Minor", "ok": "🟢 OK"}
_SEVERITY_LEGEND = "🔴 Critical(반드시 수정) · 🟠 Warning(수정 권장) · 🔵 Minor(참고)"
_CATEGORY_KO = {
    "typo": "오타",
    "terminology": "용어 통일",
    "data": "데이터",
    "conclusion": "결론 검증",
    "improvement": "개선 제안",
    "logic": "논리·강도",
}
_THESIS_ANSWERED_KO = {"yes": "✅ 답변됨", "partial": "🟡 부분 답변", "no": "❌ 미답변"}
_GRADE_KO = {
    "excellent": "🟢 Excellent",
    "good": "🔵 Good",
    "fair": "🟡 Fair",
    "needs_work": "🔴 Needs Work",
}
_DOC_AXES = [
    ("story_flow", "스토리라인 흐름 / 구성 균형"),
    ("decision_information", "결정 정보 충분성"),
    ("audience_fit", "청중 적합성"),
]


def _finding_categories(f: dict[str, Any]) -> list[str]:
    """finding의 카테고리 리스트를 반환. 복수 `categories[]` 우선, 없으면 단수 `category` 1개 리스트."""
    cats = f.get("categories")
    if cats:
        return list(cats)
    cat = f.get("category")
    return [cat] if cat else []


def _category_label(cats: list[str]) -> str:
    """카테고리 리스트 → ' / '로 결합된 한국어 라벨."""
    return " / ".join(_CATEGORY_KO.get(c, c) for c in cats)


def _source_ids_suffix(f: dict[str, Any]) -> str:
    """source_finding_ids가 자기 자신 1개가 아닐 때 '(원본 A·B·C)' 형태 접미사 반환. 그 외 빈 문자열."""
    src = f.get("source_finding_ids") or []
    if len(src) >= 2:
        return f" (원본 {'·'.join(src)})"
    return ""


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

    # 심각도 범례 (항상 표시)
    lines.append(f"심각도 척도: {_SEVERITY_LEGEND}")
    lines.append("")

    # 요약 헤더 — 발견 건수 분포
    by_sev = summary.get("by_severity", {})
    if by_sev:
        sev_parts = []
        for sk, ko in [("critical", "Critical"), ("warning", "Warning"), ("minor", "Minor")]:
            n = by_sev.get(sk, 0)
            if n > 0:
                sev_parts.append(f"{_SEVERITY_KO.get(sk, ko)}: {n}")
        if sev_parts:
            lines.append("발견 건수: " + " · ".join(sev_parts))
            lines.append("")

    # 문서 전체 평가 (거시 SA 결과) — 슬라이드별 이슈 위에 배치
    doc = findings.get("document_review")
    if doc:
        lines.extend(_format_document_review(doc))
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
                cat_label = _category_label(_finding_categories(f))
                lines.append(f"- [{f.get('id', '?')}]{_source_ids_suffix(f)} {sev} · {cat_label} · {f.get('issue', '')}")
            lines.append("")

        # 카테고리별 그룹 — 통합 finding은 자신의 모든 카테고리 그룹에 중복 표시(A안)
        lines.append("## 카테고리별 이슈")
        lines.append("")
        by_cat: dict[str, list[dict[str, Any]]] = {}
        for f in findings.get("findings", []):
            for c in _finding_categories(f):
                by_cat.setdefault(c, []).append(f)
        for cat in sorted(by_cat.keys()):
            cat_ko = _CATEGORY_KO.get(cat, cat)
            lines.append(f"### {cat_ko} ({len(by_cat[cat])}건)")
            lines.append("")
            for f in by_cat[cat]:
                sev = _SEVERITY_KO.get(f.get("severity", ""), "")
                lines.append(f"- [{f.get('id', '?')}]{_source_ids_suffix(f)} {sev} · 슬라이드 {f.get('slide_index')} · {f.get('issue', '')}")
            lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


def _format_document_review(doc: dict[str, Any]) -> list[str]:
    """document_review 객체 → 마크다운 블록(라인 리스트)."""
    lines: list[str] = ["## 문서 전체 평가", ""]

    grade = doc.get("overall_grade")
    if grade:
        lines.append(f"**종합 등급**: {_GRADE_KO.get(grade, grade)}")
        lines.append("")

    overall = doc.get("overall_assessment")
    if overall:
        lines.append(f"> {overall}")
        lines.append("")

    # 핵심 질문 + 답변 여부
    thesis_q = doc.get("thesis_question")
    thesis_a = doc.get("thesis_answered")
    if thesis_q or thesis_a:
        lines.append("### 핵심 질문 답변 여부")
        lines.append("")
        if thesis_q:
            lines.append(f"- **핵심 질문**: {thesis_q}")
        if thesis_a:
            lines.append(f"- **답변 여부**: {_THESIS_ANSWERED_KO.get(thesis_a, thesis_a)}")
        summary_text = doc.get("thesis_answer_summary")
        if summary_text:
            lines.append(f"- **판단 근거**: {summary_text}")
        lines.append("")

    # 5개 축 중 3개 단일 항목 (story_flow / decision_information / audience_fit)
    for key, label in _DOC_AXES:
        sev = doc.get(f"{key}_severity")
        assessment = doc.get(f"{key}_assessment")
        if not (sev or assessment):
            continue
        sev_label = _SEVERITY_KO.get(sev, sev or "")
        lines.append(f"### {label}")
        lines.append("")
        if sev_label:
            lines.append(f"- **평가**: {sev_label}")
        if assessment:
            lines.append(f"- {assessment}")
        lines.append("")

    # 슬라이드 간 모순·중복
    concerns = doc.get("cross_slide_concerns") or []
    if concerns:
        lines.append("### 슬라이드 간 모순·중복")
        lines.append("")
        for c in concerns:
            sev = _SEVERITY_KO.get(c.get("severity", "info"), c.get("severity", ""))
            idxs = c.get("slide_indexes") or []
            idx_label = ", ".join(f"슬라이드 {i}" for i in idxs) if idxs else "(슬라이드 미지정)"
            lines.append(f"- {sev} · {idx_label}")
            issue = c.get("issue")
            if issue:
                lines.append(f"  - **문제**: {issue}")
            suggestion = c.get("suggestion")
            if suggestion:
                lines.append(f"  - **제안**: {suggestion}")
        lines.append("")

    return lines


def _format_finding(f: dict[str, Any]) -> list[str]:
    """단일 finding을 마크다운 블록(라인 리스트)으로. categories 복수와 source_finding_ids 통합 표시 포함."""
    sev = _SEVERITY_KO.get(f.get("severity", "info"), f.get("severity", "info"))
    cat_label = _category_label(_finding_categories(f))
    block = [
        f"### [{f.get('id', '?')}]{_source_ids_suffix(f)} {sev} · {cat_label} · {f.get('position_hint', '')}",
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
