from __future__ import annotations
from pathlib import Path
from typing import Any


_SEVERITY_KO = {"critical": "🔴 Critical", "warning": "🟠 Warning", "info": "🔵 Info", "ok": "🟢 OK"}
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

    # 요약 헤더 보강
    by_sev = summary.get("by_severity", {})
    if by_sev:
        sev_parts = []
        for sk, ko in [("critical", "Critical"), ("warning", "Warning"), ("info", "Info")]:
            n = by_sev.get(sk, 0)
            if n > 0:
                sev_parts.append(f"{_SEVERITY_KO.get(sk, ko)}: {n}")
        if sev_parts:
            lines.append("심각도: " + " · ".join(sev_parts))
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
                cat = _CATEGORY_KO.get(f.get("category", ""), "")
                lines.append(f"- [{f.get('id', '?')}] {sev} · {cat} · {f.get('issue', '')}")
            lines.append("")

        # 카테고리별 그룹
        lines.append("## 카테고리별 이슈")
        lines.append("")
        by_cat: dict[str, list[dict[str, Any]]] = {}
        for f in findings.get("findings", []):
            by_cat.setdefault(f.get("category", ""), []).append(f)
        for cat in sorted(by_cat.keys()):
            cat_ko = _CATEGORY_KO.get(cat, cat)
            lines.append(f"### {cat_ko} ({len(by_cat[cat])}건)")
            lines.append("")
            for f in by_cat[cat]:
                sev = _SEVERITY_KO.get(f.get("severity", ""), "")
                lines.append(f"- [{f.get('id', '?')}] {sev} · 슬라이드 {f.get('slide_index')} · {f.get('issue', '')}")
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
