from pathlib import Path


def test_skill_md_has_frontmatter():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    assert text.startswith("---\n")
    end = text.find("\n---\n", 4)
    assert end > 0, "frontmatter 종료 마커 없음"
    fm = text[4:end]
    assert "name: report-reviewer" in fm
    assert "description:" in fm


def test_skill_md_workflow_steps_present():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    for step_heading in [
        "### Step 0", "### Step 1", "### Step 2",
        "### Step 3", "### Step 4", "### Step 5",
    ]:
        assert step_heading in text, f"누락: {step_heading}"


def test_skill_md_subagent_names():
    skill_md = Path(__file__).parent.parent.parent / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    for sa in [
        "report-reviewer-slide-analyzer",
        "report-reviewer-typo",
        "report-reviewer-terminology",
        "report-reviewer-data",
        "report-reviewer-conclusion",
        "report-reviewer-improvement",
        "report-reviewer-logic",
    ]:
        assert sa in text, f"SKILL.md에 {sa} 참조 없음"
