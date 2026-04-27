from pathlib import Path

AGENTS_DIR = Path(__file__).parent.parent.parent.parent / "agents_src"

EXPECTED_AGENTS = [
    "report-reviewer-slide-analyzer",
    "report-reviewer-typo",
    "report-reviewer-terminology",
    "report-reviewer-data",
    "report-reviewer-conclusion",
    "report-reviewer-improvement",
    "report-reviewer-logic",
]


CATEGORY_PREFIX = {
    "typo": "T",
    "terminology": "TM",
    "data": "D",
    "conclusion": "C",
    "improvement": "I",
    "logic": "L",
}


def test_agents_have_frontmatter():
    for name in EXPECTED_AGENTS:
        path = AGENTS_DIR / f"{name}.md"
        assert path.exists(), f"누락: {path}"
        text = path.read_text(encoding="utf-8")
        assert text.startswith("---\n")
        end = text.find("\n---\n", 4)
        assert end > 0
        fm = text[4:end]
        assert f"name: {name}" in fm
        assert "description:" in fm
        assert "tools:" in fm


def test_slide_analyzer_specifics():
    text = (AGENTS_DIR / "report-reviewer-slide-analyzer.md").read_text(encoding="utf-8")
    assert "tools: Read" in text
    for key in ["key_message", "claims", "data_points", "vision_observations"]:
        assert key in text


def test_category_subagents_have_correct_prefix():
    for cat, prefix in CATEGORY_PREFIX.items():
        path = AGENTS_DIR / f"report-reviewer-{cat}.md"
        text = path.read_text(encoding="utf-8")
        assert f"{prefix}001" in text, f"{cat} SA에 ID 예시 '{prefix}001' 누락"
        assert f'"{cat}"' in text or f"`{cat}`" in text


def test_category_subagents_have_output_schema():
    for cat in CATEGORY_PREFIX:
        path = AGENTS_DIR / f"report-reviewer-{cat}.md"
        text = path.read_text(encoding="utf-8")
        for key in ["id", "severity", "slide_index", "shape_id", "position_pct", "quoted_text", "issue", "suggestion"]:
            assert key in text, f"{cat} SA에 {key} 누락"
