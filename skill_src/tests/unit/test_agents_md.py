from pathlib import Path

AGENTS_DIR = Path(__file__).parent.parent.parent.parent / "agents_src"

EXPECTED_AGENTS = [
    "report-reviewer-slide-analyzer",
]


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
