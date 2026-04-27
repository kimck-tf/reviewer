import sys
from pathlib import Path
import pytest


def test_cli_missing_pptx_returns_error(tmp_path, monkeypatch):
    """존재하지 않는 PPTX 입력 시 종료 코드 2."""
    monkeypatch.syspath_prepend(str(Path(__file__).parent.parent.parent))
    from extract_and_render import main

    rc = main([str(tmp_path / "nonexistent.pptx"), "--out", str(tmp_path / "out")])
    assert rc == 2
