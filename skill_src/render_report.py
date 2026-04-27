"""CLI: <work_dir>/extracted.json + findings.json → review.md + review.html."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

from src.reporter_md import render as render_md
from src.reporter_html import render as render_html


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="findings.json + extracted.json → 마크다운·HTML 리포트"
    )
    parser.add_argument("work_dir", type=Path, help="extract_and_render.py가 생성한 작업 디렉토리")
    parser.add_argument("--out", type=Path, required=True, help="리포트 출력 디렉토리")
    args = parser.parse_args(argv)

    work_dir: Path = args.work_dir.resolve()
    out_dir: Path = args.out.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    extracted_path = work_dir / "extracted.json"
    findings_path = work_dir / "findings.json"

    if not extracted_path.exists():
        print(f"ERROR: extracted.json 없음: {extracted_path}", file=sys.stderr)
        return 2
    if not findings_path.exists():
        print(f"ERROR: findings.json 없음: {findings_path}", file=sys.stderr)
        return 2

    extracted = json.loads(extracted_path.read_text(encoding="utf-8"))
    findings = json.loads(findings_path.read_text(encoding="utf-8"))

    md_path = render_md(findings, extracted, out_dir / "review.md")
    html_path = render_html(findings, extracted, out_dir)

    print(f"Markdown: {md_path}")
    print(f"HTML: {html_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
