"""CLI: PPTX → extracted.json + 슬라이드 이미지 + 슬라이드별 SA 입력 파일."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

from src.config import Config
from src.extractor import extract
from src.slide_renderer import render


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="PPTX 추출 + 슬라이드 이미지 변환 + 1단계 SA 입력 생성"
    )
    parser.add_argument("pptx_path", type=Path, help="입력 PPTX 파일 경로")
    parser.add_argument("--out", type=Path, required=True, help="출력 작업 디렉토리")
    args = parser.parse_args(argv)

    pptx_path: Path = args.pptx_path.resolve()
    work_dir: Path = args.out.resolve()
    work_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        print(f"ERROR: PPTX 파일 없음: {pptx_path}", file=sys.stderr)
        return 2

    cfg = Config.from_env(dict(__import__("os").environ))

    # ① 구조 추출
    print(f"[1/3] 구조 추출 중: {pptx_path.name}")
    extracted = extract(pptx_path)

    # ② 슬라이드 이미지 변환 (자동 모드 선택)
    print("[2/3] 슬라이드 이미지 변환 중 (모드 자동 선택)")
    rendered = render(
        pptx_path, work_dir, extracted,
        width=cfg.image_max_dim, height=int(cfg.image_max_dim * 9 / 16),
        pdf_dpi=cfg.pdf_dpi, thumbnail_max_dim=cfg.thumbnail_max_dim,
    )

    # extracted에 image_path / thumbnail_path / embedded_image_paths 채우기
    thumb_dir = work_dir / "thumbnails"
    for slide in extracted["slides"]:
        idx = slide["index"]
        paths = rendered.get(idx, [])
        if paths:
            slide["image_path"] = str(paths[0])
            slide["embedded_image_paths"] = [str(p) for p in paths[1:]]
        thumb = thumb_dir / f"slide_{idx:03d}.jpg"
        if thumb.exists():
            slide["thumbnail_path"] = str(thumb)

    # extracted.json 저장
    extracted_path = work_dir / "extracted.json"
    extracted_path.write_text(
        json.dumps(extracted, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  → {extracted_path}")

    # ③ 슬라이드별 SA 입력 JSON 생성
    print("[3/3] 슬라이드별 SA 입력 JSON 생성")
    inputs_dir = work_dir / "slide_inputs"
    inputs_dir.mkdir(parents=True, exist_ok=True)
    for slide in extracted["slides"]:
        idx = slide["index"]
        target = inputs_dir / f"slide_{idx:03d}.json"
        target.write_text(
            json.dumps(slide, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(f"  → {inputs_dir}/ ({len(extracted['slides'])}개)")

    print("완료. 다음 단계: SKILL.md 워크플로 따라 1단계 SA dispatch.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
