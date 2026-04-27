"""슬라이드 → 이미지 변환 (PowerPoint COM 기반).

두 모드:
  - convert_pptx_to_images_without_embedding: 단순 변환 (slide.Export)
  - convert_pptx_with_embedded_to_images: 임베디드 OLE 포함 변환 (DoVerb 분기)

오케스트레이터:
  - render: extracted dict의 has_embedded 플래그 기반 슬라이드별 자동 모드 선택
"""
from __future__ import annotations
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


def convert_pptx_to_images_without_embedding(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
) -> dict[int, Path]:
    """모드 1: 단순 PPTX → 슬라이드 JPG 변환. 슬라이드 인덱스 → 출력 경로 dict."""
    import win32com.client
    import pythoncom

    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.Visible = 0
        except Exception:
            logger.debug("PowerPoint.Visible=0 거부됨, WithWindow=False에 의존")

        pres = app.Presentations.Open(str(pptx_path), WithWindow=False)

        out_paths: dict[int, Path] = {}
        for slide in pres.Slides:
            idx = slide.SlideIndex
            target = out_dir / f"slide_{idx:03d}.jpg"
            slide.Export(str(target), "JPG", width, height)
            out_paths[idx] = target
        return out_paths
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            logger.warning("Presentation.Close() 실패", exc_info=True)
        try:
            if app is not None:
                app.Quit()
        except Exception:
            logger.warning("Application.Quit() 실패", exc_info=True)
        pythoncom.CoUninitialize()


def convert_pptx_with_embedded_to_images(
    pptx_path: Path,
    out_dir: Path,
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
) -> dict[int, list[Path]]:
    """모드 2: 임베디드 OLE 포함 변환. 슬라이드 인덱스 → [원본 슬라이드 경로, *임베디드 경로들]."""
    raise NotImplementedError


def render(
    pptx_path: Path,
    out_dir: Path,
    extracted: dict[str, Any],
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
    thumbnail_max_dim: int = 480,
) -> dict[int, list[Path]]:
    """자동 모드 선택 오케스트레이터. extracted dict의 has_embedded 플래그로 슬라이드별 분기."""
    raise NotImplementedError
