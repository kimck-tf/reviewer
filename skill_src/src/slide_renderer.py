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
    """모드 2: 임베디드 OLE 포함 변환.

    1차 MVP: PDF 임베디드만 별도 추출. PowerPoint/Excel/Word 임베디드는 슬라이드 캡처에 의존.
    """
    import win32com.client
    import pythoncom

    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    embed_dir = out_dir / "embedded"
    embed_dir.mkdir(parents=True, exist_ok=True)

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

        result: dict[int, list[Path]] = {}
        for slide in pres.Slides:
            idx = slide.SlideIndex
            paths: list[Path] = []

            # 1) 원본 슬라이드 캡처
            main_target = out_dir / f"slide_{idx:03d}.jpg"
            slide.Export(str(main_target), "JPG", width, height)
            paths.append(main_target)

            # 2) 임베디드 OLE 도형 순회
            for j, shape in enumerate(slide.Shapes, start=1):
                if getattr(shape, "Type", None) != 7:  # msoEmbeddedOLEObject
                    continue
                try:
                    progid = shape.OLEFormat.ProgID or ""
                except Exception:
                    progid = ""

                emb_paths = _process_embedded(
                    shape=shape,
                    progid=progid,
                    slide_idx=idx,
                    shape_idx=j,
                    embed_dir=embed_dir,
                    pdf_dpi=pdf_dpi,
                )
                paths.extend(emb_paths)

            result[idx] = paths
        return result
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


def _process_embedded(
    shape,
    progid: str,
    slide_idx: int,
    shape_idx: int,
    embed_dir: Path,
    pdf_dpi: int,
) -> list[Path]:
    """임베디드 OLE 도형 1개 처리. PDF만 별도 추출, 그 외는 ProgID 기록만."""
    out: list[Path] = []
    pid_lower = progid.lower()

    # PDF 임베디드: 별도 추출
    if "acroexch" in pid_lower or "pdf" in pid_lower:
        try:
            shape.OLEFormat.DoVerb(3)
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF DoVerb(3) 실패. slide_%03d 캡처에 의존.",
                slide_idx, slide_idx, exc_info=True,
            )
            return out

        tmp_pdf = embed_dir / f"slide_{slide_idx:03d}_emb{shape_idx:02d}.pdf"
        try:
            shape.OLEFormat.Object.SaveAs(str(tmp_pdf))
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF SaveAs 실패. slide_%03d 캡처에 의존.",
                slide_idx, slide_idx, exc_info=True,
            )
            return out

        if not tmp_pdf.exists():
            logger.warning("슬라이드 %d 임베디드 PDF 저장 결과 파일 없음.", slide_idx)
            return out

        try:
            import pdf2image
            pages = pdf2image.convert_from_path(str(tmp_pdf), dpi=pdf_dpi)
            for p_idx, page in enumerate(pages, start=1):
                out_png = embed_dir / f"slide_{slide_idx:03d}_emb{shape_idx:02d}_p{p_idx:02d}.png"
                page.save(out_png, "PNG")
                out.append(out_png)
        except Exception:
            logger.warning(
                "슬라이드 %d 임베디드 PDF → PNG 변환 실패 (Poppler 미설치 가능).",
                slide_idx, exc_info=True,
            )
        return out

    # PowerPoint 임베디드: 1차 MVP 미지원, 슬라이드 캡처에 의존
    if "powerpoint.show" in pid_lower:
        logger.info(
            "슬라이드 %d에 PowerPoint 임베디드(%s) 발견. 1차 MVP는 슬라이드 캡처에 의존.",
            slide_idx, progid,
        )
        return out

    # Excel/Word: 슬라이드 캡처에 의존
    if "excel" in pid_lower or "word" in pid_lower:
        logger.info(
            "슬라이드 %d에 Office 임베디드(%s) 발견. 슬라이드 캡처에 의존.",
            slide_idx, progid,
        )
        return out

    # 그 외 알 수 없는 ProgID
    logger.warning(
        "슬라이드 %d에 알 수 없는 임베디드 ProgID(%s). 슬라이드 캡처에 의존.",
        slide_idx, progid,
    )
    return out


def render(
    pptx_path: Path,
    out_dir: Path,
    extracted: dict[str, Any],
    width: int = 1928,
    height: int = 1080,
    pdf_dpi: int = 150,
    thumbnail_max_dim: int = 480,
) -> dict[int, list[Path]]:
    """자동 모드 선택 오케스트레이터.

    has_embedded가 하나라도 True면 모드 2, 아니면 모드 1 사용.
    썸네일은 항상 생성.
    """
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    has_any_embedded = any(s.get("has_embedded") for s in extracted.get("slides", []))

    if has_any_embedded:
        result = convert_pptx_with_embedded_to_images(
            pptx_path, out_dir, width=width, height=height, pdf_dpi=pdf_dpi
        )
    else:
        paths_dict = convert_pptx_to_images_without_embedding(
            pptx_path, out_dir, width=width, height=height
        )
        # 모드 1은 dict[int, Path] 반환 → dict[int, list[Path]]로 통일
        result = {idx: [p] for idx, p in paths_dict.items()}

    # 썸네일 생성
    thumb_dir = out_dir / "thumbnails"
    thumb_dir.mkdir(parents=True, exist_ok=True)
    _generate_thumbnails(result, thumb_dir, max_dim=thumbnail_max_dim)

    return result


def _generate_thumbnails(
    rendered: dict[int, list[Path]],
    thumb_dir: Path,
    max_dim: int,
) -> None:
    """원본 슬라이드 이미지(각 슬라이드의 첫 번째 경로)를 썸네일로 축소."""
    from PIL import Image

    for slide_idx, paths in rendered.items():
        if not paths:
            continue
        src = paths[0]
        target = thumb_dir / f"slide_{slide_idx:03d}.jpg"
        with Image.open(src) as img:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(target, "JPEG", quality=85)
