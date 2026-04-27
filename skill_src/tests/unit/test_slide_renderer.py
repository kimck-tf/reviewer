import pytest
from pathlib import Path
from unittest.mock import MagicMock
from PIL import Image as _PILImage
from src.slide_renderer import (
    convert_pptx_to_images_without_embedding,
    convert_pptx_with_embedded_to_images,
    render,
)


def test_module_imports():
    """모든 공개 함수가 import 가능 (구현 완료)."""
    assert callable(convert_pptx_to_images_without_embedding)
    assert callable(convert_pptx_with_embedded_to_images)
    assert callable(render)


def test_mode1_creates_one_image_per_slide(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=3)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    result = convert_pptx_to_images_without_embedding(pptx_path, out_dir)

    assert set(result.keys()) == {1, 2, 3}
    for idx, p in result.items():
        assert p.exists()
        assert p.suffix.lower() == ".jpg"


def test_mode1_calls_export_with_correct_size(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=2)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    convert_pptx_to_images_without_embedding(pptx_path, out_dir, width=1280, height=720)

    for sl in slides:
        args = sl.Export.call_args
        assert args[0][1] == "JPG"
        assert args[0][2] == 1280
        assert args[0][3] == 720


def test_mode1_quits_powerpoint_even_on_error(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=1)
    slides[0].Export.side_effect = RuntimeError("export failed")
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    with pytest.raises(RuntimeError):
        convert_pptx_to_images_without_embedding(pptx_path, out_dir)

    pres.Close.assert_called()
    app.Quit.assert_called()


def test_mode2_no_embedded_acts_like_mode1(mock_powerpoint_com, tmp_path):
    app, pres, slides = mock_powerpoint_com(slide_count=2)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    # dict[int, list[Path]] 반환, 임베디드 없으면 각 리스트 길이 1
    assert set(result.keys()) == {1, 2}
    for paths in result.values():
        assert len(paths) == 1
        assert paths[0].exists()


def test_mode2_pptshow_uses_slide_capture_only(mock_powerpoint_com, tmp_path, caplog):
    """PowerPoint 임베디드는 별도 추출 안 함 (1차 MVP), 슬라이드 캡처만."""
    import logging
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["PowerPoint.Show.12"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    with caplog.at_level(logging.INFO, logger="src.slide_renderer"):
        result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    # 슬라이드 캡처만 (임베디드 별도 이미지 없음)
    assert result[1] == [out_dir / "slide_001.jpg"]
    # ProgID가 로그에 INFO로 기록되었는지
    assert any("PowerPoint.Show" in rec.message for rec in caplog.records)
    # DoVerb(1) 호출되지 않아야 함
    embedded_shape = slides[0].Shapes[0]
    embedded_shape.OLEFormat.DoVerb.assert_not_called()


def test_mode2_pdf_progid_extracts_separately(mock_powerpoint_com, tmp_path, monkeypatch):
    """PDF 임베디드는 DoVerb(3) + SaveAs + pdf2image 호출, 별도 PNG 생성."""
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["AcroExch.Document.DC"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    # SaveAs가 실제로 PDF를 생성하도록 side_effect
    def _fake_saveas(path):
        Path(path).write_bytes(b"%PDF-1.4 dummy")
    slides[0].Shapes[0].OLEFormat.Object.SaveAs.side_effect = _fake_saveas

    fake_pages = [_PILImage.new("RGB", (200, 200), color=(255, 0, 0))]
    fake_convert = MagicMock(return_value=fake_pages)
    monkeypatch.setattr("pdf2image.convert_from_path", fake_convert)

    result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    embedded_shape = slides[0].Shapes[0]
    embedded_shape.OLEFormat.DoVerb.assert_any_call(3)
    fake_convert.assert_called_once()
    # 슬라이드 캡처 + 임베디드 PDF 추출 = 2개 경로
    assert len(result[1]) == 2
    # 두 번째 경로는 임베디드 PNG
    assert result[1][1].suffix.lower() == ".png"


def test_mode2_pdf_saveas_failure_logs_warning(mock_powerpoint_com, tmp_path, caplog):
    """PDF 임베디드 SaveAs 실패 시 logging.warning."""
    import logging
    app, pres, slides = mock_powerpoint_com(
        slide_count=1,
        embedded_progids_per_slide={1: ["AcroExch.Document.DC"]},
    )
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"

    slides[0].Shapes[0].OLEFormat.Object.SaveAs.side_effect = RuntimeError("Acrobat blocked")

    with caplog.at_level(logging.WARNING, logger="src.slide_renderer"):
        result = convert_pptx_with_embedded_to_images(pptx_path, out_dir)

    assert any("PDF 임베디드" in rec.message or "SaveAs" in rec.message for rec in caplog.records)
    # 슬라이드 캡처는 성공
    assert result[1][0].exists()


def test_render_calls_mode1_when_no_embedded(monkeypatch, tmp_path):
    """render가 has_embedded=False 시 mode 1을 호출, mode 2는 호출 안 함."""
    from PIL import Image as _PI
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    fake_jpg = out_dir / "slide_001.jpg"
    _PI.new("RGB", (100, 100)).save(fake_jpg)

    mode1_calls = []
    def fake_mode1(pptx, out, **kwargs):
        mode1_calls.append((pptx, out))
        return {1: fake_jpg}

    mode2_calls = []
    def fake_mode2(pptx, out, **kwargs):
        mode2_calls.append((pptx, out))
        return {}

    monkeypatch.setattr("src.slide_renderer.convert_pptx_to_images_without_embedding", fake_mode1)
    monkeypatch.setattr("src.slide_renderer.convert_pptx_with_embedded_to_images", fake_mode2)

    extracted = {"slides": [{"index": 1, "has_embedded": False}]}
    render(tmp_path / "in.pptx", out_dir, extracted)

    assert len(mode1_calls) == 1
    assert len(mode2_calls) == 0


def test_render_calls_mode2_when_any_embedded(monkeypatch, tmp_path):
    """render가 has_embedded=True가 하나라도 있으면 mode 2를 호출."""
    from PIL import Image as _PI
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    fake_jpg1 = out_dir / "slide_001.jpg"
    fake_jpg2 = out_dir / "slide_002.jpg"
    _PI.new("RGB", (100, 100)).save(fake_jpg1)
    _PI.new("RGB", (100, 100)).save(fake_jpg2)

    mode1_calls = []
    def fake_mode1(pptx, out, **kwargs):
        mode1_calls.append((pptx, out))
        return {}

    mode2_calls = []
    def fake_mode2(pptx, out, **kwargs):
        mode2_calls.append((pptx, out))
        return {1: [fake_jpg1], 2: [fake_jpg2]}

    monkeypatch.setattr("src.slide_renderer.convert_pptx_to_images_without_embedding", fake_mode1)
    monkeypatch.setattr("src.slide_renderer.convert_pptx_with_embedded_to_images", fake_mode2)

    extracted = {"slides": [
        {"index": 1, "has_embedded": False},
        {"index": 2, "has_embedded": True},
    ]}
    render(tmp_path / "in.pptx", out_dir, extracted)

    assert len(mode2_calls) == 1
    assert len(mode1_calls) == 0


def test_render_creates_thumbnails(mock_powerpoint_com, tmp_path):
    """실제 mode 1 + 썸네일 생성 통합 검증."""
    app, pres, slides = mock_powerpoint_com(slide_count=1)
    pptx_path = tmp_path / "in.pptx"
    pptx_path.write_bytes(b"")
    out_dir = tmp_path / "out"
    extracted = {"slides": [{"index": 1, "has_embedded": False}]}

    render(pptx_path, out_dir, extracted, thumbnail_max_dim=300)

    thumb = out_dir / "thumbnails" / "slide_001.jpg"
    assert thumb.exists()
    from PIL import Image
    with Image.open(thumb) as img:
        assert max(img.size) <= 300
