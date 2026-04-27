import pytest
from pathlib import Path
from src.slide_renderer import (
    convert_pptx_to_images_without_embedding,
    convert_pptx_with_embedded_to_images,
    render,
)


def test_module_imports():
    """모드 2와 render는 아직 NotImplementedError."""
    with pytest.raises(NotImplementedError):
        convert_pptx_with_embedded_to_images(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        render(Path("x.pptx"), Path("out"), {"slides": []})


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
