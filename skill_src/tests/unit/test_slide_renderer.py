import pytest
from pathlib import Path
from src.slide_renderer import (
    convert_pptx_to_images_without_embedding,
    convert_pptx_with_embedded_to_images,
    render,
)


def test_module_imports():
    """세 공개 함수가 import되고 NotImplementedError를 던지는지 확인."""
    with pytest.raises(NotImplementedError):
        convert_pptx_to_images_without_embedding(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        convert_pptx_with_embedded_to_images(Path("x.pptx"), Path("out"))
    with pytest.raises(NotImplementedError):
        render(Path("x.pptx"), Path("out"), {"slides": []})
