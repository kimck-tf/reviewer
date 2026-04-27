import os
from src.config import Config


def test_default_config():
    cfg = Config.from_env({})
    assert cfg.batch_size == 5
    assert cfg.image_max_dim == 1928
    assert cfg.pdf_dpi == 150
    assert cfg.thumbnail_max_dim == 480


def test_env_override_batch_size():
    cfg = Config.from_env({"REPORT_REVIEWER_BATCH_SIZE": "10"})
    assert cfg.batch_size == 10


def test_env_override_image_dim():
    cfg = Config.from_env({"REPORT_REVIEWER_IMAGE_MAX_DIM": "1280"})
    assert cfg.image_max_dim == 1280


def test_invalid_int_raises():
    import pytest
    with pytest.raises(ValueError):
        Config.from_env({"REPORT_REVIEWER_BATCH_SIZE": "abc"})
