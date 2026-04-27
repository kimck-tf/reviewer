from __future__ import annotations
from dataclasses import dataclass
from typing import Mapping


@dataclass(frozen=True)
class Config:
    batch_size: int = 5
    image_max_dim: int = 1928
    pdf_dpi: int = 150
    thumbnail_max_dim: int = 480

    @classmethod
    def from_env(cls, env: Mapping[str, str]) -> "Config":
        def _int(key: str, default: int) -> int:
            raw = env.get(key)
            if raw is None:
                return default
            return int(raw)  # ValueError on invalid

        return cls(
            batch_size=_int("REPORT_REVIEWER_BATCH_SIZE", 5),
            image_max_dim=_int("REPORT_REVIEWER_IMAGE_MAX_DIM", 1928),
            pdf_dpi=_int("REPORT_REVIEWER_PDF_DPI", 150),
            thumbnail_max_dim=_int("REPORT_REVIEWER_THUMBNAIL_MAX_DIM", 480),
        )
