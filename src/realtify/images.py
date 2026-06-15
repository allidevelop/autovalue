from __future__ import annotations

from pathlib import Path

from PIL import Image


def create_report_image(
    source_path: Path,
    output_path: Path,
    *,
    max_width_px: int = 1600,
    jpeg_quality: int = 82,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source_path) as image:
        image = image.convert("RGB")
        if image.width > max_width_px:
            ratio = max_width_px / image.width
            new_height = max(1, int(image.height * ratio))
            image = image.resize((max_width_px, new_height), Image.Resampling.LANCZOS)
        image.save(output_path, "JPEG", quality=jpeg_quality, optimize=True)
    return output_path

