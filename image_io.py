"""Receipt image decoding and conversion, including HEIC/HEIF."""

from __future__ import annotations

import base64
import io
from pathlib import Path

from PIL import Image, ImageOps
from pillow_heif import register_heif_opener


SUPPORTED_EXTENSIONS = frozenset({".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"})

# Register once for Pillow's Image.open(). Disable embedded thumbnails so small
# text in long receipts is decoded from the full image before resizing.
register_heif_opener(thumbnails=False)


def supported_image(path: str | Path) -> bool:
    return Path(path).suffix.lower() in SUPPORTED_EXTENSIONS


def _rgb_image(path: str | Path, bounds: tuple[int, int]) -> Image.Image:
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source)
        image.thumbnail(bounds, Image.Resampling.LANCZOS)
        if image.mode == "RGB":
            return image.copy()
        if "A" in image.getbands():
            background = Image.new("RGB", image.size, "white")
            background.paste(image, mask=image.getchannel("A"))
            return background
        return image.convert("RGB")


def preview_png(path: str | Path, bounds=(1000, 1400)) -> bytes:
    image = _rgb_image(path, bounds)
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def model_jpeg_data_url(path: str | Path) -> str:
    image = _rgb_image(path, (2200, 4200))
    output = io.BytesIO()
    image.save(output, format="JPEG", quality=92, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(output.getvalue()).decode("ascii")
