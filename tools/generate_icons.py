"""Rebuild PWA icons with Pillow (included with desktop matplotlib)."""
from pathlib import Path

from PIL import Image, ImageDraw

root = Path(__file__).resolve().parents[1] / "web" / "static"
image = Image.new("RGB", (512, 512), "#087f74")
ImageDraw.Draw(image).polygon(
    [(116, 364), (116, 148), (172, 148), (256, 260), (340, 148), (396, 148),
     (396, 364), (340, 364), (340, 238), (256, 344), (172, 238), (172, 364)],
    fill="white",
)
for size, name in ((180, "apple-touch-icon.png"), (192, "icon-192.png"), (512, "icon-512.png")):
    image.resize((size, size), Image.Resampling.LANCZOS).save(root / name)
