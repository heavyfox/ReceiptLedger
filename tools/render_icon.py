"""Render our MIT-licensed vector icon to PNG and multi-resolution Windows ICO."""
import io
from pathlib import Path
from PIL import Image
from PySide6.QtCore import QByteArray, QBuffer, QIODevice, Qt
from PySide6.QtGui import QImage, QPainter
from PySide6.QtSvg import QSvgRenderer

assets = Path(__file__).resolve().parents[1] / "assets"
canvas = QImage(512, 512, QImage.Format.Format_ARGB32)
canvas.fill(Qt.GlobalColor.transparent)
painter = QPainter(canvas)
renderer = QSvgRenderer(str(assets / "receipt-ledger.svg"))
assert renderer.isValid()
renderer.render(painter)
painter.end()
assert canvas.save(str(assets / "receipt-ledger.png"))
Image.open(assets / "receipt-ledger.png").save(assets / "receipt-ledger.ico", format="ICO",
    sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
