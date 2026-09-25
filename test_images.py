import os
import tempfile
import unittest
import shutil
from pathlib import Path

from PIL import Image
from PySide6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication

from app import MainWindow
from config import Settings
from image_io import model_jpeg_data_url, preview_png


class ImageInputTests(unittest.TestCase):
    def test_jpeg_orientation_and_supported_formats(self):
        with tempfile.TemporaryDirectory() as folder:
            for suffix, fmt in (("jpg", "JPEG"), ("jpeg", "JPEG"), ("png", "PNG"), ("webp", "WEBP")):
                with self.subTest(format=fmt, extension=suffix):
                    path = Path(folder) / ("image." + suffix)
                    image = Image.new("RGB", (80, 120), "white")
                    exif = Image.Exif()
                    exif[274] = 6
                    image.save(path, format=fmt, exif=exif)
                    with Image.open(__import__("io").BytesIO(preview_png(path))) as preview:
                        self.assertEqual(preview.size, (120, 80))
                    self.assertTrue(model_jpeg_data_url(path).startswith("data:image/jpeg;base64,"))

    def test_heic_decodes_for_preview_and_model(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "receipt.HEIC"
            shutil.copyfile(Path(__file__).resolve().parent / "tests/fixtures/receipt.heic", path)
            with Image.open(__import__("io").BytesIO(preview_png(path))) as preview:
                self.assertEqual(preview.format, "PNG")
                self.assertEqual(preview.size, (420, 900))
            self.assertTrue(model_jpeg_data_url(path).startswith("data:image/jpeg;base64,"))

    def test_drop_on_form_field_adds_heic_to_queue(self):
        with tempfile.TemporaryDirectory() as folder:
            os.environ["RECEIPT_LEDGER_DATA_DIR"] = folder
            Settings(workbook_path=str(Path(folder) / "ledger.xlsx")).save()
            path = Path(folder) / "receipt.heic"
            shutil.copyfile(Path(__file__).resolve().parent / "tests/fixtures/receipt.heic", path)
            app = QApplication.instance() or QApplication([])
            window = MainWindow()
            window.show()
            app.processEvents()
            try:
                mime = QMimeData()
                mime.setUrls([QUrl.fromLocalFile(str(path))])
                enter = QDragEnterEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                                       Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
                QApplication.sendEvent(window.merchant_edit, enter)
                self.assertTrue(enter.isAccepted())
                drop = QDropEvent(QPointF(10, 10), Qt.DropAction.CopyAction, mime,
                                  Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
                QApplication.sendEvent(window.merchant_edit, drop)
                self.assertEqual(window.queue.count(), 1)
                self.assertEqual(Path(window.current_image), path.resolve())
                self.assertFalse(window.preview.pixmap().isNull())
            finally:
                window.close()
                os.environ.pop("RECEIPT_LEDGER_DATA_DIR", None)


if __name__ == "__main__":
    unittest.main()
