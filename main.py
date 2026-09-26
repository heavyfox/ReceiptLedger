import os
import sys
import tempfile
import shutil
from pathlib import Path

from app import run


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        from datetime import date
        from unittest.mock import patch
        from PySide6.QtWidgets import QApplication, QMessageBox
        from app import MainWindow
        from config import Settings
        from batch import BatchReadWorker
        from lm_client import _parse_receipt
        from core import ExcelLedger, LineItem, Receipt
        from image_io import Image, model_jpeg_data_url, preview_png
        from data_backup import create_bundle, restore_bundle
        from version import APP_VERSION

        with tempfile.TemporaryDirectory() as folder, patch("app.get_token", return_value=""), patch("app.save_token"):
            os.environ["RECEIPT_LEDGER_DATA_DIR"] = folder
            Settings(workbook_path=str(Path(folder) / "test.xlsx")).save()
            app = QApplication([])
            window = MainWindow()
            window.theme_selector.setCurrentIndex(window.theme_selector.findData("dark"))
            assert Settings.load().theme == "dark"
            window.theme_selector.setCurrentIndex(window.theme_selector.findData("system"))
            assert Settings.load().theme == "system"
            window._save_workbook()
            ledger = window.ledger
            ledger.save([Receipt(purchased_on=date.today(), merchant="テスト店", amount=123,
                                 category="食費", tax=11, items=[LineItem("品物", 1, 123, 123, "食費", "補足")])])
            assert ledger.load()[0].amount == 123
            assert ledger.load()[0].tax == 11
            assert ledger.load()[0].items[0].note == "補足"
            heic = (Path(folder) / "test.heic").resolve()
            shutil.copyfile(Path(__file__).resolve().parent / "tests/fixtures/receipt.heic", heic)
            assert preview_png(heic).startswith(b"\x89PNG")
            assert model_jpeg_data_url(heic).startswith("data:image/jpeg;base64,")
            for extension, image_format in (("jpg", "JPEG"), ("jpeg", "JPEG"), ("png", "PNG"), ("webp", "WEBP")):
                sample = Path(folder) / ("sample." + extension)
                Image.new("RGB", (320, 640), "white").save(sample, format=image_format)
                assert preview_png(sample).startswith(b"\x89PNG")
                assert model_jpeg_data_url(sample).startswith("data:image/jpeg;base64,")
            heif = Path(folder) / "sample.heif"
            shutil.copyfile(heic, heif)
            assert preview_png(heif).startswith(b"\x89PNG")
            assert not window.windowIcon().isNull()
            class TestClient:
                def extract(self, image_path, model):
                    return _parse_receipt({"receipt_date": date.today().isoformat(),
                                           "merchant": "一括テスト", "total_yen": 123})

            batch_results = []
            window.concurrency_box.setCurrentIndex(window.concurrency_box.findData(2))
            assert Settings.load().concurrent_reads == 2
            reader = BatchReadWorker([str(heic), str(heic)], TestClient(), "test", window, concurrency=2)
            reader.image_succeeded.connect(lambda path, result: batch_results.append(result))
            reader.start()
            assert reader.wait(10000)
            app.processEvents()
            assert len(batch_results) == 2
            backup = ledger.delete()
            assert not ledger.path.exists()
            assert ExcelLedger(backup).load()[0].items[0].note == "補足"
            window.merchant_edit.setText("復元確認の店")
            window.memo_edit.setText("編集中の内容")
            window._set_image(str(heic))
            window.preview.rotate(90)
            window.preview.actual_size()
            window.close()
            recovered = MainWindow()
            assert recovered.merchant_edit.text() == "復元確認の店"
            assert recovered.memo_edit.text() == "編集中の内容"
            assert recovered.preview.angle == 90
            assert recovered.preview.view.transform().m11() == 1
            recovered._enqueue_images([str(heic)])
            recovered._batch_image_succeeded(str(heic), TestClient().extract(str(heic), "test"))
            recovered.queue_filter.setCurrentIndex(recovered.queue_filter.findData("ready"))
            recovered.queue_search.setText("一括テスト")
            recovered.check_visible_btn.click()
            assert recovered.entries[str(heic)].checked
            recovered.queue_search.setText("該当なし")
            assert recovered.queue.item(0).isHidden()
            assert "非表示 1 件" in recovered.review_count_label.text()
            recovered.close()
            filtered = MainWindow()
            assert filtered.queue_filter.currentData() == "all"
            assert filtered.queue_search.text() == ""
            assert not filtered.queue.item(0).isHidden()
            assert filtered.entries[str(heic)].checked
            filtered.clear_checks_btn.click()
            assert not filtered.entries[str(heic)].checked
            filtered.queue_search.clear()
            filtered.check_visible_btn.click()
            with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
                filtered.batch_save_btn.click()
            assert len(filtered.ledger.load()) == 1
            filtered.queue_filter.setCurrentIndex(filtered.queue_filter.findData("saved"))
            filtered.check_visible_btn.click()
            assert filtered.entries[str(heic)].checked
            assert not filtered.batch_save_btn.isEnabled()
            with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
                filtered.remove_image_actions["checked"].trigger()
            assert not filtered.entries
            assert heic.exists()
            assert len(filtered.ledger.load()) == 1
            filtered.close()
            removed = MainWindow()
            assert not removed.entries
            assert removed.queue.count() == 0
            assert removed.undo_remove_btn.isEnabled()
            removed.undo_remove_btn.click()
            assert removed.entries[str(heic)].state == "saved"
            assert removed.entries[str(heic)].checked
            assert "未記録 0 件・記録済み 1 件" in removed.review_count_label.text()
            assert removed.remove_checked_btn.text() == "選択 1 件を一覧から削除"
            assert not removed.batch_save_btn.isEnabled()
            assert len(removed.ledger.load()) == 1
            assert APP_VERSION in removed.windowTitle()
            stored = removed.entries[str(heic)].stored_image
            assert Path(stored).is_file()
            heic.unlink()
            removed.queue.setCurrentRow(0)
            removed._set_image(str(heic))
            assert removed.current_image == stored
            bundle = Path(folder) / "all-data.zip"
            create_bundle(bundle, removed.settings, removed._session_snapshot())
            restored = restore_bundle(bundle, Path(folder) / "restored")
            restored_records = ExcelLedger(restored["settings"]["workbook_path"]).load()
            assert len(restored_records) == 1
            assert Path(restored_records[0].image_path).is_file()
            removed.close()
        raise SystemExit(0)
    raise SystemExit(run())
