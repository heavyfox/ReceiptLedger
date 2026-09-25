import json
import os
import shutil
import tempfile
import unittest
import zipfile
from dataclasses import asdict
from datetime import date
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox
from app import MainWindow
from config import Settings, app_data_dir, default_data_dir
from core import ExcelLedger, Receipt
from data_backup import create_bundle, restore_bundle, activate_restored_state, recover_activation
from lm_client import _parse_receipt
from session_store import SessionStore
from storage import copy_image, copy_workbook_rebased, regular_backups, prune_backups, file_usage
from version import APP_VERSION


class StorageBundleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.application = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name).resolve()
        self.env = patch.dict(os.environ, {"RECEIPT_LEDGER_DATA_DIR": str(self.root / "state")})
        self.env.start()
        self.token = patch("app.get_token", return_value="test-secret-not-for-export")
        self.token.start()
        self.settings = Settings(workbook_path=str(self.root / "excel" / "家計簿.xlsx"),
                                 image_folder=str(self.root / "images"))
        self.settings.save()
        self.window = MainWindow()

    def tearDown(self):
        self.window.close()
        self.token.stop()
        self.env.stop()
        self.temp.cleanup()

    def ready(self, name="photo.png"):
        path = self.root / name
        Image.new("RGB", (80, 160), "white").save(path)
        self.window._enqueue_images([str(path)])
        self.window._batch_image_succeeded(str(path), _parse_receipt({
            "receipt_date": date.today().isoformat(), "merchant": "テスト店", "total_yen": 108,
            "tax_yen": 8, "items": [{"name": "商品", "quantity": 1, "unit_price_yen": 100, "line_total_yen": 100}]}))
        return path

    def record(self, path):
        self.window.queue.setCurrentItem(self.window._queue_item(str(path)))
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._save_receipt()
        return self.window.receipts[-1]

    def wait_worker(self):
        for _ in range(400):
            self.application.processEvents()
            if not self.window.workers:
                return
            QTest.qWait(10)
        self.fail("Worker did not finish")

    def test_default_locations_and_legacy_image_folder(self):
        self.assertEqual(default_data_dir(), self.root / "state")
        state = app_data_dir()
        (state / "settings.json").write_text(json.dumps({"workbook_path": str(self.root / "old.xlsx")}), encoding="utf-8")
        loaded = Settings.load()
        self.assertEqual(loaded.workbook_path, str(self.root / "old.xlsx"))
        self.assertEqual(loaded.image_folder, str(state / "images"))
        with patch.dict(os.environ, {}, clear=True), patch("config.application_dir", return_value=self.root / "installed"):
            self.assertEqual(default_data_dir(), self.root / "installed" / "data")

    def test_portable_settings_session_and_excel_survive_directory_move(self):
        original, moved = self.root / "app-a", self.root / "app-b"
        with patch("config.application_dir", return_value=original):
            source = self.ready()
            image, _ = copy_image(source, original / "data" / "images")
            workbook = original / "data" / "excel" / "ledger.xlsx"
            ExcelLedger(workbook).save([Receipt(merchant="portable", amount=100, image_path=str(image))])
            Settings(workbook_path=str(workbook), image_folder=str(image.parent)).save()
            SessionStore(app_data_dir()).save({"version": 1, "workbook": str(workbook), "entries": [], "image": str(image)})
        shutil.copytree(original, moved)
        with patch("config.application_dir", return_value=moved):
            settings = Settings.load()
            self.assertEqual(Path(settings.workbook_path), moved / "data" / "excel" / "ledger.xlsx")
            self.assertTrue(Path(ExcelLedger(settings.workbook_path).load()[0].image_path).exists())
            session = SessionStore(app_data_dir()).load()
            self.assertTrue(session["image"].startswith(str(moved)))

    def test_missing_original_uses_saved_copy_and_restores_after_restart(self):
        path = self.ready()
        receipt = self.record(path)
        saved = Path(receipt.image_path)
        self.assertEqual(saved.parent, Path(self.settings.image_folder))
        path.unlink()
        self.window.queue.setCurrentItem(self.window._queue_item(str(path)))
        self.assertEqual(self.window.current_image, str(saved))
        self.assertFalse(self.window.preview.pixmap().isNull())
        self.window.close()
        self.window = MainWindow()
        self.assertEqual(self.window.current_image, str(saved))

    def test_legacy_queue_without_saved_path_recovers_by_receipt_details(self):
        path = self.ready()
        receipt = self.record(path)
        entry = self.window.entries[str(path)]
        entry.stored_image = entry.image_hash = ""
        path.unlink()
        self.window.queue.setCurrentItem(self.window._queue_item(str(path)))
        self.assertEqual(self.window.current_image, receipt.image_path)

    def test_corrupt_original_uses_saved_copy(self):
        path = self.ready()
        receipt = self.record(path)
        path.write_bytes(b"broken image")
        self.window.queue.setCurrentItem(self.window._queue_item(str(path)))
        self.assertEqual(self.window.current_image, receipt.image_path)

    def test_setting_new_folder_affects_future_copies_and_keeps_existing(self):
        first = self.ready()
        old_image = Path(self.record(first).image_path)
        new_folder = self.root / "other-images"
        self.window.image_folder_edit.setText(str(new_folder))
        self.window.backup_limit_spin.setValue(3)
        with patch("app.save_token"), patch("app.QMessageBox.warning") as warning:
            self.window._save_settings()
        warning.assert_not_called()
        self.assertEqual(Settings.load().image_folder, str(new_folder))
        self.assertEqual(self.window.ledger.backup_limit, 3)
        second = self.ready("second.png")
        self.assertEqual(Path(self.record(second).image_path).parent, new_folder)
        self.assertTrue(old_image.exists())

    def test_invalid_setting_does_not_change_active_storage(self):
        self.window.image_folder_edit.setText(str(self.root / "unwritable"))
        with patch("app.writable_directory", side_effect=PermissionError("access denied")), patch("app.QMessageBox.warning") as warning:
            self.window._save_settings()
        warning.assert_called_once()
        self.assertEqual(self.window.settings.image_folder, self.settings.image_folder)
        self.assertEqual(Settings.load().image_folder, self.settings.image_folder)

    def test_retention_only_removes_exact_regular_backups(self):
        ledger = self.window.ledger
        ledger.backup_limit = 2
        ledger.save([Receipt(merchant="店", amount=100)])
        safety = ledger.path.with_name(ledger.path.stem + ".before-delete-20260925-000000-000000.xlsx")
        shutil.copy2(ledger.path, safety)
        unrelated = ledger.path.with_name("unrelated.backup-20260925-000000-000000.xlsx")
        shutil.copy2(ledger.path, unrelated)
        for amount in range(101, 106):
            ledger.save([Receipt(merchant="店", amount=amount)])
        self.assertEqual(len(regular_backups(ledger.path)), 2)
        self.assertTrue(safety.exists())
        self.assertTrue(unrelated.exists())
        self.assertEqual(ledger.load()[0].amount, 105)
        self.assertEqual(file_usage([safety, safety])[0], 1)

    def test_retention_failure_does_not_report_excel_save_failure(self):
        ledger = self.window.ledger
        ledger.save([Receipt(merchant="店", amount=100)])
        with patch("storage.prune_backups", side_effect=PermissionError("locked")):
            ledger.save([Receipt(merchant="店", amount=200)])
        self.assertIn("Excelは保存済み", ledger.maintenance_warning)
        self.assertEqual(ledger.load()[0].amount, 200)

    def test_bundle_restores_images_drafts_undo_and_workbook_without_overwrite(self):
        path = self.ready()
        receipt = self.record(path)
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._remove_images("checked")
        draft = self.ready("draft.png")
        self.window.memo_edit.setText("未記録の編集")
        snapshot = self.window._session_snapshot()
        bundle = self.root / "all.zip"
        original_excel = self.window.ledger.path.read_bytes()
        result = create_bundle(bundle, self.window.settings, snapshot)
        self.assertFalse(result["missing"])
        with zipfile.ZipFile(bundle) as archive:
            manifest_text = archive.read("manifest.json").decode()
            self.assertNotIn("test-secret-not-for-export", manifest_text)
        restored = restore_bundle(bundle, self.root / "restores")
        self.assertEqual(self.window.ledger.path.read_bytes(), original_excel)
        new_ledger = ExcelLedger(restored["settings"]["workbook_path"])
        self.assertEqual(new_ledger.load()[0].receipt_id, receipt.receipt_id)
        self.assertTrue(Path(new_ledger.load()[0].image_path).exists())
        self.assertEqual(restored["session"]["entries"][0]["draft"]["memo"], "未記録の編集")
        self.assertEqual(len(restored["session"]["removed_batches"]), 1)
        # Identical images imported from distinct paths remain distinct entries/history.
        self.assertNotEqual(restored["session"]["entries"][0]["path"],
                            restored["session"]["removed_batches"][0]["entries"][0]["entry"]["path"])
        activate_restored_state(restored["settings"], restored["session"])
        self.window._restoring_session = True
        self.window.close()
        self.window = MainWindow()
        self.assertEqual(self.window.memo_edit.text(), "未記録の編集")
        self.window.undo_remove_btn.click()
        self.assertEqual(len(self.window.entries), 2)
        self.assertEqual(len(self.window.ledger.load()), 1)

    def test_bundle_with_missing_original_keeps_saved_preview(self):
        path = self.ready()
        self.record(path)
        path.unlink()
        bundle = self.root / "missing.zip"
        result = create_bundle(bundle, self.window.settings, self.window._session_snapshot())
        self.assertIn(str(path), result["missing"])
        restored = restore_bundle(bundle, self.root / "restores")
        entry = restored["session"]["entries"][0]
        self.assertTrue(Path(entry["stored_image"]).exists())

    def test_corrupt_or_traversing_bundle_does_not_publish_restore_folder(self):
        self.record(self.ready())
        bundle = self.root / "good.zip"
        create_bundle(bundle, self.window.settings, self.window._session_snapshot())
        for mode in ("checksum", "traversal", "extra"):
            bad = self.root / (mode + ".zip")
            with zipfile.ZipFile(bundle) as original, zipfile.ZipFile(bad, "w") as output:
                manifest = json.loads(original.read("manifest.json"))
                if mode == "checksum":
                    manifest["files"][0]["sha256"] = "0" * 64
                elif mode == "traversal":
                    manifest["files"][0]["sha256"] = "../../outside"
                else:
                    output.writestr("../../outside", "malicious")
                for name in original.namelist():
                    output.writestr(name, json.dumps(manifest) if name == "manifest.json" else original.read(name))
            with self.assertRaises(ValueError):
                restore_bundle(bad, self.root / "restores")
            self.assertFalse(list((self.root / "restores").glob("ReceiptLedger-restored-*")))

    def test_activation_journal_completes_after_interrupted_session_write(self):
        self.record(self.ready())
        settings = asdict(self.window.settings)
        session = self.window._session_snapshot()
        settings["model"] = "restored-model"
        with patch("data_backup.SessionStore.save", side_effect=OSError("interrupted")):
            with self.assertRaises(OSError):
                activate_restored_state(settings, session)
        self.assertTrue((app_data_dir() / "restore-pending.json").exists())
        recover_activation()
        self.assertEqual(Settings.load().model, "restored-model")
        self.assertFalse((app_data_dir() / "restore-pending.json").exists())
        self.assertEqual(SessionStore(app_data_dir()).load()["entries"], session["entries"])

    def test_ui_backup_restore_and_capacity_controls(self):
        self.record(self.ready())
        bundle = self.root / "ui.zip"
        with patch("app.QFileDialog.getSaveFileName", return_value=(str(bundle), "ZIP")), \
                patch("app.QMessageBox.information"), patch("app.QMessageBox.warning") as warning:
            self.window._export_bundle()
            self.wait_worker()
        warning.assert_not_called()
        self.assertTrue(bundle.exists())
        with patch("app.QFileDialog.getOpenFileName", return_value=(str(bundle), "ZIP")), \
                patch("app.QFileDialog.getExistingDirectory", return_value=str(self.root / "restores")), \
                patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes), \
                patch("app.QMessageBox.information"), patch("app.QMessageBox.warning") as warning:
            self.window._import_bundle()
            self.wait_worker()
        warning.assert_not_called()
        self.assertIn("ReceiptLedger-restored-", self.window.settings.workbook_path)
        self.assertEqual(len(self.window.receipts), 1)
        self.assertTrue(list((app_data_dir() / "backups").glob("before-data-restore-*.zip")))
        self.window.tabs.setCurrentIndex(2)
        self.window._refresh_storage_usage()
        self.wait_worker()
        self.assertEqual(self.window.tabs.currentIndex(), 2)
        self.assertIn("保存済み画像", self.window.storage_usage_label.text())
        self.assertIn(f"v{APP_VERSION}", self.window.windowTitle())

    def test_excel_restore_to_different_folder_rebases_image_reference(self):
        receipt = self.record(self.ready())
        target = self.root / "other-folder" / "restored.xlsx"
        copy_workbook_rebased(self.window.ledger.path, target)
        self.assertEqual(ExcelLedger(target).load()[0].image_path, receipt.image_path)

    def test_reread_uses_saved_copy_without_adding_second_queue_entry(self):
        path = self.ready()
        receipt = self.record(path)
        path.unlink()
        self.window.queue.setCurrentRow(0)
        self.window.model_box.setCurrentText("test")
        result = _parse_receipt({"receipt_date": date.today().isoformat(), "merchant": "再読取", "total_yen": 108})
        with patch("app.LMStudioClient") as client, patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            client.return_value.extract.return_value = result
            self.window._read_image()
            self.wait_worker()
            self.assertEqual(client.return_value.extract.call_args.args[0], receipt.image_path)
        self.assertEqual(self.window.queue.count(), 1)
        self.assertEqual(self.window.entries[str(path)].draft.merchant, "再読取")


if __name__ == "__main__":
    unittest.main()
