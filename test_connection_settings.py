import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QElapsedTimer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from app import MainWindow
from config import Settings
from lm_client import LMStudioError, _parse_receipt
from PIL import Image


class ConnectionSettingsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name).resolve()
        self.env = patch.dict(os.environ, RECEIPT_LEDGER_DATA_DIR=str(self.root))
        self.env.start()
        self.credential = "old-test-token"
        self.patches = [patch("app.get_token", side_effect=lambda: self.credential),
                        patch("app.save_token", side_effect=self.set_credential),
                        patch("app.QMessageBox.information"), patch("app.QMessageBox.warning")]
        self.mocks = [p.start() for p in self.patches]
        Settings(model="previous-model", workbook_path=str(self.root / "ledger.xlsx"),
                 image_folder=str(self.root / "images")).save()
        self.window = MainWindow()

    def set_credential(self, token):
        self.credential = token

    def wait(self):
        timer = QElapsedTimer()
        timer.start()
        while self.window.workers and timer.elapsed() < 3000:
            self.qt.processEvents()
            QTest.qWait(5)
        self.assertFalse(self.window.workers)

    def tearDown(self):
        self.wait()
        self.window.close()
        for p in reversed(self.patches):
            p.stop()
        self.env.stop()
        self.temp.cleanup()

    def reopen(self):
        self.window.close()
        self.window = MainWindow()

    def test_connection_success_restores_url_token_models_without_save_button(self):
        self.window.url_edit.setText("http://127.0.0.1:2345/")
        self.window.token_edit.setText("new-test-token")
        self.window.workbook_edit.setText("unfinished-storage-edit")
        with patch("app.LMStudioClient.models", return_value=["qwen-test", "second-model"]):
            self.window._test_connection()
            self.wait()
        self.reopen()
        self.assertEqual(self.window.url_edit.text(), "http://127.0.0.1:2345/v1")
        self.assertEqual(self.window.token_edit.text(), "new-test-token")
        self.assertEqual(self.window.model_box.currentText(), "qwen-test")
        self.assertEqual([self.window.model_box.itemText(i) for i in range(2)], ["qwen-test", "second-model"])
        self.assertEqual(self.window.ledger.path, self.root / "ledger.xlsx")
        self.assertNotIn("new-test-token", (self.root / "settings.json").read_text())

    def test_selected_and_manually_entered_models_are_restored(self):
        self.window._show_models(["previous-model", "second-model"])
        self.window.model_box.setCurrentIndex(1)
        self.window.model_box.activated.emit(1)
        self.reopen()
        self.assertEqual(self.window.model_box.currentText(), "second-model")
        self.window.model_box.setCurrentText("manual-model")
        self.window.model_box.lineEdit().editingFinished.emit()
        self.reopen()
        self.assertEqual(self.window.model_box.currentText(), "manual-model")
        self.assertEqual(self.window.model_box.count(), 2)

    def test_failed_connection_keeps_previous_successful_settings(self):
        self.window._show_models(["previous-model"])
        before = (self.root / "settings.json").read_bytes()
        self.window.url_edit.setText("http://127.0.0.1:9999/v1")
        self.window.token_edit.setText("failed-test-token")
        with patch("app.LMStudioClient.models", side_effect=LMStudioError("test connection failed")):
            self.window._test_connection()
            self.wait()
        self.assertEqual((self.root / "settings.json").read_bytes(), before)
        self.assertEqual(self.credential, "old-test-token")
        self.reopen()
        self.assertEqual(self.window.url_edit.text(), "http://localhost:1234/v1")
        self.assertEqual(self.window.model_box.currentText(), "previous-model")

    def test_write_failure_is_reported_and_rolls_back_credential(self):
        before = (self.root / "settings.json").read_bytes()
        with patch.object(Settings, "save", side_effect=OSError("test disk full")):
            self.window._show_models(["new-model"], url="http://127.0.0.1:2345/v1", token="new-test-token")
        self.assertEqual(self.credential, "old-test-token")
        self.assertEqual((self.root / "settings.json").read_bytes(), before)
        self.mocks[3].assert_called_once()
        self.mocks[2].assert_not_called()

    def test_successful_receipt_read_remembers_connection_without_model_fetch(self):
        sample = self.root / "sample.jpg"
        Image.new("RGB", (100, 200), "white").save(sample)
        self.window._enqueue_images([str(sample)])
        self.window.url_edit.setText("http://127.0.0.1:2345/v1")
        self.window.model_box.setCurrentText("read-model")
        self.window.token_edit.setText("read-test-token")
        with patch("app.LMStudioClient.extract", return_value=_parse_receipt({"merchant":"test", "total_yen":100})):
            self.window._read_batch()
            self.wait()
        self.reopen()
        self.assertEqual(self.window.url_edit.text(), "http://127.0.0.1:2345/v1")
        self.assertEqual(self.window.model_box.currentText(), "read-model")
        self.assertEqual(self.window.token_edit.text(), "read-test-token")

    def test_old_settings_load_and_restore_selected_model(self):
        path = self.root / "settings.json"
        old = json.loads(path.read_text())
        old.pop("model_choices", None)
        path.write_text(json.dumps(old))
        self.reopen()
        self.assertEqual(self.window.model_box.currentText(), "previous-model")
        self.assertEqual(self.window.settings.model_choices, [])

    def test_manual_settings_save_failure_preserves_previous_credential_and_settings(self):
        before = (self.root / "settings.json").read_bytes()
        self.window.url_edit.setText("http://127.0.0.1:2345/v1")
        self.window.token_edit.setText("new-test-token")
        with patch.object(Settings, "save", side_effect=OSError("test disk full")):
            self.window._save_settings()
        self.assertEqual(self.credential, "old-test-token")
        self.assertEqual((self.root / "settings.json").read_bytes(), before)
        self.mocks[3].assert_called_once()
        self.reopen()
        self.assertEqual(self.window.url_edit.text(), "http://localhost:1234/v1")
        self.assertEqual(self.window.token_edit.text(), "old-test-token")

    def test_manual_settings_save_success_restores_all_connection_fields(self):
        self.window.url_edit.setText("http://127.0.0.1:2345/v1")
        self.window.token_edit.setText("new-test-token")
        self.window.model_box.setCurrentText("manual-model")
        self.window._save_settings()
        self.mocks[3].assert_not_called()
        self.reopen()
        self.assertEqual(self.window.url_edit.text(), "http://127.0.0.1:2345/v1")
        self.assertEqual(self.window.token_edit.text(), "new-test-token")
        self.assertEqual(self.window.model_box.currentText(), "manual-model")

    def test_invalid_ipv6_url_reports_warning_without_starting_worker(self):
        self.window.url_edit.setText("http://[::1:1234/v1")
        self.window._test_connection()
        self.assertFalse(self.window.workers)
        self.mocks[3].assert_called_once()

    def test_keyring_failure_does_not_save_new_connection(self):
        before = (self.root / "settings.json").read_bytes()
        with patch("app.save_token", side_effect=RuntimeError("test keyring failure")):
            self.window._show_models(["new-model"], url="http://127.0.0.1:2345/v1", token="new-test-token")
        self.assertEqual((self.root / "settings.json").read_bytes(), before)
        self.assertEqual(self.window.settings.server_url, "http://localhost:1234/v1")
        self.mocks[3].assert_called_once()


if __name__ == "__main__":
    unittest.main()
