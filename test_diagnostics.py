import json
import os
from pathlib import Path
import socket
import tempfile
import unittest
import urllib.error
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import diagnostics
from lm_client import LMStudioClient, LMStudioError


class DiagnosticsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.env = patch.dict(os.environ, RECEIPT_LEDGER_DATA_DIR=self.temp.name)
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.temp.cleanup()

    def rows(self):
        return [json.loads(line) for line in (diagnostics.folder() / "diagnostics.jsonl").read_text().splitlines()]

    def test_only_safe_metadata_reaches_disk_or_clipboard_report(self):
        secret = "PRIVATE-token-name-receipt-text"
        diagnostics.record("extract_finished", format=".heic", elapsed_ms=125, error="timeout",
                           token=secret, path=secret, body=secret, job=secret, endpoint=secret)
        diagnostics.record(secret, error="timeout")
        diagnostics.record("task_failed", error=[secret], format={"private": secret})
        rows = self.rows()
        self.assertEqual(rows[0]["format"], ".heic")
        self.assertEqual(rows[0]["elapsed_ms"], 125)
        self.assertEqual(rows[0]["error"], "timeout")
        self.assertNotIn(secret, json.dumps(rows))
        # An edited log must also be sanitized before copying diagnostic text.
        with (diagnostics.folder() / "diagnostics.jsonl").open("a") as stream:
            stream.write(json.dumps({"event": "task_failed", "time": secret, "token": secret}) + "\n")
        self.assertNotIn(secret, diagnostics.report())

    def test_unwritable_logging_does_not_break_processing(self):
        with patch("diagnostics.folder", side_effect=PermissionError("private path")):
            diagnostics.record("task_failed", error="permission")
            self.assertIn("まだありません", diagnostics.report())

    def test_rotation_is_bounded_and_thread_safe(self):
        with patch("diagnostics.MAX_BYTES", 1024):
            with ThreadPoolExecutor(max_workers=2) as pool:
                list(pool.map(lambda n: diagnostics.record("extract_finished", elapsed_ms=n, error="none"), range(100)))
        paths = list(diagnostics.folder().glob("diagnostics.jsonl*"))
        self.assertEqual(len(paths), 4)
        for path in paths:
            self.assertLessEqual(path.stat().st_size, 1024)
            for line in path.read_text().splitlines():
                self.assertEqual(json.loads(line)["event"], "extract_finished")

    def test_retry_and_truncated_output_have_diagnostic_codes(self):
        client = LMStudioClient("http://localhost:1234/v1", token="secret-token")
        response = {"choices": [{"finish_reason": "length", "message": {"content": "secret-receipt"}}]}
        with patch("lm_client._image_data_url", return_value="secret-image"), patch.object(client, "_request", return_value=response):
            with self.assertRaises(LMStudioError) as error:
                client.extract(Path(self.temp.name) / "secret-file.heic", "secret-model")
        self.assertEqual(error.exception.code, "output_limit")
        rows = self.rows()
        self.assertEqual([row["event"] for row in rows], ["extract_started", "extract_retry", "extract_finished"])
        self.assertEqual(rows[-1]["error"], "output_limit")
        self.assertEqual(rows[0]["job"], rows[-1]["job"])
        self.assertNotIn("secret", json.dumps(rows))

    def test_request_timeout_classified_without_url_or_exception_text(self):
        client = LMStudioClient("http://localhost:1234/v1")
        with patch("lm_client.urllib.request.build_opener") as factory:
            factory.return_value.open.side_effect = urllib.error.URLError(socket.timeout("secret detail"))
            with self.assertRaises(LMStudioError) as error:
                client.models()
        self.assertEqual(error.exception.code, "timeout")
        self.assertEqual(self.rows()[-1]["endpoint"], "models")
        self.assertEqual(self.rows()[-1]["error"], "timeout")
        self.assertNotIn("secret detail", diagnostics.report())

    def test_copy_button_after_view_refactor(self):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        from app import MainWindow
        from config import Settings
        qt = QApplication.instance() or QApplication([])
        Settings(model="private-model", workbook_path=str(Path(self.temp.name) / "test.xlsx")).save()
        with patch("app.get_token", return_value="private-token"), patch("app.save_token"):
            window = MainWindow()
            try:
                diagnostics.record("extract_finished", format=".jpg", elapsed_ms=42, error="none")
                window.copy_diagnostics_btn.click()
                copied = qt.clipboard().text()
                self.assertIn('"elapsed_ms": 42', copied)
                self.assertNotIn("private-token", copied)
                self.assertNotIn("private-model", copied)
                self.assertEqual(window.tabs.count(), 3)
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()
