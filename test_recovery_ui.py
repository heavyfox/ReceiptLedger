import json
import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox
from app import MainWindow
from config import Settings
from core import ExcelLedger, Receipt
from lm_client import _parse_receipt
from session_store import SessionStore


class RecoveryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.application = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name).resolve()
        self.env = patch.dict(os.environ, {"RECEIPT_LEDGER_DATA_DIR": str(self.root)})
        self.env.start()
        self.token = patch("app.get_token", return_value="")
        self.token.start()
        Settings(workbook_path=str(self.root / "ledger.xlsx")).save()
        self.window = MainWindow()

    def tearDown(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.close()
        self.token.stop()
        self.env.stop()
        self.temp.cleanup()

    def add_ready(self):
        image = self.root / "receipt.png"
        Image.new("RGB", (1200, 2400), "white").save(image)
        self.window._enqueue_images([str(image)])
        self.window._batch_image_succeeded(str(image), _parse_receipt({"receipt_date": date.today().isoformat(),
            "merchant": "店", "total_yen": 1080, "tax_yen": 80,
            "items": [{"name": "商品", "quantity": 2, "unit_price_yen": 500, "line_total_yen": 1000}]}))
        return str(image)

    def restart(self, crash=False):
        if crash:
            self.window._restoring_session = True
        self.window.close()
        self.window = MainWindow()

    def test_timer_saves_corrections_checks_and_view_then_restores(self):
        path = self.add_ready()
        self.window.memo_edit.setText("再起動後に残すメモ")
        self.window.item_table.item(0, 4).setText("修正した補足")
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        self.window.preview.rotate(90)
        self.window.preview.actual_size()
        self.window.details_toggle.setChecked(True)
        QTest.qWait(1200)
        data = json.loads(self.window.session_store.path.read_text(encoding="utf-8"))
        self.assertEqual(data["entries"][0]["draft"]["memo"], "再起動後に残すメモ")
        self.restart(crash=True)
        self.assertEqual(self.window.active_queue_path, path)
        self.assertEqual(self.window.memo_edit.text(), "再起動後に残すメモ")
        self.assertEqual(self.window.item_table.item(0, 4).text(), "修正した補足")
        self.assertTrue(self.window.entries[path].checked)
        self.assertEqual(self.window.preview.angle, 90)
        self.assertFalse(self.window.preview.fitted)
        self.assertTrue(self.window.details_toggle.isChecked())

    def test_interrupted_reading_restores_pending_without_network(self):
        path = self.add_ready()
        self.window.entries[path].state = "reading"
        self.window._save_session()
        with patch("app.LMStudioClient") as client:
            self.restart(crash=True)
            client.assert_not_called()
        self.assertEqual(self.window.entries[path].state, "pending")
        self.assertEqual(self.window.entries[path].draft.amount, "1080")
        self.assertFalse(self.window.entries[path].checked)

    def test_crash_after_excel_commit_does_not_restore_as_unrecorded(self):
        path = self.add_ready()
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        receipt = self.window.entries[path].draft.to_receipt(path)
        self.window._stage_commit([(path, receipt)])
        self.window.ledger.save([receipt])
        self.restart(crash=True)
        self.assertEqual(self.window.entries[path].state, "saved")
        self.assertFalse(self.window.entries[path].checked)
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.assertEqual(len(self.window.receipts), 1)

    def test_editing_existing_record_restores_identity_and_updates_once(self):
        receipt = Receipt(purchased_on=date.today(), merchant="元の店", amount=200)
        self.window.ledger.save([receipt])
        self.window._load_ledger()
        self.window.ledger_table.selectRow(0)
        self.window._edit_selected()
        self.window.memo_edit.setText("修正途中")
        self.restart()
        self.assertEqual(self.window.editing_id, receipt.receipt_id)
        self.assertEqual(self.window.memo_edit.text(), "修正途中")
        self.window._save_receipt()
        self.restart()
        records = self.window.ledger.load()
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].memo, "修正途中")
        self.assertEqual(self.window.memo_edit.text(), "")

    def test_missing_original_keeps_draft_and_can_record_text(self):
        path = self.add_ready()
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        self.window._save_session()
        Path(path).unlink()
        self.restart(crash=True)
        self.assertEqual(self.window.amount_edit.text(), "1080")
        self.assertTrue(self.window.preview.pixmap().isNull())
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.batch_save_btn.click()
        self.assertEqual(len(self.window.ledger.load()), 1)

    def test_atomic_write_failure_leaves_previous_recovery_file(self):
        store = SessionStore(self.root)
        before = {"version": 1, "workbook": str(self.root / "old.xlsx"), "entries": []}
        store.save(before)
        with patch("session_store.os.replace", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                store.save({**before, "workbook": "new.xlsx"})
        self.assertEqual(store.load(), before)
        self.assertFalse(list(self.root.glob(".draft-*")))

    def test_view_zoom_rotate_fit_does_not_modify_original(self):
        path = self.add_ready()
        original = Path(path).read_bytes()
        view = self.window.preview
        view.actual_size()
        view.zoom(1.25)
        self.assertAlmostEqual(view.view.transform().m11(), 1.25)
        view.rotate(90)
        self.assertEqual(view.item.pixmap().size().width(), 2400)
        view.rotate(-90)
        self.assertEqual(view.item.pixmap().size().width(), 1200)
        self.window.show()
        self.application.processEvents()
        view.fit()
        self.assertTrue(view.fitted)
        self.assertLess(view.view.transform().m11(), 1)
        self.assertEqual(Path(path).read_bytes(), original)

    def test_corrupt_session_is_preserved_before_new_autosave(self):
        self.window._restoring_session = True
        self.window.close()
        damaged = '{"version":1,"entries":'
        (self.root / "draft-session.json").write_text(damaged, encoding="utf-8")
        self.window = MainWindow()
        self.assertIn("復元できなかった", self.window.autosave_label.text())
        backups = list(self.root.glob("draft-session.invalid-*.json"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_text(encoding="utf-8"), damaged)
        self.assertTrue(self.window._save_session())

    def test_manual_form_and_discard_do_not_touch_excel_or_source(self):
        path = self.add_ready()
        self.window.ledger.save([Receipt(merchant="記録済み", amount=200)])
        self.window._new_form()
        self.window.tax_edit.setText("0")
        self.window.memo_edit.setText("手入力の途中")
        self.restart()
        self.assertEqual(self.window.memo_edit.text(), "手入力の途中")
        self.assertEqual(self.window.tax_edit.text(), "0")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._discard_session()
        self.restart()
        self.assertEqual(self.window.queue.count(), 0)
        self.assertEqual(self.window.memo_edit.text(), "")
        self.assertEqual(len(self.window.ledger.load()), 1)
        self.assertTrue(Path(path).exists())
