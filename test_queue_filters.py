import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMessageBox
from app import MainWindow
from batch import ReceiptDraft
from config import Settings
from lm_client import _parse_receipt


class QueueFilterTests(unittest.TestCase):
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
        self.paths = []
        for name in ("Apple.PNG", "shop.png", "pending.png", "failed.png", "saved.png"):
            path = self.root / name
            Image.new("RGB", (80, 160), "white").save(path)
            self.paths.append(str(path))
        self.window._enqueue_images(self.paths)
        self.window._new_form()
        for path, state, merchant in zip(self.paths,
                ("ready", "ready", "pending", "failed", "saved"),
                ("青空スーパー", "中央薬局", "", "", "青空スーパー")):
            entry = self.window.entries[path]
            entry.state = state
            if state in ("ready", "saved"):
                entry.draft = ReceiptDraft(purchased_on=date.today().isoformat(), merchant=merchant, amount="108")
            self.window._refresh_queue_item(path)

    def tearDown(self):
        self.window.close()
        self.token.stop()
        self.env.stop()
        self.temp.cleanup()

    def filter(self, state):
        self.window.queue_filter.setCurrentIndex(self.window.queue_filter.findData(state))

    def visible(self):
        return [self.window.queue.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.window.queue.count()) if not self.window.queue.item(i).isHidden()]

    def test_state_and_merchant_or_filename_search_intersect(self):
        for state, indices in (("all", range(5)), ("pending", [2]), ("ready", [0, 1]),
                               ("failed", [3]), ("saved", [4])):
            self.filter(state)
            self.assertEqual(self.visible(), [self.paths[i] for i in indices])
        self.filter("ready")
        self.window.queue_search.setText("青空")
        self.assertEqual(self.visible(), [self.paths[0]])
        self.window.queue_search.setText(" apple.png ")
        self.assertEqual(self.visible(), [self.paths[0]])
        self.window.queue_search.setText("存在しない店")
        self.assertEqual(self.visible(), [])
        self.assertIn("該当する画像はありません", self.window.queue_filter_notice.text())
        self.assertFalse(self.window.check_visible_btn.isEnabled())
        self.assertEqual(len(self.window.entries), 5)

    def test_select_visible_eligible_only_and_clear_including_hidden(self):
        self.filter("ready")
        self.window.queue_search.setText("青空")
        self.window.check_visible_btn.click()
        self.assertEqual([p for p, e in self.window.entries.items() if e.checked], [self.paths[0]])
        self.window.queue_search.setText("薬局")
        self.window.check_visible_btn.click()
        self.assertTrue(self.window.entries[self.paths[0]].checked)
        self.assertIn("非表示 1 件", self.window.review_count_label.text())
        self.assertIn("2 件", self.window.batch_save_btn.text())
        self.window.clear_checks_btn.click()
        self.assertFalse(any(e.checked for e in self.window.entries.values()))
        self.window.queue_search.clear()
        self.filter("all")
        self.window.check_visible_btn.click()
        self.assertEqual([p for p, e in self.window.entries.items() if e.checked], self.paths)
        self.assertFalse(self.window.check_visible_btn.isEnabled())

    def test_filter_change_retains_editor_and_draft(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("編集中のメモ")
        self.window.merchant_edit.setText("変更した店舗")
        self.window.queue_search.setText("薬局")
        self.assertEqual(self.window.active_queue_path, self.paths[0])
        self.assertEqual(self.window.memo_edit.text(), "編集中のメモ")
        self.assertEqual(self.window.entries[self.paths[0]].draft.merchant, "変更した店舗")
        self.assertEqual(self.visible(), [self.paths[1]])
        self.assertIn("一覧では非表示", self.window.queue_filter_notice.text())
        self.window.queue.setCurrentRow(1)
        self.window.queue_search.clear()
        self.window.queue.setCurrentRow(0)
        self.assertEqual(self.window.memo_edit.text(), "編集中のメモ")

    def test_background_results_keep_open_image_and_update_other_rows(self):
        self.window.queue.setCurrentRow(2)
        self.filter("pending")
        self.window._batch_image_started(self.paths[2], 1, 1)
        self.assertEqual(self.visible(), [self.paths[2]])
        self.assertIn("条件の対象外", self.window.queue_filter_notice.text())
        result = _parse_receipt({"receipt_date": date.today().isoformat(), "merchant": "読取結果", "total_yen": 108})
        self.window._batch_image_succeeded(self.paths[2], result)
        self.assertEqual(self.window.merchant_edit.text(), "読取結果")
        self.assertEqual(self.visible(), [self.paths[2]])
        self.window._new_form()
        self.assertEqual(self.visible(), [])
        self.filter("ready")
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("ほかの読取中の修正")
        self.window._batch_image_succeeded(self.paths[3], result)
        self.assertIn(self.paths[3], self.visible())
        self.assertEqual(self.window.active_queue_path, self.paths[0])
        self.assertEqual(self.window.memo_edit.text(), "ほかの読取中の修正")

    def test_pending_reread_is_not_bulk_checked(self):
        self.window.pending_reads.add(self.paths[0])
        self.window._refresh_queue_item(self.paths[0])
        self.window.check_visible_btn.click()
        self.assertFalse(self.window.entries[self.paths[0]].checked)
        self.assertTrue(self.window.entries[self.paths[1]].checked)

    def test_startup_shows_all_but_retains_checks_and_editor(self):
        self.window.queue.setCurrentRow(0)
        self.filter("ready")
        self.window.check_visible_btn.click()
        self.window.memo_edit.setText("再開する編集")
        self.filter("ready")
        self.window.queue_search.setText("薬局")
        self.window.close()
        self.window = MainWindow()
        self.assertEqual(self.window.queue_filter.currentData(), "all")
        self.assertEqual(self.window.queue_search.text(), "")
        self.assertEqual(self.visible(), self.paths)
        self.assertTrue(self.window.entries[self.paths[0]].checked)
        self.assertEqual(self.window.memo_edit.text(), "再開する編集")
        self.assertNotIn("非表示 1 件", self.window.review_count_label.text())

    def test_hidden_checked_receipts_record_once_with_confirmation(self):
        self.filter("ready")
        self.window.check_visible_btn.click()
        self.window.queue_search.setText("薬局")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.No) as question:
            self.window.batch_save_btn.click()
        self.assertIn("非表示のチェック済みレシート 1 件", question.call_args.args[2])
        self.assertEqual(len(self.window.receipts), 0)
        self.assertTrue(self.window.entries[self.paths[0]].checked)
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.batch_save_btn.click()
        self.assertEqual(len(self.window.ledger.load()), 2)
        self.assertFalse(any(e.checked for e in self.window.entries.values()))
        self.assertFalse(self.window.batch_save_btn.isEnabled())

    def test_import_under_filter_does_not_replace_current_draft(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("保持する")
        self.filter("saved")
        extra = self.root / "new.png"
        Image.new("RGB", (80, 160), "white").save(extra)
        self.window._enqueue_images([str(extra)])
        self.assertEqual(self.visible(), [self.paths[4]])
        self.assertEqual(self.window.memo_edit.text(), "保持する")
        self.assertEqual(self.window.active_queue_path, self.paths[0])

    def test_remove_cancel_preserves_edits_and_checks(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("削除をキャンセル")
        self.window.check_visible_btn.click()
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.No):
            self.window.remove_image_actions["current"].trigger()
        self.assertEqual(len(self.window.entries), 5)
        self.assertEqual(self.window.memo_edit.text(), "削除をキャンセル")
        self.assertTrue(self.window.entries[self.paths[0]].checked)

    def test_remove_hidden_checked_then_restart_and_reimport(self):
        self.window.queue.setCurrentRow(0)
        self.filter("ready")
        self.window.check_visible_btn.click()
        self.window.queue_search.setText("薬局")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes) as question:
            self.window.remove_image_actions["checked"].trigger()
        self.assertIn("非表示の画像 1 件", question.call_args.args[2])
        self.assertIn("未記録の画像: 2 件", question.call_args.args[2])
        self.assertEqual(list(self.window.entries), self.paths[2:])
        self.assertIsNone(self.window.active_queue_path)
        self.assertEqual(self.window.merchant_edit.text(), "")
        self.assertTrue(self.window.preview.pixmap().isNull())
        # Restart without another autosave: deletion must already be durable.
        self.window._restoring_session = True
        self.window.close()
        self.window = MainWindow()
        self.assertEqual(list(self.window.entries), self.paths[2:])
        self.assertEqual(self.window.queue_search.text(), "")
        self.assertEqual(self.window.merchant_edit.text(), "")
        self.window.queue_search.clear()
        self.filter("all")
        self.window._enqueue_images([self.paths[0]])
        self.assertEqual(self.window.entries[self.paths[0]].state, "pending")
        self.assertIsNone(self.window.entries[self.paths[0]].draft)

    def test_remove_saved_keeps_original_excel_and_archived_image(self):
        from core import Receipt
        source = Path(self.paths[4])
        archived = self.root / "archived.png"
        archived.write_bytes(source.read_bytes())
        record = Receipt(merchant="保存済み", amount=108, image_path=str(archived))
        self.window.ledger.save([record])
        self.window.receipts = [record]
        excel_bytes = self.window.ledger.path.read_bytes()
        image_bytes = source.read_bytes()
        self.window.queue.setCurrentRow(4)
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.remove_image_actions["current"].trigger()
        self.assertNotIn(self.paths[4], self.window.entries)
        self.assertEqual(self.window.ledger.path.read_bytes(), excel_bytes)
        self.assertEqual(source.read_bytes(), image_bytes)
        self.assertEqual(archived.read_bytes(), image_bytes)
        self.assertEqual(self.window.receipts[0].receipt_id, record.receipt_id)

    def test_remove_visible_includes_unread_failed_and_saved_only_within_filter(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            for state, index in (("pending", 2), ("failed", 3), ("saved", 4)):
                self.filter(state)
                self.assertTrue(self.window.remove_image_actions["visible"].isEnabled())
                self.window.remove_image_actions["visible"].trigger()
                self.assertNotIn(self.paths[index], self.window.entries)
                self.assertFalse(self.window.remove_image_actions["visible"].isEnabled())
            self.filter("all")
            self.window.remove_image_actions["visible"].trigger()
        self.assertEqual(self.window.queue.count(), 0)
        self.assertFalse(self.window.remove_images_btn.isEnabled())
        self.assertTrue(all(Path(path).exists() for path in self.paths))

    def test_remove_save_failure_keeps_queue_draft_and_previous_snapshot(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("残す内容")
        self.window.check_visible_btn.click()
        self.window._save_session()
        before = self.window.session_store.path.read_bytes()
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes), \
                patch("session_store.os.replace", side_effect=OSError("disk full")), \
                patch("app.QMessageBox.warning") as warning:
            self.window.remove_image_actions["checked"].trigger()
        warning.assert_called_once()
        self.assertEqual(self.window.session_store.path.read_bytes(), before)
        self.assertEqual(len(self.window.entries), 5)
        self.assertEqual(self.window.memo_edit.text(), "残す内容")
        self.assertTrue(self.window.entries[self.paths[0]].checked)

    def test_remove_is_blocked_during_worker_and_does_not_affect_callbacks(self):
        self.window.queue.setCurrentRow(0)
        self.window.workers.append(object())
        try:
            self.window._update_read_controls()
            self.assertFalse(self.window.remove_images_btn.isEnabled())
            with patch("app.QMessageBox.question") as question:
                self.window._remove_images("visible")
            question.assert_not_called()
            self.assertEqual(len(self.window.entries), 5)
        finally:
            self.window.workers.clear()
            self.window._update_read_controls()
        self.assertTrue(self.window.remove_images_btn.isEnabled())

    def test_remove_other_rows_preserves_current_editor_view_and_checks(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("現在の編集")
        self.window.preview.rotate(90)
        self.window.preview.actual_size()
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        self.filter("failed")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.remove_image_actions["visible"].trigger()
        self.assertEqual(self.window.active_queue_path, self.paths[0])
        self.assertEqual(self.window.memo_edit.text(), "現在の編集")
        self.assertEqual(self.window.preview.angle, 90)
        self.assertEqual(self.window.preview.view.transform().m11(), 1)
        self.assertTrue(self.window.entries[self.paths[0]].checked)
        self.assertIn("非表示 1 件", self.window.review_count_label.text())

    def test_remove_all_preserves_independent_manual_form(self):
        self.window._new_form()
        self.window.merchant_edit.setText("手入力の途中")
        self.window.memo_edit.setText("消さない")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.remove_image_actions["visible"].trigger()
        self.assertEqual(self.window.merchant_edit.text(), "手入力の途中")
        self.window._restoring_session = True
        self.window.close()
        self.window = MainWindow()
        self.assertFalse(self.window.entries)
        self.assertEqual(self.window.memo_edit.text(), "消さない")

    def test_saved_individual_check_restores_and_does_not_enable_recording(self):
        self.filter("saved")
        item = self.window._queue_item(self.paths[4])
        self.assertTrue(item.flags() & Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        self.assertTrue(self.window.entries[self.paths[4]].checked)
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.assertIn("0 件", self.window.batch_save_btn.text())
        self.assertTrue(self.window.remove_image_actions["checked"].isEnabled())
        self.window.close()
        self.window = MainWindow()
        self.assertTrue(self.window.entries[self.paths[4]].checked)
        self.assertEqual(self.window._queue_item(self.paths[4]).checkState(), Qt.CheckState.Checked)
        self.assertFalse(self.window.batch_save_btn.isEnabled())

    def test_bulk_select_and_remove_saved_images_keeps_excel_records(self):
        self.filter("ready")
        self.window.check_visible_btn.click()
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.batch_save_btn.click()
        before = self.window.ledger.path.read_bytes()
        self.filter("saved")
        self.window.check_visible_btn.click()
        self.assertEqual([p for p, e in self.window.entries.items() if e.checked], self.paths[:2] + [self.paths[4]])
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.window.queue_search.setText("Apple")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes) as question:
            self.window.remove_image_actions["checked"].trigger()
        self.assertIn("非表示の画像 2 件", question.call_args.args[2])
        self.assertEqual(list(self.window.entries), self.paths[2:4])
        self.assertEqual(self.window.ledger.path.read_bytes(), before)
        self.assertEqual(len(self.window.ledger.load()), 2)
        self.assertTrue(all(Path(p).exists() for p in self.paths))

    def test_mixed_checks_record_only_unsaved_and_leave_saved_selected(self):
        self.window.check_visible_btn.click()
        self.assertIn("選択 5 件（未記録 4 件・記録済み 1 件", self.window.review_count_label.text())
        self.assertEqual(self.window.remove_checked_btn.text(), "選択 5 件を一覧から削除")
        self.assertEqual(self.window.batch_save_btn.text(), "未記録 2 件をExcelに記録")
        self.assertIn("2 件", self.window.batch_save_btn.text())
        self.window.queue_search.setText("saved")
        self.assertIn("非表示 2 件", self.window.batch_save_btn.toolTip())
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes) as question:
            self.window.batch_save_btn.click()
        self.assertIn("チェックした 2 枚", question.call_args.args[2])
        self.assertIn("未読・失敗の 2 枚は記録対象外", question.call_args.args[2])
        self.assertEqual(len(self.window.ledger.load()), 2)
        self.assertTrue(self.window.entries[self.paths[4]].checked)
        self.assertTrue(all(self.window.entries[p].checked for p in self.paths[2:4]))
        self.assertFalse(self.window.entries[self.paths[0]].checked)
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.assertIn("0 件", self.window.batch_save_btn.text())

    def test_unread_and_failed_can_be_checked_by_keyboard_and_restored(self):
        from PySide6.QtTest import QTest
        for index in (2, 3):
            self.window.queue.setCurrentRow(index)
            item = self.window.queue.item(index)
            self.assertTrue(item.flags() & Qt.ItemFlag.ItemIsUserCheckable)
            QTest.keyClick(self.window.queue, Qt.Key.Key_Space)
            self.assertTrue(self.window.entries[self.paths[index]].checked)
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.assertTrue(self.window.remove_checked_btn.isEnabled())
        self.window.close()
        self.window = MainWindow()
        for path, state in zip(self.paths[2:4], ("pending", "failed")):
            self.assertTrue(self.window.entries[path].checked)
            self.assertEqual(self.window.entries[path].state, state)
            self.assertEqual(self.window._queue_item(path).checkState(), Qt.CheckState.Checked)

    def test_bulk_remove_unread_and_failed_with_hidden_selection_and_undo(self):
        self.filter("pending")
        self.window.check_visible_btn.click()
        self.filter("failed")
        self.window.check_visible_btn.click()
        self.assertEqual([p for p, e in self.window.entries.items() if e.checked], self.paths[2:4])
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes) as question:
            self.window.remove_checked_btn.click()
        self.assertIn("未記録の画像: 2 件", question.call_args.args[2])
        self.assertIn("非表示の画像 1 件", question.call_args.args[2])
        self.assertTrue(all(p not in self.window.entries for p in self.paths[2:4]))
        self.assertTrue(all(Path(p).is_file() for p in self.paths))
        self.assertFalse(self.window.ledger.path.exists())
        self.window.close()
        self.window = MainWindow()
        self.assertTrue(all(p not in self.window.entries for p in self.paths[2:4]))
        self.window.undo_remove_btn.click()
        self.assertEqual(list(self.window.entries), self.paths)
        self.assertEqual([self.window.entries[p].state for p in self.paths[2:4]], ["pending", "failed"])
        self.assertTrue(all(self.window.entries[p].checked for p in self.paths[2:4]))

    def test_waiting_and_reading_images_are_excluded_from_checking(self):
        self.window.pending_reads.add(self.paths[2])
        self.window.entries[self.paths[3]].state = "reading"
        for path in self.paths[2:4]:
            self.window._refresh_queue_item(path)
            self.assertFalse(self.window._queue_item(path).flags() & Qt.ItemFlag.ItemIsUserCheckable)
        self.window.check_visible_btn.click()
        self.assertFalse(any(self.window.entries[p].checked for p in self.paths[2:4]))

    def test_undo_restores_order_draft_checks_and_view_after_restart(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("消す前に修正したメモ")
        self.window.preview.rotate(90)
        self.window.preview.actual_size()
        self.window.queue.item(0).setCheckState(Qt.CheckState.Checked)
        self.window.queue.item(1).setCheckState(Qt.CheckState.Checked)
        self.window.queue_search.setText("薬局")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.remove_checked_btn.click()
        self.assertIn("2 件", self.window.undo_remove_btn.text())
        self.window.close()
        self.window = MainWindow()
        self.window.undo_remove_btn.click()
        self.assertEqual(list(self.window.entries), self.paths)
        self.assertEqual(self.window.active_queue_path, self.paths[0])
        self.assertEqual(self.window.memo_edit.text(), "消す前に修正したメモ")
        self.assertEqual(self.window.preview.angle, 90)
        self.assertEqual(self.window.preview.view.transform().m11(), 1)
        self.assertTrue(all(self.window.entries[p].checked for p in self.paths[:2]))
        self.assertFalse(self.window.undo_remove_btn.isEnabled())
        self.window._restoring_session = True
        self.window.close()
        self.window = MainWindow()
        self.assertEqual(list(self.window.entries), self.paths)
        self.assertEqual(self.window.memo_edit.text(), "消す前に修正したメモ")
        self.assertEqual(self.window.preview.angle, 90)

    def test_undo_multiple_deletions_in_reverse_order_preserves_other_edits(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.queue.setCurrentRow(1)
            self.window._remove_images("current")
            self.window.queue.setCurrentItem(self.window._queue_item(self.paths[3]))
            self.window._remove_images("current")
        self.window.queue.setCurrentItem(self.window._queue_item(self.paths[0]))
        self.window.memo_edit.setText("削除後の別レシートの編集")
        self.window.undo_remove_btn.click()
        self.assertIn(self.paths[3], self.window.entries)
        self.assertNotIn(self.paths[1], self.window.entries)
        self.window.undo_remove_btn.click()
        self.assertEqual(list(self.window.entries), self.paths)
        self.assertEqual(self.window.memo_edit.text(), "削除後の別レシートの編集")
        self.assertEqual(self.window.active_queue_path, self.paths[0])

    def test_undo_save_failure_keeps_history_and_current_state(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._remove_images("visible")
        before = self.window.session_store.path.read_bytes()
        with patch("session_store.os.replace", side_effect=OSError("disk full")), patch("app.QMessageBox.warning") as warning:
            self.window.undo_remove_btn.click()
        warning.assert_called_once()
        self.assertEqual(self.window.session_store.path.read_bytes(), before)
        self.assertFalse(self.window.entries)
        self.assertTrue(self.window.undo_remove_btn.isEnabled())
        self.window.undo_remove_btn.click()
        self.assertEqual(list(self.window.entries), self.paths)

    def test_undo_reimport_conflict_requires_consent_and_restores_old_draft(self):
        self.window.queue.setCurrentRow(0)
        self.window.memo_edit.setText("削除前")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._remove_images("current")
        self.window._enqueue_images([self.paths[0]])
        self.window.queue.setCurrentItem(self.window._queue_item(self.paths[0]))
        self.window.memo_edit.setText("再追加後の新しい内容")
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.No):
            self.window.undo_remove_btn.click()
        self.assertEqual(self.window.memo_edit.text(), "再追加後の新しい内容")
        self.assertTrue(self.window.undo_remove_btn.isEnabled())
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.undo_remove_btn.click()
        self.assertEqual(self.window.memo_edit.text(), "削除前")
        self.assertEqual(self.window.queue.count(), 5)
        self.assertEqual(list(self.window.entries), self.paths)

    def test_undo_missing_image_recovers_text_and_keeps_manual_form(self):
        self.window.queue.setCurrentRow(0)
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._remove_images("current")
        Path(self.paths[0]).unlink()
        self.window.merchant_edit.setText("独立した手入力")
        self.window.undo_remove_btn.click()
        self.assertEqual(self.window.merchant_edit.text(), "独立した手入力")
        self.assertIsNone(self.window.active_queue_path)
        self.assertEqual(self.window.entries[self.paths[0]].draft.amount, "108")

    def test_undo_saved_receipts_does_not_modify_workbook_or_rerecord(self):
        self.filter("ready")
        self.window.check_visible_btn.click()
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.batch_save_btn.click()
            self.filter("saved")
            self.window.check_visible_btn.click()
            self.window.remove_checked_btn.click()
        before = self.window.ledger.path.read_bytes()
        self.window.undo_remove_btn.click()
        self.assertEqual(self.window.ledger.path.read_bytes(), before)
        self.assertTrue(self.window.entries[self.paths[0]].checked)
        self.assertEqual(self.window.entries[self.paths[0]].state, "saved")
        self.assertFalse(self.window.batch_save_btn.isEnabled())
        self.assertIn("未記録 0 件・記録済み 3 件", self.window.review_count_label.text())

    def test_undo_blocked_during_work_and_discard_clears_history(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._remove_images("visible")
        self.window.workers.append(object())
        try:
            self.window._update_read_controls()
            self.assertFalse(self.window.undo_remove_btn.isEnabled())
            self.window._undo_remove_images()
            self.assertFalse(self.window.entries)
        finally:
            self.window.workers.clear()
            self.window._update_read_controls()
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window._discard_session()
        self.assertFalse(self.window.undo_remove_btn.isEnabled())
        self.window.close()
        self.window = MainWindow()
        self.assertFalse(self.window.removed_batches)

    def test_undo_history_keeps_last_ten_operations(self):
        with patch("app.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            for index in range(11):
                path = self.root / f"extra-{index}.png"
                Image.new("RGB", (80, 160), "white").save(path)
                self.window._enqueue_images([str(path)])
                self.window.queue.setCurrentItem(self.window._queue_item(str(path)))
                self.window._remove_images("current")
        self.assertEqual(len(self.window.removed_batches), 10)
        self.assertTrue(self.window.removed_batches[0]["entries"][0]["entry"]["path"].endswith("extra-1.png"))
        for _ in range(10):
            self.window.undo_remove_btn.click()
        self.assertFalse(self.window.undo_remove_btn.isEnabled())
        self.assertEqual(len(self.window.entries), 15)


if __name__ == "__main__":
    unittest.main()
