"""Windows desktop UI for ReceiptLedger."""

from __future__ import annotations

import os
import shutil
from dataclasses import asdict, replace
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QEvent, QObject, Qt, QThread, Signal, QSignalBlocker, QTimer, QLockFile
from PySide6.QtGui import QPixmap, QIcon
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFileDialog, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QStatusBar, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from config import Settings, app_data_dir, default_data_dir, get_token, save_token
from data_backup import create_bundle, restore_bundle, activate_restored_state, recover_activation
from storage import copy_image, copy_workbook_rebased, writable_directory, file_usage, folder_files, regular_backups, prune_backups, size_text
from version import APP_VERSION
from core import (
    ExcelLedger, LedgerError, Receipt, WorkbookChanged,
    discrepancy, duplicate_candidates, validate_receipt,
)
from batch import BatchReadWorker, QueueEntry, ReceiptDraft
from appearance import apply_theme, system_theme
from lm_client import LMStudioClient, LMStudioError, ExtractedReceipt, normalize_base_url
from image_io import preview_png, supported_image
from session_store import SessionStore, receipt_signature
import receipt_panel
import settings_panel
import diagnostics


class TaskWorker(QThread):
    result = Signal(object)
    failed = Signal(str)

    def __init__(self, task, parent=None):
        super().__init__(parent)
        self.task = task

    def run(self):
        try:
            self.result.emit(self.task())
        except Exception as exc:
            diagnostics.record("task_failed", error=diagnostics.error_code(exc))
            self.failed.emit(str(exc))


class ImageDropFilter(QObject):
    """Accept local image drops on any part of the main window."""

    def __init__(self, window: "MainWindow"):
        super().__init__(window)
        self.window = window

    def eventFilter(self, watched, event):
        if event.type() not in (QEvent.Type.DragEnter, QEvent.Type.DragMove, QEvent.Type.Drop):
            return False
        if not isinstance(watched, QWidget) or watched.window() is not self.window:
            return False
        mime = event.mimeData()
        if not mime.hasUrls():
            return False
        paths = [url.toLocalFile() for url in mime.urls() if url.isLocalFile()]
        paths = [path for path in paths if supported_image(path)]
        if not paths:
            return False
        event.acceptProposedAction()
        if event.type() == QEvent.Type.Drop:
            self.window._enqueue_images(paths)
        return True


def field(label: str, placeholder: str = "") -> QLineEdit:
    editor = QLineEdit()
    editor.setPlaceholderText(placeholder)
    editor.setAccessibleName(label)
    return editor


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"ReceiptLedger — レシート家計簿 v{APP_VERSION}")
        self.setWindowIcon(QIcon(str(Path(__file__).resolve().parent / "assets/receipt-ledger.ico")))
        self.resize(1260, 820)
        recover_activation()
        self.settings = Settings.load()
        if self.settings.theme not in ("system", "light", "dark"):
            self.settings.theme = "system"
        if not self.settings.workbook_path:
            self.settings.workbook_path = str(default_data_dir() / "excel" / "レシート家計簿.xlsx")
        if not self.settings.image_folder:
            self.settings.image_folder = str(default_data_dir() / "images")
        self.ledger = self._ledger_for(self.settings.workbook_path)
        self.receipts: list[Receipt] = []
        self.current_image = ""
        self.editing_id: str | None = None
        self.workers: list[QThread] = []
        self.entries: dict[str, QueueEntry] = {}
        self.active_queue_path: str | None = None
        self._filter_pinned_path: str | None = None
        self.removed_batches: list[dict] = []
        self.batch_worker: BatchReadWorker | None = None
        self.pending_reads: set[str] = set()
        self.session_store = SessionStore(app_data_dir())
        self._restoring_session = True
        self._pending_commit = {}
        self._close_after_batch = False
        self._session_notice = ""
        self._session_write_blocked = False
        self._data_busy = False
        self._setup_ui()
        diagnostics.record("app_started")
        self.theme_timer = QTimer(self)
        self.theme_timer.setInterval(2000)
        self.theme_timer.timeout.connect(self._sync_system_theme)
        self.theme_timer.start()
        self.drop_filter = ImageDropFilter(self)
        QApplication.instance().installEventFilter(self.drop_filter)
        self._new_form()
        self._load_ledger()
        self._restore_session(reset_filters=True)
        self._restoring_session = False
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(1000)
        self.autosave_timer.timeout.connect(self._save_session)
        self.autosave_timer.start()
        self._update_read_controls()

    def _ledger_for(self, path):
        ledger = ExcelLedger(path)
        ledger.backup_limit = self.settings.backup_generations
        return ledger

    def _setup_ui(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.addTab(self._receipt_tab(), "レシート登録")
        self.tabs.addTab(self._ledger_tab(), "家計簿・Excel")
        self.tabs.addTab(self._settings_tab(), "設定")
        self.theme_selector = self._theme_selector()
        self.tabs.setCornerWidget(self.theme_selector, Qt.Corner.TopRightCorner)
        for editor in (self.date_edit, self.merchant_edit, self.amount_edit, self.tax_edit,
                       self.discount_edit, self.payment_edit, self.memo_edit):
            editor.textEdited.connect(self._form_edited)
        self.category_box.activated.connect(self._form_edited)
        self.item_table.itemChanged.connect(self._form_edited)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("準備完了")
        self._apply_theme()

    def _session_snapshot(self):
        self._stash_current_draft()
        form = self._capture_draft()
        return {"version": 1, "workbook": str(self.ledger.path),
                    "entries": [asdict(entry) for entry in self.entries.values()],
                    "active": self.active_queue_path, "image": self.current_image,
                    "editing_id": self.editing_id,
                    "form": asdict(form) if self.active_queue_path is None and form.has_content() else None,
                    "view": self.preview.state(), "details_open": self.details_toggle.isChecked(),
                    "queue_filter": self.queue_filter.currentData(), "queue_search": self.queue_search.text(),
                    "filter_pinned": self._filter_pinned_path,
                    "commit": self._pending_commit, "removed_batches": self.removed_batches}

    def _save_session(self):
        if self._session_write_blocked:
            return False
        if self._restoring_session:
            return True
        try:
            self.session_store.save(self._session_snapshot())
            self._session_log_error = None
            self.autosave_label.setText(self._session_notice + "下書き保存済み・Excelへの保存は記録ボタンから")
            return True
        except (OSError, ValueError, TypeError) as exc:
            code = diagnostics.error_code(exc)
            if getattr(self, "_session_log_error", None) != code:
                diagnostics.record("session_failed", error=code)
                self._session_log_error = code
            self.autosave_label.setText(f"下書きを保存できません: {exc}")
            return False

    @staticmethod
    def _decode_draft(data):
        if data is None:
            return None
        if not isinstance(data, dict):
            raise ValueError("入力内容の形式が不正です")
        draft = ReceiptDraft(**data)
        if any(not isinstance(value, str) for key, value in asdict(draft).items() if key != "items"):
            raise ValueError("入力内容の形式が不正です")
        if not isinstance(draft.items, list) or any(not isinstance(row, list) or len(row) != 5
                or any(not isinstance(value, str) for value in row) for row in draft.items):
            raise ValueError("商品明細の形式が不正です")
        return draft

    def _restore_session(self, *, reset_filters=False):
        try:
            data = self.session_store.load()
            if not data:
                return
            removed_batches = data.get("removed_batches", [])
            if not isinstance(removed_batches, list):
                raise ValueError("削除履歴の形式が不正です")
            for batch in removed_batches:
                if (not isinstance(batch, dict) or not isinstance(batch.get("entries"), list)
                        or not batch["entries"] or not isinstance(batch.get("view", {}), dict)
                        or not isinstance(batch.get("commit", []), list)):
                    raise ValueError("削除履歴の形式が不正です")
                for saved in batch["entries"]:
                    if not isinstance(saved, dict) or not isinstance(saved.get("index"), int) or saved["index"] < 0:
                        raise ValueError("削除履歴の位置が不正です")
                    self._decode_queue_entry(saved["entry"])
            recovered = []
            for raw in data["entries"]:
                raw = dict(raw)
                raw["draft"] = self._decode_draft(raw.get("draft"))
                entry = QueueEntry(**raw)
                if not isinstance(entry.path, str) or entry.state not in ("pending", "reading", "ready", "failed", "saved"):
                    raise ValueError("画像一覧の形式が不正です")
                if entry.state == "reading":
                    entry.state = "pending"
                    entry.checked = entry.reviewed = False
                recovered.append(entry)
            form = self._decode_draft(data.get("form"))
            if not data["workbook"] or Path(data["workbook"]).suffix.lower() != ".xlsx":
                raise ValueError("下書きの記録先が不正です")
            ledger = self._ledger_for(data["workbook"])
            records = ledger.load()
            commits = data.get("commit") or []
            recorded_paths = set()
            recorded_form = False
            for commit in commits:
                match = next((r for r in records if r.receipt_id == commit["id"]), None)
                current = next((entry.draft for entry in recovered if entry.path == commit["path"]), None) if commit["path"] else form
                unchanged = current is not None and asdict(current) == commit.get("draft")
                if match and unchanged and receipt_signature(match) == commit["signature"]:
                    if commit["path"]:
                        recorded_paths.add(commit["path"])
                    else:
                        recorded_form = True
            self.removed_batches = removed_batches[-10:]
            self.ledger, self.receipts = ledger, records
            self.settings.workbook_path = str(ledger.path)
            self.workbook_edit.setText(str(ledger.path))
            self._refresh_months()
            self._refresh_ledger()
            with QSignalBlocker(self.queue):
                for entry in recovered:
                    if entry.path in recorded_paths:
                        entry.state, entry.checked, entry.reviewed = "saved", False, False
                        entry.saved_to = str(ledger.path)
                    self.entries[entry.path] = entry
                    self.queue.addItem(Path(entry.path).name)
                    self.queue.item(self.queue.count() - 1).setData(Qt.ItemDataRole.UserRole, entry.path)
                    self._refresh_queue_item(entry.path)
            active = data.get("active")
            if active in self.entries:
                self.queue.setCurrentItem(self._queue_item(active))
            elif form and not recorded_form:
                self._restore_draft(form)
                # Preserve edit identity so resuming an edit updates the same row.
                editing_id = data.get("editing_id")
                if editing_id and not any(r.receipt_id == editing_id for r in records):
                    self._session_notice = "元の支出が見つからないため新規の下書きとして復元しました。 "
                else:
                    self.editing_id = editing_id
                if self.editing_id:
                    self.save_btn.setText("変更を保存")
                if data.get("image"):
                    self._set_image(data["image"])
            self.preview.restore_state(data.get("view") or {})
            self.details_toggle.setChecked(bool(data.get("details_open")))
            with QSignalBlocker(self.queue_filter), QSignalBlocker(self.queue_search):
                index = self.queue_filter.findData("all" if reset_filters else data.get("queue_filter", "all"))
                self.queue_filter.setCurrentIndex(max(0, index))
                self.queue_search.setText("" if reset_filters else str(data.get("queue_search", "")))
            self._filter_pinned_path = active if not reset_filters and data.get("filter_pinned") == active else None
            self._apply_queue_filter(preserve_active=False, reset_pin=False)
            if recovered or form:
                self._session_notice += "前回の下書きを復元しました。 "
                self.batch_status.setText("下書きを復元しました。中断した画像は「未読を一括読み込み」で再開できます。")
            self.autosave_label.setText(self._session_notice or "下書きを自動保存します")
        except (OSError, ValueError, TypeError, KeyError, LedgerError) as exc:
            self._session_notice = f"復元できなかった下書きがあります: {exc} "
            self.autosave_label.setText(self._session_notice)
            # Never overwrite recovery data that could not be loaded.
            try:
                if self.session_store.path.exists():
                    self.session_store.preserve_invalid()
            except OSError:
                self._session_write_blocked = True

    def _stage_commit(self, prepared):
        self._stash_current_draft()
        self._pending_commit = [{"path": path or "", "id": receipt.receipt_id,
                                 "signature": receipt_signature(receipt),
                                 "draft": asdict(self.entries[path].draft if path in self.entries else self._capture_draft())}
                                for path, receipt in prepared]
        if not self._save_session():
            raise LedgerError("下書きの保存に失敗しました。保存先の空き容量・アクセス権を確認してください。")

    def _discard_session(self):
        if self.workers:
            return
        answer = QMessageBox.question(self, "下書きを消去", "未記録の入力内容・取り込んだ画像一覧・削除を元に戻すための履歴を消去します。Excelと元画像は残ります。続けますか？",
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                      QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.active_queue_path = None
        self.entries.clear()
        self.removed_batches.clear()
        with QSignalBlocker(self.queue):
            self.queue.clear()
        self._pending_commit = []
        self._session_notice = ""
        self._new_form()
        self._save_session()

    def _apply_theme(self):
        self.effective_theme = system_theme() if self.settings.theme == "system" else self.settings.theme
        apply_theme(self, self.effective_theme)
        self.queue_delegate.theme = self.effective_theme
        self.queue.viewport().update()

    def _theme_selector(self):
        selector = QComboBox()
        selector.setAccessibleName("表示モード")
        for label, value in (("Windowsに合わせる", "system"), ("ライト", "light"), ("ダーク", "dark")):
            selector.addItem(label, value)
        selector.setCurrentIndex(selector.findData(self.settings.theme))
        selector.currentIndexChanged.connect(lambda: self._change_theme(selector.currentData()))
        return selector

    def _sync_system_theme(self):
        if self.settings.theme == "system" and system_theme() != self.effective_theme:
            self._apply_theme()

    def _change_theme(self, mode):
        self.settings.theme = mode
        for selector in (self.theme_box, self.theme_selector):
            with QSignalBlocker(selector):
                selector.setCurrentIndex(selector.findData(mode))
        self._apply_theme()
        try:
            self.settings.save()
        except OSError as exc:
            QMessageBox.warning(self, "表示設定を保存できませんでした", str(exc))

    def _receipt_tab(self) -> QWidget:
        return receipt_panel.build(self, field)

    def _ledger_tab(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)
        file_group = QGroupBox("Excelファイル")
        file_layout = QVBoxLayout(file_group)
        self.workbook_label = QLabel()
        self.workbook_label.setWordWrap(True)
        self.workbook_label.setTextFormat(Qt.TextFormat.PlainText)
        self.workbook_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        file_layout.addWidget(self.workbook_label)
        file_hint = QLabel("登録した内容はテンプレート形式で自動保存されます。Excelを選ぶと、そのファイルへ追記します。")
        file_hint.setWordWrap(True)
        file_layout.addWidget(file_hint)
        file_actions = QHBoxLayout()
        self.save_excel_btn = QPushButton("Excelを保存")
        self.save_excel_btn.setObjectName("primary")
        self.save_excel_btn.setToolTip("登録済みの全レシートを保存し、ダッシュボードとレシート別表示を更新します。")
        self.save_excel_btn.clicked.connect(self._save_workbook)
        save_as_btn = QPushButton("名前を付けて保存")
        save_as_btn.clicked.connect(self._save_workbook_as)
        choose_btn = QPushButton("保存先Excelを選ぶ")
        choose_btn.clicked.connect(self._select_workbook)
        open_btn = QPushButton("Excelで表示")
        open_btn.clicked.connect(self._open_workbook)
        self.delete_excel_btn = QPushButton("Excelファイルを削除")
        self.delete_excel_btn.clicked.connect(self._delete_workbook)
        for button in (self.save_excel_btn, save_as_btn, choose_btn, open_btn, self.delete_excel_btn):
            file_actions.addWidget(button)
        file_layout.addLayout(file_actions)
        layout.addWidget(file_group)
        tools = QHBoxLayout()
        self.search_edit = field("検索", "店舗・費目・メモを検索")
        self.search_edit.textChanged.connect(self._refresh_ledger)
        self.month_box = QComboBox()
        self.month_box.addItem("すべての月")
        self.month_box.currentIndexChanged.connect(self._refresh_ledger)
        reload_btn = QPushButton("再読み込み")
        reload_btn.clicked.connect(self._load_ledger)
        tools.addWidget(self.search_edit, 2)
        tools.addWidget(self.month_box)
        tools.addWidget(reload_btn)
        layout.addLayout(tools)
        self.total_label = QLabel("合計: 0 円")
        self.total_label.setStyleSheet("font-size: 18px; font-weight: bold; padding: 8px;")
        layout.addWidget(self.total_label)
        self.category_summary_label = QLabel("")
        self.category_summary_label.setWordWrap(True)
        self.category_summary_label.setObjectName("categorySummary")
        layout.addWidget(self.category_summary_label)
        self.ledger_table = QTableWidget(0, 7)
        self.ledger_table.setHorizontalHeaderLabels(["購入日", "店舗", "税込合計", "商品数", "費目", "支払方法", "ID"])
        self.ledger_table.setColumnHidden(6, True)
        self.ledger_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.ledger_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.ledger_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.ledger_table.cellDoubleClicked.connect(lambda *_: self._show_selected_details())
        layout.addWidget(self.ledger_table, 1)
        bottom = QHBoxLayout()
        detail_btn = QPushButton("選択したレシートの明細を見る")
        detail_btn.clicked.connect(self._show_selected_details)
        edit_btn = QPushButton("選択した支出を編集")
        edit_btn.clicked.connect(self._edit_selected)
        delete_btn = QPushButton("選択した支出を削除")
        delete_btn.clicked.connect(self._delete_selected)
        bottom.addWidget(detail_btn)
        bottom.addWidget(edit_btn)
        bottom.addWidget(delete_btn)
        bottom.addStretch()
        layout.addLayout(bottom)
        return root

    def _settings_tab(self) -> QWidget:
        return settings_panel.build(self, field, get_token())

    def _copy_diagnostics(self):
        QApplication.clipboard().setText(diagnostics.report(
            count=len(self.entries), concurrency=self.settings.concurrent_reads))
        self.statusBar().showMessage("診断情報をコピーしました。", 5000)

    def _start_task(self, task, on_success, busy_text: str):
        if self.workers:
            return
        worker = TaskWorker(task, self)
        self.workers.append(worker)
        self._update_read_controls()
        self.statusBar().showMessage(busy_text)
        worker.result.connect(on_success)
        worker.failed.connect(lambda message: QMessageBox.warning(self, "処理できませんでした", message))
        worker.finished.connect(lambda: self._task_finished(worker))
        worker.start()

    def _task_finished(self, worker: TaskWorker):
        self.statusBar().showMessage("準備完了")
        if worker in self.workers:
            self.workers.remove(worker)
        worker.deleteLater()
        self._data_busy = False
        self._update_read_controls()

    def _update_read_controls(self):
        busy = bool(self.workers)
        self.add_images_btn.setEnabled(not self._data_busy)
        self.theme_selector.setEnabled(not self._data_busy)
        for button in (self.read_btn, self.batch_read_btn, self.retry_btn, self.new_btn, self.batch_save_btn):
            button.setEnabled(not busy)
        entry = self.entries.get(self.active_queue_path)
        awaiting_result = bool(entry and (entry.path in self.pending_reads or entry.state == "reading"))
        can_edit = not awaiting_result and (not busy or (self.batch_worker is not None
                                                       and entry is not None and entry.state in ("ready", "failed")))
        self.receipt_editor.setEnabled(can_edit)
        self.save_btn.setEnabled(not busy)
        self._update_review_selection()
        self.stop_btn.setEnabled(self.batch_worker is not None and not self.batch_worker.isInterruptionRequested())
        self.tabs.setTabEnabled(0, not self._data_busy)
        self.tabs.setTabEnabled(1, self.batch_worker is None and not self._data_busy)
        self.tabs.setTabEnabled(2, self.batch_worker is None)
        self.tabs.widget(2).setEnabled(not busy)

    def _can_review(self, entry):
        return entry.state == "ready" and entry.draft is not None and entry.path not in self.pending_reads

    def _can_check(self, entry):
        return entry.state in ("pending", "failed", "ready", "saved") and entry.path not in self.pending_reads

    def _update_review_selection(self):
        checked = sum(entry.checked for entry in self.entries.values())
        saved = sum(entry.state == "saved" for entry in self.entries.values())
        visible = [self.queue.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.queue.count())
                   if not self.queue.item(i).isHidden()]
        hidden_checked = sum(entry.checked and entry.path not in visible for entry in self.entries.values())
        checked_saved = sum(entry.checked and entry.state == "saved" for entry in self.entries.values())
        self.review_count_label.setText(f"表示 {len(visible)} 件 ／ 全 {len(self.entries)} 件 ・ 記録済み {saved} 件\n"
                                       f"選択 {checked} 件（未記録 {checked - checked_saved} 件・記録済み {checked_saved} 件／非表示 {hidden_checked} 件）")
        self.check_visible_btn.setEnabled(any(self._can_check(self.entries[path]) and not self.entries[path].checked
                                              for path in visible))
        self.clear_checks_btn.setEnabled(bool(checked))
        recordable = [entry for entry in self.entries.values() if entry.checked and self._can_review(entry)]
        hidden_recordable = sum(entry.path not in visible for entry in recordable)
        self.batch_save_btn.setText(f"未記録 {len(recordable)} 件をExcelに記録")
        self.batch_save_btn.setToolTip(f"読み取り済みの未記録 {len(recordable)} 件が対象です（非表示 {hidden_recordable} 件を含む）。未読・失敗の画像は一覧からの削除用に選択できます。")
        self.batch_save_btn.setEnabled(not self.workers and bool(recordable))
        self.remove_images_btn.setEnabled(not self.workers and bool(self.entries))
        self.remove_checked_btn.setText(f"選択 {checked} 件を一覧から削除")
        self.remove_checked_btn.setToolTip(f"未記録 {checked - checked_saved} 件・記録済み {checked_saved} 件が対象です。非表示 {hidden_checked} 件も含みます。")
        self.remove_checked_btn.setEnabled(not self.workers and checked > 0)
        undo_count = len(self.removed_batches[-1]["entries"]) if self.removed_batches else 0
        self.undo_remove_btn.setText(f"削除を元に戻す（{undo_count} 件）" if undo_count else "削除を元に戻す")
        self.undo_remove_btn.setEnabled(not self.workers and undo_count > 0)
        for scope, label in (("current", "選択中の画像"), ("checked", "チェックした画像"), ("visible", "表示中の画像")):
            count = len(self._image_removal_targets(scope))
            self.remove_image_actions[scope].setText(f"{label}を削除（{count} 件）")
            self.remove_image_actions[scope].setEnabled(not self.workers and count > 0)

    def _image_removal_targets(self, scope):
        if scope == "current":
            return [self.active_queue_path] if self.active_queue_path in self.entries else []
        if scope == "checked":
            return [path for path, entry in self.entries.items() if entry.checked]
        if scope == "visible":
            return [self.queue.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.queue.count())
                    if not self.queue.item(i).isHidden()]
        return []

    def _remove_images(self, scope):
        # Batch callbacks retain references to all their paths until completion.
        if self.workers:
            return
        self._stash_current_draft()
        paths = self._image_removal_targets(scope)
        if not paths:
            return
        hidden = sum(self._queue_item(path).isHidden() for path in paths)
        unrecorded = sum(self.entries[path].state != "saved" for path in paths)
        message = (f"画像 {len(paths)} 件を取り込んだ一覧と下書きから削除しますか？\n"
                   "元の画像ファイルとExcelの記録・保存済み画像は残ります。\n"
                   f"未記録の画像: {unrecorded} 件（読み取り結果・編集内容も削除します）。")
        message += "\n削除後は「削除を元に戻す」で復元できます（直近10回まで）。"
        if hidden:
            message += f"\n一覧で非表示の画像 {hidden} 件を含みます。"
        message += "\n\n" + "\n".join(Path(path).name for path in paths[:8])
        if len(paths) > 8:
            message += f"\nほか {len(paths) - 8} 件"
        answer = QMessageBox.question(self, "画像を一覧から削除", message,
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                      QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        removed = set(paths)
        clear_form = self.active_queue_path in removed
        try:
            if self._session_write_blocked:
                raise OSError("下書きの保存を再開できないため、削除を確定できません。")
            data = self._session_snapshot()
            undo = {"entries": [{"index": index, "entry": entry} for index, entry in enumerate(data["entries"])
                                if entry["path"] in removed],
                    "active": data["active"] if clear_form else None,
                    "view": data["view"], "details_open": data["details_open"],
                    "commit": [commit for commit in self._pending_commit if commit["path"] in removed]}
            data["removed_batches"] = (self.removed_batches + [undo])[-10:]
            data["entries"] = [entry for entry in data["entries"] if entry["path"] not in removed]
            data["commit"] = [commit for commit in self._pending_commit if commit["path"] not in removed]
            if data["filter_pinned"] in removed:
                data["filter_pinned"] = None
            if clear_form:
                data.update(active=None, image="", editing_id=None, form=None, view={})
            # Commit the removal before changing the UI so failures retain all drafts.
            self.session_store.save(data)
        except (OSError, ValueError, TypeError) as exc:
            QMessageBox.warning(self, "削除できません", f"画像一覧と下書きは保持しています。\n{exc}")
            return
        with QSignalBlocker(self.queue):
            for index in range(self.queue.count() - 1, -1, -1):
                if self.queue.item(index).data(Qt.ItemDataRole.UserRole) in removed:
                    self.queue.takeItem(index)
            for path in paths:
                del self.entries[path]
            if clear_form:
                self.queue.setCurrentRow(-1)
        self._pending_commit = data["commit"]
        self.removed_batches = data["removed_batches"]
        self._filter_pinned_path = data["filter_pinned"]
        if clear_form:
            self._new_form()
        else:
            self._apply_queue_filter(preserve_active=False, reset_pin=False)
            self._update_read_controls()
        self.autosave_label.setText("一覧からの削除を保存しました。")
        self.batch_status.setText(f"画像 {len(paths)} 件を一覧から削除しました。")
        self.statusBar().showMessage(f"画像 {len(paths)} 件を一覧から削除しました。元画像とExcelは残っています。", 7000)

    def _decode_queue_entry(self, raw):
        raw = dict(raw)
        raw["draft"] = self._decode_draft(raw.get("draft"))
        entry = QueueEntry(**raw)
        if not isinstance(entry.path, str) or entry.state not in ("pending", "reading", "ready", "failed", "saved"):
            raise ValueError("画像一覧の形式が不正です")
        return entry

    def _undo_remove_images(self):
        if self.workers or not self.removed_batches:
            return
        undo = self.removed_batches[-1]
        self._stash_current_draft()
        restored = [(saved["index"], self._decode_queue_entry(saved["entry"])) for saved in undo["entries"]]
        restored_keys = {entry.path.casefold() for _, entry in restored}
        conflicts = [entry for entry in self.entries.values() if entry.path.casefold() in restored_keys]
        if conflicts:
            answer = QMessageBox.question(self, "同じ画像が取り込み済みです",
                f"同じ画像 {len(conflicts)} 件が再度取り込まれています。現在の編集内容・チェックを、削除前の内容で置き換えて復元しますか？\n\n"
                + "\n".join(Path(entry.path).name for entry in conflicts[:8]),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            if self._session_write_blocked:
                raise OSError("下書きを保存できないため、復元を確定できません。")
            data = self._session_snapshot()
            merged = [entry for entry in self.entries.values() if entry.path.casefold() not in restored_keys]
            for index, entry in sorted(restored, key=lambda pair: pair[0]):
                merged.insert(min(index, len(merged)), entry)
            data["entries"] = [asdict(entry) for entry in merged]
            data["removed_batches"] = self.removed_batches[:-1]
            data["commit"] = [commit for commit in self._pending_commit if commit["path"].casefold() not in restored_keys] + undo.get("commit", [])
            # Keep unrelated work open. Only reopen the recovered receipt if the form is empty,
            # or if the user explicitly agreed to replace the currently open duplicate.
            target = None
            if self.active_queue_path and self.active_queue_path.casefold() in restored_keys:
                target = next(entry.path for _, entry in restored if entry.path.casefold() == self.active_queue_path.casefold())
            elif self.active_queue_path is None and self.editing_id is None and not self._capture_draft().has_content():
                target = undo.get("active") or (restored[0][1].path if restored else None)
            if target:
                view = undo.get("view", {}) if target == undo.get("active") else {}
                data.update(active=target, image=target, form=None, editing_id=None, view=view,
                            details_open=undo.get("details_open", False), filter_pinned=target)
            self.session_store.save(data)
        except (OSError, ValueError, TypeError) as exc:
            QMessageBox.warning(self, "元に戻せません", f"削除履歴と現在の入力内容は保持しています。\n{exc}")
            return
        self.entries = {entry.path: entry for entry in merged}
        self.removed_batches = data["removed_batches"]
        self._pending_commit = data["commit"]
        with QSignalBlocker(self.queue):
            self.queue.clear()
            for entry in merged:
                self.queue.addItem(Path(entry.path).name)
                self.queue.item(self.queue.count() - 1).setData(Qt.ItemDataRole.UserRole, entry.path)
            for entry in merged:
                self._refresh_queue_item(entry.path)
            if target or self.active_queue_path:
                self.queue.setCurrentItem(self._queue_item(target or self.active_queue_path))
            else:
                self.queue.setCurrentRow(-1)
        if target:
            # Do not stash the old form over the recovered copy of a duplicate.
            self.active_queue_path = None
            self._select_queue_image(self.queue.currentRow())
            self.preview.restore_state(data["view"])
            self.details_toggle.setChecked(data["details_open"])
        self._filter_pinned_path = data["filter_pinned"]
        self._apply_queue_filter(preserve_active=False, reset_pin=False)
        self._update_read_controls()
        hidden = sum(self._queue_item(entry.path).isHidden() for _, entry in restored)
        self.batch_status.setText(f"画像 {len(restored)} 件を復元しました（現在の条件で非表示 {hidden} 件）。")
        self.autosave_label.setText("削除を元に戻した状態を保存しました。")
        self.statusBar().showMessage(f"削除を取り消し、画像 {len(restored)} 件と編集内容を復元しました。", 7000)

    def _matches_queue_filter(self, entry):
        state = self.queue_filter.currentData()
        query = self.queue_search.text().strip().casefold()
        merchant = entry.draft.merchant if entry.draft else ""
        return (state == "all" or entry.state == state) and (not query or
                query in Path(entry.path).name.casefold() or query in merchant.casefold())

    def _queue_filter_changed(self, *_):
        self._stash_current_draft()
        self._apply_queue_filter(preserve_active=False)

    def _apply_queue_filter(self, preserve_active=True, reset_pin=True):
        if not preserve_active and reset_pin:
            self._filter_pinned_path = None
        # Hiding a row must never select a different receipt or replace the editor.
        with QSignalBlocker(self.queue):
            for index in range(self.queue.count()):
                item = self.queue.item(index)
                path = item.data(Qt.ItemDataRole.UserRole)
                matches = self._matches_queue_filter(self.entries[path])
                if (preserve_active and path == self.active_queue_path and not item.isHidden() and not matches):
                    self._filter_pinned_path = path
                pinned = path == self._filter_pinned_path and path == self.active_queue_path
                item.setHidden(not matches and not pinned)
        notices = []
        if not any(not self.queue.item(i).isHidden() for i in range(self.queue.count())):
            notices.append("該当する画像はありません。条件を変更すると再表示できます。")
        active = self.entries.get(self.active_queue_path)
        if active and not self._matches_queue_filter(active):
            if self._filter_pinned_path == active.path:
                notices.append(f"開いている画像を表示中: {Path(active.path).name}（条件の対象外）。別の画像を選ぶと非表示になります。")
            else:
                notices.append(f"編集中: {Path(active.path).name}（一覧では非表示）。編集内容は保持されます。")
        self.queue_filter_notice.setText("\n".join(notices))
        self.queue_filter_notice.setVisible(bool(notices))
        self._update_review_selection()

    def _check_visible(self):
        self._stash_current_draft()
        for index in range(self.queue.count()):
            item = self.queue.item(index)
            entry = self.entries[item.data(Qt.ItemDataRole.UserRole)]
            if not item.isHidden() and self._can_check(entry):
                entry.checked = True
                with QSignalBlocker(self.queue):
                    item.setCheckState(Qt.CheckState.Checked)
        self._update_review_selection()
        self._save_session()

    def _clear_checks(self):
        for entry in self.entries.values():
            entry.checked = False
        with QSignalBlocker(self.queue):
            for index in range(self.queue.count()):
                self.queue.item(index).setCheckState(Qt.CheckState.Unchecked)
        self._update_review_selection()
        self._save_session()

    def _queue_item(self, path):
        return next((self.queue.item(i) for i in range(self.queue.count())
                     if self.queue.item(i).data(Qt.ItemDataRole.UserRole) == path), None)

    def _refresh_queue_item(self, path):
        entry = self.entries[path]
        item = self._queue_item(path)
        if item is None:
            return
        labels = {"pending": "未読", "reading": "読み取り中", "ready": "確認待ち",
                  "failed": "失敗", "saved": "記録済み"}
        label = "確認済み" if entry.state == "ready" and entry.reviewed else labels[entry.state]
        with QSignalBlocker(self.queue):
            merchant = entry.draft.merchant if entry.draft else ""
            item.setText(f"[{label}] {Path(path).name}" + (f" ｜ {merchant}" if merchant else ""))
            flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            if self._can_check(entry):
                flags |= Qt.ItemFlag.ItemIsUserCheckable
            item.setFlags(flags)
            item.setCheckState(Qt.CheckState.Checked if entry.checked else Qt.CheckState.Unchecked)
            item.setToolTip("\n".join(value for value in (path, entry.error,
                            f"記録先: {entry.saved_to}" if entry.saved_to else "") if value))
        self._apply_queue_filter()

    def _queue_item_changed(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path not in self.entries:
            return
        checked = item.checkState() == Qt.CheckState.Checked
        if path == self.active_queue_path:
            self._stash_current_draft()
        entry = self.entries[path]
        entry.checked = self._can_check(entry) and checked
        self._refresh_queue_item(path)

    def _capture_draft(self):
        draft = ReceiptDraft(
            purchased_on=self.date_edit.text(), merchant=self.merchant_edit.text(),
            amount=self.amount_edit.text(), tax=self.tax_edit.text(), discount=self.discount_edit.text(),
            category=self.category_box.currentText(), payment=self.payment_edit.text(), memo=self.memo_edit.text(),
            items=[[(self.item_table.item(row, col).text() if self.item_table.item(row, col) else "")
                    for col in range(5)] for row in range(self.item_table.rowCount())],
            warning=self.warning_label.text(),
        )
        editor = QApplication.focusWidget()
        row, col = self.item_table.currentRow(), self.item_table.currentColumn()
        if isinstance(editor, QLineEdit) and self.item_table.isAncestorOf(editor) and row >= 0 and col >= 0:
            draft.items[row][col] = editor.text()
        return draft

    def _restore_draft(self, draft):
        for editor, value in ((self.date_edit, draft.purchased_on), (self.merchant_edit, draft.merchant),
                              (self.amount_edit, draft.amount), (self.tax_edit, draft.tax),
                              (self.discount_edit, draft.discount), (self.payment_edit, draft.payment),
                              (self.memo_edit, draft.memo)):
            editor.setText(value)
        self.category_box.setCurrentText(draft.category)
        with QSignalBlocker(self.item_table):
            self.item_table.setRowCount(len(draft.items))
            for row, values in enumerate(draft.items):
                for col, value in enumerate(values):
                    self.item_table.setItem(row, col, QTableWidgetItem(value))
        self.warning_label.setText(draft.warning)

    def _form_edited(self, *_):
        entry = self.entries.get(self.active_queue_path)
        if entry and entry.reviewed:
            entry.reviewed = False
            self._refresh_queue_item(entry.path)
        self._stash_current_draft()
        self._update_read_controls()

    def _stash_current_draft(self):
        entry = self.entries.get(self.active_queue_path)
        if entry is None or entry.state in ("reading", "saved") or entry.path in self.pending_reads:
            return
        draft = self._capture_draft()
        if draft != entry.draft:
            entry.draft = draft
            entry.reviewed = False
            if draft.has_content():
                entry.state = "ready"
                entry.error = ""
            self._refresh_queue_item(entry.path)

    def _mark_reviewed(self):
        self._stash_current_draft()
        entry = self.entries.get(self.active_queue_path)
        if entry is None or not self._can_review(entry):
            QMessageBox.information(self, "確認する画像を選択", "読み取り済みの画像を選んで、内容を確認してください。")
            return
        try:
            validate_receipt(entry.draft.to_receipt(entry.path))
            entry.reviewed = True
            entry.checked = True
            self._refresh_queue_item(entry.path)
            self.statusBar().showMessage("確認済みにしてチェックを付けました。チェックしたレシートをまとめて記録できます。", 5000)
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "入力内容を確認してください", str(exc))

    def _add_images(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "レシート画像を選択", "", "画像 (*.jpg *.jpeg *.png *.webp *.heic *.heif)")
        self._enqueue_images(paths)

    def _enqueue_images(self, paths):
        if self._data_busy:
            self.statusBar().showMessage("データのバックアップ・復元が終わってから画像を追加してください。", 5000)
            return
        added = 0
        queued = {str(Path(self.queue.item(i).data(Qt.ItemDataRole.UserRole)).resolve()).casefold()
                  for i in range(self.queue.count())}
        for path in paths:
            if not supported_image(path) or not Path(path).is_file():
                continue
            resolved = str(Path(path).resolve())
            if resolved.casefold() not in queued:
                self.entries[resolved] = QueueEntry(resolved)
                self.queue.addItem(Path(resolved).name)
                self.queue.item(self.queue.count() - 1).setData(Qt.ItemDataRole.UserRole, resolved)
                self._refresh_queue_item(resolved)
                queued.add(resolved.casefold())
                added += 1
        if self.queue.count() and self.queue.currentRow() < 0:
            first = next((i for i in range(self.queue.count()) if not self.queue.item(i).isHidden()), None)
            if first is not None:
                self.queue.setCurrentRow(first)
        if added:
            self.statusBar().showMessage(f"画像を {added} 件追加しました。", 5000)

    def _select_queue_image(self, row: int):
        if row < 0:
            return
        path = self.queue.item(row).data(Qt.ItemDataRole.UserRole)
        self._stash_current_draft()
        self.active_queue_path = path
        self.editing_id = None
        self.save_btn.setText("この1件を記録")
        entry = self.entries[path]
        self._restore_draft(entry.draft or ReceiptDraft())
        if entry.state == "failed":
            self.warning_label.setText(f"読み取り失敗: {entry.error}\n再試行するか、手入力できます。")
        elif entry.state == "saved":
            self.warning_label.setText(f"この画像は記録済みです。記録先: {entry.saved_to}")
        self._set_image(path)
        self._apply_queue_filter(preserve_active=False)
        self._update_read_controls()

    def _set_image(self, path: str):
        error = "画像ファイルが見つかりません"
        for candidate in dict.fromkeys([path, *self._saved_image_candidates(path)]):
            try:
                pixmap = QPixmap()
                if not pixmap.loadFromData(preview_png(candidate, bounds=(4000, 6000)), "PNG"):
                    raise ValueError("画像を描画できません")
                self.current_image = candidate
                self.preview.setPixmap(pixmap)
                self.preview.setToolTip(f"表示中: {candidate}")
                if candidate != path:
                    self.statusBar().showMessage("取り込み元の画像を表示できないため、保存済みの画像を表示しています。", 7000)
                return
            except Exception as exc:
                error = str(exc)
        self.current_image = ""
        self.preview.setPixmap(QPixmap())
        self.preview.setText(f"画像を表示できません: {error}")

    def _saved_image_candidates(self, path):
        entry = self.entries.get(path)
        candidates, hashes = [], []
        if entry:
            candidates.append(entry.stored_image)
            hashes.append(entry.image_hash)
        records = self.receipts
        if entry and entry.saved_to and Path(entry.saved_to) != self.ledger.path:
            try:
                records = self._ledger_for(entry.saved_to).load()
            except LedgerError:
                records = []
        matching = [r for r in records if r.receipt_id == self.editing_id
                    or r.image_path == path or (entry and entry.image_hash and r.image_hash == entry.image_hash)]
        if entry and entry.state == "saved" and not matching and entry.draft:
            try:
                signature = receipt_signature(entry.draft.to_receipt())
                matches = [r for r in records if receipt_signature(r) == signature]
                if len(matches) == 1:
                    matching = matches
            except (LedgerError, ValueError):
                pass
        for receipt in matching:
            candidates.append(receipt.image_path)
            hashes.append(receipt.image_hash)
        for folder in (Path(self.settings.image_folder), app_data_dir() / "images", default_data_dir() / "images"):
            for candidate in list(candidates):
                if candidate:
                    candidates.append(str(folder / Path(candidate).name))
            for digest in set(hashes):
                if digest and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest):
                    candidates.extend(str(item) for item in folder.glob(digest + ".*"))
        return [candidate for candidate in dict.fromkeys(candidates) if candidate and Path(candidate).is_file()]

    def _resolve_image_path(self, path):
        if path and Path(path).is_file():
            return path
        candidates = self._saved_image_candidates(path)
        return candidates[0] if candidates else path

    def _new_form(self, clear_image: bool = True):
        self._stash_current_draft()
        self.active_queue_path = None
        with QSignalBlocker(self.queue):
            self.queue.setCurrentRow(-1)
        self.editing_id = None
        self.date_edit.clear()
        self.merchant_edit.clear()
        self.amount_edit.clear()
        self.tax_edit.clear()
        self.discount_edit.clear()
        self.category_box.setCurrentText("その他")
        self.payment_edit.clear()
        self.memo_edit.clear()
        self.item_table.setRowCount(0)
        self.warning_label.setText("元画像と金額を確認してから登録してください。")
        self.save_btn.setText("この1件を記録")
        if clear_image:
            self.current_image = ""
            self.preview.setPixmap(QPixmap())
            self.preview.setText("画像を追加してください")
        self._apply_queue_filter(preserve_active=False)
        self._update_read_controls()

    def _read_image(self):
        if not self.current_image:
            QMessageBox.information(self, "画像がありません", "先にレシート画像を追加してください。")
            return
        self._stash_current_draft()
        path = self.active_queue_path or str(Path(self.current_image).resolve())
        self.current_image = path
        entry = self.entries.get(path)
        if entry and entry.state in ("ready", "saved"):
            answer = QMessageBox.question(self, "もう一度読み取る", "この画像の読み取り結果と修正内容を置き換えますか？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        if entry is None:
            if not Path(path).is_file():
                QMessageBox.warning(self, "画像を開けません", "画像ファイルが見つかりません。画像を追加し直してください。")
                return
            # An image opened from the ledger can also be read again.
            with QSignalBlocker(self.queue):
                self._enqueue_images([path])
            self.active_queue_path = path
        self._begin_batch([path])

    def _read_batch(self):
        self._stash_current_draft()
        self._begin_batch([path for path, entry in self.entries.items() if entry.state == "pending"])

    def _retry_failed(self):
        self._stash_current_draft()
        self._begin_batch([path for path, entry in self.entries.items() if entry.state == "failed"])

    def _begin_batch(self, paths):
        if self.workers:
            return
        if not paths:
            QMessageBox.information(self, "対象の画像がありません", "画像を追加してください。読み取り済みや記録済みの画像は一括読み込みの対象になりません。")
            return
        try:
            model = self.model_box.currentText().strip()
            if not model:
                raise LMStudioError("設定でモデルを選択してください。")
            client = LMStudioClient(normalize_base_url(self.url_edit.text()), self.token_edit.text().strip())
        except LMStudioError as exc:
            QMessageBox.warning(self, "設定を確認してください", str(exc))
            self.tabs.setCurrentIndex(2)
            return
        worker = BatchReadWorker(paths, client, model, self,
                                 source_paths={path: self._resolve_image_path(path) for path in paths},
                                 concurrency=self.settings.concurrent_reads)
        worker._connection_to_save = (normalize_base_url(self.url_edit.text()), model, self.token_edit.text().strip())
        diagnostics.record("batch_started", count=len(paths), concurrency=self.settings.concurrent_reads)
        self.batch_worker = worker
        self.pending_reads = set(paths)
        self.workers.append(worker)
        for path in paths:
            self._refresh_queue_item(path)
        self.batch_progress.setRange(0, len(paths))
        self.batch_progress.setValue(0)
        self.batch_status.setText(f"{len(paths)} 枚の読み取りを開始します。")
        worker.image_started.connect(self._batch_image_started)
        worker.image_succeeded.connect(self._batch_image_succeeded)
        worker.image_failed.connect(self._batch_image_failed)
        worker.progress.connect(self._batch_progress_changed)
        worker.finished.connect(self._batch_finished)
        self._update_read_controls()
        worker.start()

    def _batch_image_started(self, path, index, total):
        entry = self.entries[path]
        entry.state = "reading"
        entry.reviewed = False
        entry.checked = False
        self._refresh_queue_item(path)
        self._update_read_controls()
        if not self.batch_worker or not self.batch_worker.isInterruptionRequested():
            self.batch_status.setText(f"読み取り開始 {index} / {total} 枚: {Path(path).name}（同時解析: 最大 {self.batch_worker.concurrency if self.batch_worker else 1} 件）")

    def _batch_image_succeeded(self, path, result):
        self.pending_reads.discard(path)
        entry = self.entries[path]
        entry.draft = ReceiptDraft.from_extracted(result)
        entry.state = "ready"
        entry.error = ""
        entry.reviewed = False
        entry.checked = False
        entry.saved_to = ""
        self._refresh_queue_item(path)
        if self.active_queue_path == path:
            self._restore_draft(entry.draft)
        self._update_read_controls()
        self._save_session()
        connection = getattr(self.batch_worker, "_connection_to_save", None)
        if connection is not None:
            self.batch_worker._connection_to_save = None
            url, model, token = connection
            try:
                self._persist_connection(url, model, token=token,
                                         models=self.settings.model_choices if url == self.settings.server_url else [])
            except Exception as exc:
                QMessageBox.warning(self, "接続設定を保存できません",
                                    f"レシートは読み取れましたが、接続設定を保存できませんでした。\n{exc}")

    def _batch_image_failed(self, path, message):
        self.pending_reads.discard(path)
        entry = self.entries[path]
        entry.state = "failed"
        entry.error = message
        entry.reviewed = False
        entry.checked = False
        if entry.draft is None:
            entry.draft = ReceiptDraft()
        entry.draft.warning = f"読み取り失敗: {message}\n再試行するか、手入力できます。"
        self._refresh_queue_item(path)
        if self.active_queue_path == path:
            self._restore_draft(entry.draft)
        self._update_read_controls()
        self._save_session()

    def _batch_progress_changed(self, done, total):
        self.batch_progress.setValue(done)

    def _stop_batch(self):
        if self.batch_worker is not None:
            self.batch_worker.requestInterruption()
            self.stop_btn.setEnabled(False)
            self.batch_status.setText("停止を予約しました。送信済みの画像の応答または通信タイムアウトを待って停止します。")

    def _batch_finished(self):
        worker = self.batch_worker
        if worker is None:
            return
        diagnostics.record("batch_finished", count=self.batch_progress.value(),
                           outcome="stopped" if worker.isInterruptionRequested() else "success")
        done = sum(self.entries[path].state == "ready" for path in worker.paths)
        failed = sum(self.entries[path].state == "failed" for path in worker.paths)
        remaining = len(worker.paths) - done - failed
        self.batch_status.setText(f"読み取り完了: {done} 枚 ／ 失敗: {failed} 枚 ／ 未処理: {remaining} 枚。内容を確認してチェックしてください。")
        self.workers.remove(worker)
        self.batch_worker = None
        remaining_paths = self.pending_reads
        self.pending_reads = set()
        for path in remaining_paths:
            self._refresh_queue_item(path)
        worker.deleteLater()
        self._update_read_controls()
        self._save_session()
        if self._close_after_batch:
            QTimer.singleShot(0, self.close)

    def _apply_extraction(self, extracted: ExtractedReceipt):
        self._restore_draft(ReceiptDraft.from_extracted(extracted))
        self.statusBar().showMessage("読み取りが完了しました。内容を確認してください。", 6000)

    def _delete_item_rows(self):
        if self.item_table.selectedIndexes():
            self._form_edited()
        for row in sorted({index.row() for index in self.item_table.selectedIndexes()}, reverse=True):
            self.item_table.removeRow(row)

    def _collect_form(self) -> Receipt:
        existing = next((r for r in self.receipts if r.receipt_id == self.editing_id), None)
        return self._capture_draft().to_receipt(self.current_image, existing)

    def _keep_receipt_image(self, receipt, image_path):
        image_path = self._resolve_image_path(image_path) if image_path else ""
        if image_path and not Path(image_path).is_file():
            return
        if image_path and self.keep_check.isChecked():
            destination, receipt.image_hash = copy_image(image_path, self.settings.image_folder)
            receipt.image_path = str(destination)
        elif image_path:
            receipt.image_path = ""

    def _mark_recorded(self, path, draft, receipt=None):
        entry = self.entries.get(path)
        if entry is not None:
            entry.draft = draft
            entry.state = "saved"
            entry.reviewed = False
            entry.checked = False
            entry.saved_to = str(self.ledger.path)
            if receipt is not None:
                entry.stored_image = receipt.image_path
                entry.image_hash = receipt.image_hash
            self._refresh_queue_item(path)

    def _save_batch(self):
        if self.workers:
            return
        self._stash_current_draft()
        selected = [entry for entry in self.entries.values() if entry.checked and self._can_review(entry)]
        if not selected:
            QMessageBox.information(self, "記録する画像を選択", "読み取り済みのレシートを確認し、一覧の左側にチェックを付けてください。")
            return
        prepared = []
        errors = []
        warnings = []
        for entry in selected:
            try:
                receipt = entry.draft.to_receipt(self._resolve_image_path(entry.path))
                validate_receipt(receipt)
                if duplicate_candidates(receipt, self.receipts + [r for _, r in prepared]):
                    warnings.append(f"{Path(entry.path).name}: 重複候補があります")
                diff = discrepancy(receipt)
                if diff:
                    warnings.append(f"{Path(entry.path).name}: 明細と合計の差 {diff:+,} 円")
                prepared.append((entry, receipt))
            except (LedgerError, OSError) as exc:
                errors.append(f"{Path(entry.path).name}: {exc}")
        if errors:
            QMessageBox.warning(self, "修正が必要です", "まだ保存していません。次の画像を修正するか、チェックを外してください。\n\n" + "\n".join(errors[:10]))
            return
        message = (f"チェックした {len(prepared)} 枚を確認済みとして、税込合計 {sum(r.amount for _, r in prepared):,} 円を記録しますか？\n\n"
                   f"記録先: {self.ledger.path}")
        excluded = sum(entry.checked and entry.state in ("pending", "failed") for entry in self.entries.values())
        if excluded:
            message += f"\n未読・失敗の {excluded} 枚は記録対象外です（チェックは保持します）。"
        hidden_count = sum(self._queue_item(entry.path).isHidden() for entry, _ in prepared)
        if hidden_count:
            message += f"\n一覧で非表示のチェック済みレシート {hidden_count} 件を含みます。"
        known_taxes = [r.tax for _, r in prepared if r.tax is not None]
        message += f"\nうち消費税額（入力済み分）: {sum(known_taxes):,} 円"
        if len(known_taxes) < len(prepared):
            message += f"\n消費税額が未記入: {len(prepared) - len(known_taxes)} 枚"
            missing_tax = [f"・{Path(entry.path).name}（{receipt.merchant} / {receipt.purchased_on:%Y/%m/%d}）"
                           for entry, receipt in prepared if receipt.tax is None]
            message += "\n" + "\n".join(missing_tax[:10])
            if len(missing_tax) > 10:
                message += f"\nほか {len(missing_tax) - 10} 枚（詳細を表示で全件確認できます）"
        if warnings:
            message += "\n\n確認事項:\n" + "\n".join(warnings[:10])
            if len(warnings) > 10:
                message += f"\nほか {len(warnings) - 10} 件"
        if len(prepared) - len(known_taxes) > 10:
            dialog = QMessageBox(self)
            dialog.setWindowTitle("まとめてExcelに記録")
            dialog.setIcon(QMessageBox.Icon.Question)
            dialog.setTextFormat(Qt.TextFormat.PlainText)
            dialog.setText(message)
            dialog.setDetailedText("消費税額が未記入の画像（全件）\n" + "\n".join(missing_tax))
            dialog.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            dialog.setDefaultButton(QMessageBox.StandardButton.No)
            answer = dialog.exec()
        else:
            answer = QMessageBox.question(self, "まとめてExcelに記録", message,
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                          QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            for entry, receipt in prepared:
                self._keep_receipt_image(receipt, entry.path)
            updated = self.receipts + [r for _, r in prepared]
            self._stage_commit([(entry.path, receipt) for entry, receipt in prepared])
            self.ledger.save(updated)
            self.receipts = updated
            for entry, receipt in prepared:
                self._mark_recorded(entry.path, entry.draft, receipt)
            self._refresh_months()
            self._refresh_ledger()
            if self.active_queue_path and self.entries[self.active_queue_path].state == "saved":
                self._new_form()
            self.batch_status.setText(f"{len(prepared)} 枚をExcelに記録しました。")
            self.statusBar().showMessage(f"一括記録が完了しました: {len(prepared)} 枚", 7000)
            self._pending_commit = []
            self._save_session()
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "一括記録できません", f"読み取り結果を保持しています。原因を解消して再度記録してください。\n{exc}")

    def _save_receipt(self):
        if self.workers:
            return
        try:
            receipt = self._collect_form()
            if not receipt.merchant:
                raise LedgerError("店舗名を入力してください。")
            duplicate = duplicate_candidates(receipt, self.receipts)
            if duplicate:
                answer = QMessageBox.question(self, "重複候補", "同じ画像または日付・店舗・金額の支出があります。登録しますか？")
                if answer != QMessageBox.StandardButton.Yes:
                    return
            diff = discrepancy(receipt)
            if diff:
                answer = QMessageBox.question(self, "金額の差", f"明細の合計と支払合計に {diff:+,} 円の差があります。確認して登録しますか？")
                if answer != QMessageBox.StandardButton.Yes:
                    return
            self._keep_receipt_image(receipt, self.current_image)
            updated = [r for r in self.receipts if r.receipt_id != receipt.receipt_id] + [receipt]
            self._stage_commit([(self.active_queue_path, receipt)])
            self.ledger.save(updated)
            self.receipts = updated
            self._mark_recorded(self.active_queue_path or self.current_image, self._capture_draft(), receipt)
            self._refresh_months()
            self._refresh_ledger()
            self.statusBar().showMessage(f"テンプレート形式で記録しました: {self.ledger.path.name}", 7000)
            self._new_form()
            self._pending_commit = []
            self._save_session()
        except WorkbookChanged as exc:
            QMessageBox.warning(self, "家計簿が変更されました", str(exc))
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "登録できません", str(exc))

    def _load_ledger(self):
        try:
            self.receipts = self.ledger.load()
            self._refresh_months()
            self._refresh_ledger()
            self.statusBar().showMessage(f"家計簿を読み込みました: {len(self.receipts)} 件", 5000)
        except LedgerError as exc:
            QMessageBox.warning(self, "家計簿を開けません", str(exc))
        finally:
            self._refresh_workbook_label()

    def _refresh_months(self):
        current = self.month_box.currentText()
        months = sorted({r.purchased_on.strftime("%Y-%m") for r in self.receipts}, reverse=True)
        self.month_box.blockSignals(True)
        self.month_box.clear()
        self.month_box.addItem("すべての月")
        self.month_box.addItems(months)
        self.month_box.setCurrentText(current if current in months else "すべての月")
        self.month_box.blockSignals(False)

    def _refresh_ledger(self):
        self._refresh_workbook_label()
        query = self.search_edit.text().casefold().strip()
        month = self.month_box.currentText()
        records = [r for r in self.receipts if
                   (month == "すべての月" or r.purchased_on.strftime("%Y-%m") == month)
                   and (not query or query in (r.merchant + r.category + r.memo).casefold())]
        records.sort(key=lambda r: (r.purchased_on, r.updated_at), reverse=True)
        self.ledger_table.setRowCount(len(records))
        for row, receipt in enumerate(records):
            values = (receipt.purchased_on.isoformat(), receipt.merchant,
                      f"{receipt.amount:,} 円", str(len(receipt.items)), receipt.category,
                      receipt.payment_method, receipt.receipt_id)
            for col, value in enumerate(values):
                self.ledger_table.setItem(row, col, QTableWidgetItem(value))
        self.total_label.setText(f"表示中の合計: {sum(r.amount for r in records):,} 円  ／  {len(records)} 件")
        by_category = {}
        for receipt in records:
            by_category[receipt.category] = by_category.get(receipt.category, 0) + receipt.amount
        self.category_summary_label.setText("費目別: " + "  ／  ".join(
            f"{name} {amount:,} 円" for name, amount in sorted(by_category.items())
        ) if by_category else "費目別: 登録なし")

    def _selected_receipt(self) -> Receipt | None:
        row = self.ledger_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "選択してください", "家計簿から1件選択してください。")
            return None
        receipt_id = self.ledger_table.item(row, 6).text()
        return next((r for r in self.receipts if r.receipt_id == receipt_id), None)

    def _show_selected_details(self):
        receipt = self._selected_receipt()
        if not receipt:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{receipt.merchant} のレシート明細")
        dialog.resize(920, 560)
        layout = QVBoxLayout(dialog)
        title = QLabel(f"{receipt.purchased_on:%Y/%m/%d}　{receipt.merchant}　税込合計 {receipt.amount:,} 円")
        title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 8px;")
        layout.addWidget(title)
        tax_text = f"{receipt.tax:,} 円" if receipt.tax is not None else "未記入"
        meta = QLabel(f"費目: {receipt.category}　支払方法: {receipt.payment_method or '未記入'}　うち消費税額: {tax_text}　値引き額: {receipt.discount if receipt.discount is not None else '未記入'}")
        meta.setWordWrap(True)
        layout.addWidget(meta)
        if receipt.memo:
            memo = QLabel(f"レシートのメモ: {receipt.memo}")
            memo.setWordWrap(True)
            layout.addWidget(memo)
        table = QTableWidget(len(receipt.items), 5)
        table.setHorizontalHeaderLabels(["商品名", "個数", "単価（円）", "金額（円）", "詳細・補足"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        for row, item in enumerate(receipt.items):
            values = (item.name, item.quantity, item.unit_price, item.amount, item.note)
            for col, value in enumerate(values):
                table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))
        layout.addWidget(table)
        difference = discrepancy(receipt)
        if difference is not None and difference != 0:
            layout.addWidget(QLabel(f"明細金額と合計の差: {difference:,} 円（税・値引き等を確認）"))
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()

    def _edit_selected(self):
        receipt = self._selected_receipt()
        if not receipt:
            return
        self._new_form()
        self.editing_id = receipt.receipt_id
        self.date_edit.setText(receipt.purchased_on.isoformat())
        self.merchant_edit.setText(receipt.merchant)
        self.amount_edit.setText(str(receipt.amount))
        self.tax_edit.setText("" if receipt.tax is None else str(receipt.tax))
        self.discount_edit.setText("" if receipt.discount is None else str(receipt.discount))
        self.category_box.setCurrentText(receipt.category)
        self.payment_edit.setText(receipt.payment_method)
        self.memo_edit.setText(receipt.memo)
        for item in receipt.items:
            row = self.item_table.rowCount()
            self.item_table.insertRow(row)
            for col, value in enumerate((item.name, item.quantity, item.unit_price, item.amount, item.note)):
                self.item_table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))
        if receipt.image_path:
            self._set_image(receipt.image_path)
        self.save_btn.setText("変更を保存")
        self.tabs.setCurrentIndex(0)

    def _delete_selected(self):
        receipt = self._selected_receipt()
        if not receipt:
            return
        answer = QMessageBox.question(self, "削除の確認", f"{receipt.merchant} の {receipt.amount:,} 円を削除しますか？")
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            updated = [r for r in self.receipts if r.receipt_id != receipt.receipt_id]
            self.ledger.save(updated)
            self.receipts = updated
            self._refresh_months()
            self._refresh_ledger()
        except LedgerError as exc:
            QMessageBox.warning(self, "削除できません", str(exc))

    def _open_workbook(self):
        path = self.ledger.path
        if not path.exists():
            QMessageBox.information(self, "家計簿がありません", "「Excelを保存」または「確認してExcelに記録」でテンプレート形式のファイルを作成します。")
            return
        try:
            os.startfile(path)
        except OSError as exc:
            QMessageBox.warning(self, "Excelを開けません", str(exc))

    def _refresh_workbook_label(self):
        path = self.ledger.path
        state = "保存済み" if path.exists() else "次の保存時に作成"
        self.workbook_label.setText(f"現在の保存先: {path}\n{state} ／ 登録済み {len(self.receipts)} 件")
        name = path.name if len(path.name) <= 35 else path.name[:32] + "…"
        self.receipt_save_target.setText(f"記録先: {name}")
        self.receipt_save_target.setToolTip(str(path))
        self.delete_excel_btn.setEnabled(path.is_file() and self.ledger.fingerprint is not None)
        if self.ledger.maintenance_warning:
            self.statusBar().showMessage(self.ledger.maintenance_warning, 15000)

    def _activate_workbook(self, ledger: ExcelLedger, records: list[Receipt], *, keep_editing=False):
        self.ledger = ledger
        self.receipts = records
        self.settings.workbook_path = str(ledger.path)
        self.workbook_edit.setText(str(ledger.path))
        # Retain a just-read receipt when the user chooses where to record it.
        if not keep_editing:
            self.editing_id = None
            self.save_btn.setText("この1件を記録")
        self.search_edit.clear()
        self.month_box.setCurrentIndex(0)
        self._refresh_months()
        self._refresh_ledger()
        try:
            self.settings.save()
        except OSError as exc:
            QMessageBox.warning(self, "保存先設定を保持できません", f"この実行中の保存先は変更しましたが、次回起動用の設定を保存できません。\n{exc}")

    def _select_workbook(self):
        path, _ = QFileDialog.getOpenFileName(self, "記録先の家計簿またはテンプレートを選択", str(self.ledger.path.parent), "Excel (*.xlsx)")
        if not path:
            return
        try:
            candidate = self._ledger_for(Path(path).resolve())
            records = candidate.load()
            keep_editing = (candidate.path == self.ledger.path.resolve()
                            and any(r.receipt_id == self.editing_id for r in records))
            self._activate_workbook(candidate, records, keep_editing=keep_editing)
            self.statusBar().showMessage("記録先を変更しました。登録画面の内容は、このExcelへ追記できます。", 7000)
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "家計簿を選択できません", str(exc))

    def _save_workbook(self):
        try:
            self.ledger.save(self.receipts)
            self._refresh_workbook_label()
            self.statusBar().showMessage(f"登録済み {len(self.receipts)} 件を保存しました: {self.ledger.path}", 7000)
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "Excelを保存できません", str(exc))

    def _save_workbook_as(self):
        source = self.ledger.path
        destination, _ = QFileDialog.getSaveFileName(self, "登録済みの家計簿に名前を付けて保存", str(source.with_name(source.stem + "-copy.xlsx")), "Excel (*.xlsx)")
        if not destination:
            return
        try:
            path = Path(destination).resolve()
            if path.suffix.lower() != ".xlsx":
                raise LedgerError("ファイル名の末尾は .xlsx にしてください。")
            if path == source.resolve():
                self._save_workbook()
                return
            candidate = self._ledger_for(path)
            candidate.load()  # Validate an existing destination before replacing it.
            candidate.save(self.receipts)
            self._activate_workbook(candidate, list(self.receipts), keep_editing=True)
            self.statusBar().showMessage(f"保存しました。今後の記録先: {path}", 7000)
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "Excelを保存できません", str(exc))

    def _delete_workbook(self):
        if not self.ledger.path.is_file():
            QMessageBox.information(self, "ファイルがありません", "削除する Excel ファイルがありません。")
            return
        answer = QMessageBox.question(
            self, "Excelファイルを削除",
            f"この家計簿ファイルと、登録済みの {len(self.receipts)} 件を一覧から削除しますか？\n\n"
            f"{self.ledger.path}\n\n同じフォルダーに復元用バックアップを残します。次の保存時に空の家計簿から再開します。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            backup = self.ledger.delete()
            self.receipts = []
            self.editing_id = None
            self.save_btn.setText("この1件を記録")
            self._refresh_months()
            self._refresh_ledger()
            QMessageBox.information(self, "削除しました", f"家計簿ファイルを削除しました。\n復元用バックアップ: {backup}\n\n設定の「バックアップから復元」で戻せます。")
        except LedgerError as exc:
            QMessageBox.warning(self, "Excelを削除できません", str(exc))

    def _restore_backup(self):
        source, _ = QFileDialog.getOpenFileName(self, "復元するバックアップを選択", str(self.ledger.path.parent), "Excel (*.xlsx)")
        if not source:
            return
        if Path(source).resolve() == self.ledger.path.resolve():
            QMessageBox.information(self, "別のファイルを選択", "現在の家計簿とは別のバックアップを選択してください。")
            return
        answer = QMessageBox.question(self, "復元の確認", "現在の家計簿をバックアップしてから、選択したファイルに置き換えますか？")
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            candidate = self._ledger_for(source)
            candidate.load()  # Validate before replacing the current workbook.
            current = self.ledger.path
            current.parent.mkdir(parents=True, exist_ok=True)
            if current.exists():
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
                shutil.copy2(current, current.with_name(f"{current.stem}.before-restore-{stamp}.xlsx"))
            copy_workbook_rebased(source, current)
            self._load_ledger()
            self.statusBar().showMessage("バックアップから復元しました。", 6000)
        except (LedgerError, OSError) as exc:
            QMessageBox.warning(self, "復元できません", str(exc))

    def _choose_workbook(self):
        path, _ = QFileDialog.getSaveFileName(self, "家計簿ファイルを選択", self.workbook_edit.text(), "Excel (*.xlsx)")
        if path:
            self.workbook_edit.setText(path)

    def _choose_image_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "記録済み画像の保存先", self.image_folder_edit.text())
        if folder:
            self.image_folder_edit.setText(folder)

    def _open_folder(self, folder):
        try:
            os.startfile(writable_directory(folder))
        except OSError as exc:
            QMessageBox.warning(self, "フォルダを開けません", str(exc))

    def _refresh_storage_usage(self):
        if self.workers:
            return
        paths = [receipt.image_path for receipt in self.receipts if receipt.image_path]
        paths += [entry.stored_image for entry in self.entries.values() if entry.stored_image]
        for batch in self.removed_batches:
            paths += [item["entry"].get("stored_image", "") for item in batch["entries"] if item["entry"].get("stored_image")]
        image_folder, workbook = self.settings.image_folder, self.ledger.path
        def measure():
            images = file_usage([*paths, *(path for path in folder_files(image_folder) if supported_image(path))])
            normal = file_usage(regular_backups(workbook))
            safety = file_usage([*workbook.parent.glob(workbook.stem + ".before-delete-*.xlsx"),
                                 *workbook.parent.glob(workbook.stem + ".before-restore-*.xlsx")])
            bundles = file_usage([* (default_data_dir() / "backups").glob("*.zip"),
                                 * (app_data_dir() / "backups").glob("*.zip")])
            state = file_usage([app_data_dir() / "settings.json", app_data_dir() / "draft-session.json"])
            book = file_usage([workbook])
            return (f"現在のExcel: {size_text(book[1])}\n"
                    f"保存済み画像（設定フォルダ・参照先）: {images[0]} 件 / {size_text(images[1])}\n"
                    f"この家計簿の通常バックアップ: {normal[0]} 件 / {size_text(normal[1])}\n"
                    f"削除前・復元前の退避: {safety[0]} 件 / {size_text(safety[1])}\n"
                    f"一括ZIP（標準バックアップフォルダ）: {bundles[0]} 件 / {size_text(bundles[1])}\n"
                    f"設定・下書き・削除履歴: {size_text(state[1])}")
        self._start_task(measure, self.storage_usage_label.setText, "保存容量を確認しています…")

    def _prune_backups_now(self):
        if self.workers:
            return
        try:
            keep = self.backup_limit_spin.value()
            targets = regular_backups(self.ledger.path)[keep:]
            if not targets:
                QMessageBox.information(self, "整理するファイルはありません", f"通常バックアップは {keep} 世代以内です。")
                return
            if QMessageBox.question(self, "古いバックアップを整理", f"この家計簿の通常バックアップを新しい順に {keep} 世代残し、古い {len(targets)} 件を削除します。続けますか？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
            prune_backups(self.ledger.path, keep)
            self._refresh_storage_usage()
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "整理できません", str(exc))

    def _export_bundle(self):
        if self.workers or not self._save_session():
            return
        filename = "ReceiptLedger-data-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".zip"
        destination, _ = QFileDialog.getSaveFileName(self, "データ一式をバックアップ", str(default_data_dir() / "backups" / filename), "ZIP (*.zip)")
        if not destination:
            return
        snapshot, settings = self._session_snapshot(), replace(self.settings, workbook_path=str(self.ledger.path))
        self._data_busy = True
        def done(result):
            message = f"バックアップを保存しました。\n{result['path']}\n{result['files']} ファイル。認証トークンは含みません。"
            if result["missing"]:
                message += f"\n\n元ファイルが見つからない {len(result['missing'])} 件は含まれていません:\n" + "\n".join(result["missing"][:8])
            QMessageBox.information(self, "バックアップ完了", message)
        self._start_task(lambda: create_bundle(destination, settings, snapshot), done, "データ一式をバックアップしています…")

    def _import_bundle(self):
        if self.workers:
            return
        source, _ = QFileDialog.getOpenFileName(self, "復元する一括バックアップ", str(default_data_dir() / "backups"), "ZIP (*.zip)")
        if not source:
            return
        parent = QFileDialog.getExistingDirectory(self, "復元先（この中に新しいフォルダを作成します）", str(default_data_dir()))
        if not parent:
            return
        message = ("現在のデータ一式を退避してから、選択したZIPを新しいフォルダへ復元し、保存先・設定・下書きを切り替えます。\n"
                   "元のExcelと画像はそのまま残ります。認証トークンは変更しません。続けますか？")
        if QMessageBox.question(self, "データ一式を復元", message,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes or not self._save_session():
            return
        snapshot, settings = self._session_snapshot(), replace(self.settings, workbook_path=str(self.ledger.path))
        safety = app_data_dir() / "backups" / ("before-data-restore-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f") + ".zip")
        def task():
            create_bundle(safety, settings, snapshot)
            return restore_bundle(source, parent)
        def done(result):
            try:
                activate_restored_state(result["settings"], result["session"])
                self._restoring_session = True
                self.settings = Settings.load()
                self.entries.clear()
                self.removed_batches.clear()
                self.active_queue_path = None
                with QSignalBlocker(self.queue):
                    self.queue.clear()
                self._new_form()
                self.ledger = self._ledger_for(self.settings.workbook_path)
                self._load_ledger()
                self._pending_commit = []
                self._session_notice = ""
                self._restore_session()
                self.image_folder_edit.setText(self.settings.image_folder)
                self.workbook_edit.setText(self.settings.workbook_path)
                self.backup_limit_spin.setValue(self.settings.backup_generations)
                self.keep_check.setChecked(self.settings.keep_images)
                self.url_edit.setText(self.settings.server_url)
                with QSignalBlocker(self.concurrency_box):
                    self.concurrency_box.setCurrentIndex(self.concurrency_box.findData(self.settings.concurrent_reads))
                with QSignalBlocker(self.model_box):
                    self.model_box.clear()
                    self.model_box.addItems(self.settings.model_choices)
                    self.model_box.setCurrentText(self.settings.model)
                for selector in (self.theme_box, self.theme_selector):
                    with QSignalBlocker(selector):
                        selector.setCurrentIndex(selector.findData(self.settings.theme))
                self._apply_theme()
                self.storage_usage_label.setText("復元しました。「容量を確認」で再集計できます。")
                self._restoring_session = False
                message = f"新しい保存先へ復元しました。\n{result['folder']}\n\n復元前の退避: {safety}"
                if result["missing"]:
                    message += f"\n元のバックアップに含まれていないファイル: {len(result['missing'])} 件"
                QMessageBox.information(self, "復元完了", message)
            except (OSError, ValueError, LedgerError) as exc:
                self._session_write_blocked = True
                QMessageBox.warning(self, "復元の切り替えを完了できません", f"復元先と退避データは保持しています。再起動時に切り替えを再試行します。\n{exc}")
        self._data_busy = True
        self._start_task(task, done, "データを検証して復元しています…")

    def _show_update_guide(self):
        QMessageBox.information(self, f"更新手順 — 現在 v{APP_VERSION}",
            "1. 設定の「データ一式をバックアップ」でZIPを保存します。\n"
            "2. アプリを終了し、新版ZIPを新しいフォルダへ展開します。\n"
            "3. 実行ファイル横の data フォルダを使っている場合は、新版の実行ファイル横へ data をコピーします。外部の保存先は引き継ぎます。\n"
            "4. 新版を起動し、設定のバージョン・Excel保存先・画像保存先と記録を確認します。\n\n"
            "ReceiptLedger.exe と _internal は同じ配布版の組み合わせで使ってください。新しいPCでは一括バックアップから復元し、LM Studioの認証トークンを設定してください。")

    def _test_connection(self):
        if self.workers:
            return
        try:
            token = self.token_edit.text().strip()
            client = LMStudioClient(self.url_edit.text(), token, timeout=20)
        except LMStudioError as exc:
            QMessageBox.warning(self, "接続先を確認してください", str(exc))
            return
        selected_model = self.model_box.currentText().strip()
        self._start_task(client.models,
                         lambda models: self._show_models(models, url=client.base_url, token=token, selected_model=selected_model),
                         "LM Studio に接続しています…")

    def _persist_connection(self, url, model, *, models=None, token=None):
        candidate = replace(self.settings, server_url=url, model=model,
                            model_choices=list(models) if models is not None else list(self.settings.model_choices))
        self._commit_settings(candidate, token)

    def _commit_settings(self, candidate, token=None):
        changed_token = token is not None and token != self._saved_server_token
        if changed_token:
            save_token(token)
        try:
            candidate.save()
        except Exception:
            if changed_token:
                # Do not leave a new server's credential paired with the previous URL.
                save_token(self._saved_server_token)
            raise
        self.settings = candidate
        if token is not None:
            self._saved_server_token = token

    def _show_models(self, models: list[str], *, url=None, token=None, selected_model=None):
        url = url if url is not None else normalize_base_url(self.url_edit.text())
        token = token if token is not None else self.token_edit.text().strip()
        current = selected_model if selected_model is not None else self.model_box.currentText().strip()
        models = list(dict.fromkeys(m for m in models if isinstance(m, str)))
        selected = current if current and (url == self.settings.server_url or current in models) else (models[0] if models else current)
        with QSignalBlocker(self.model_box):
            self.model_box.clear()
            self.model_box.addItems(models)
            self.model_box.setCurrentText(selected)
        try:
            self._persist_connection(url, selected, models=models, token=token)
        except Exception as exc:
            QMessageBox.warning(self, "接続設定を保存できません",
                                f"接続は成功しましたが、設定を保存できませんでした。設定を保存してから終了してください。\n{exc}")
            return
        self.url_edit.setText(url)
        QMessageBox.information(self, "接続成功", f"LM Studio に接続しました。利用可能なモデル: {len(models)} 件\n接続設定を保存しました。次回起動時に復元します。")

    def _save_model_selection(self, *_):
        model = self.model_box.currentText().strip()
        if not model or model == self.settings.model:
            return
        try:
            url = normalize_base_url(self.url_edit.text())
            # An untested edited URL must not replace the last working server.
            if url != self.settings.server_url:
                return
            self._persist_connection(url, model)
            self.statusBar().showMessage("モデルの選択を保存しました。", 5000)
        except Exception as exc:
            QMessageBox.warning(self, "モデルを保存できません", str(exc))

    def _save_concurrency(self, *_):
        previous = self.settings.concurrent_reads
        try:
            candidate = replace(self.settings, concurrent_reads=self.concurrency_box.currentData())
            candidate.save()
            self.settings = candidate
            self.statusBar().showMessage("同時解析数を保存しました。次の読み取りから適用します。", 5000)
        except Exception as exc:
            with QSignalBlocker(self.concurrency_box):
                self.concurrency_box.setCurrentIndex(self.concurrency_box.findData(previous))
            QMessageBox.warning(self, "同時解析数を保存できません", str(exc))

    def _save_settings(self):
        if self.workers:
            return
        try:
            url = normalize_base_url(self.url_edit.text())
            workbook = Path(self.workbook_edit.text().strip()).expanduser().resolve()
            if workbook.suffix.lower() != ".xlsx":
                raise LedgerError("家計簿ファイルは .xlsx を指定してください。")
            if not self.image_folder_edit.text().strip():
                raise LedgerError("画像の保存先フォルダを指定してください。")
            image_folder = writable_directory(self.image_folder_edit.text().strip())
            writable_directory(workbook.parent)
            changed_workbook = workbook != self.ledger.path
            next_ledger = self._ledger_for(workbook) if changed_workbook else self.ledger
            next_receipts = next_ledger.load() if changed_workbook else self.receipts
            candidate = replace(self.settings, server_url=url, model=self.model_box.currentText().strip(),
                                model_choices=self.settings.model_choices if url == self.settings.server_url else [],
                                workbook_path=str(workbook), image_folder=str(image_folder),
                                keep_images=self.keep_check.isChecked(), backup_generations=self.backup_limit_spin.value())
            self._commit_settings(candidate, self.token_edit.text().strip())
            self.ledger, self.receipts = next_ledger, next_receipts
            self.ledger.backup_limit = candidate.backup_generations
            if changed_workbook:
                self.editing_id = None
                self.save_btn.setText("この1件を記録")
            self._refresh_months()
            self._refresh_ledger()
            self.image_folder_edit.setText(str(image_folder))
            self._save_session()
            self.statusBar().showMessage("設定を保存しました。", 5000)
        except Exception as exc:
            QMessageBox.warning(self, "設定を保存できません", str(exc))

    def closeEvent(self, event):
        saved = self._save_session()
        if self.workers:
            if self.batch_worker:
                self._close_after_batch = True
                self._stop_batch()
                self.batch_status.setText("下書きを保存しました。送信済みの画像の処理がすべて終わったら終了します。")
            else:
                QMessageBox.information(self, "処理中", "処理の完了を待ってください。")
            event.ignore()
        else:
            if not saved:
                answer = QMessageBox.question(self, "下書きを保存できません", "最新の入力を復元できない可能性があります。終了しますか？",
                                              QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                              QMessageBox.StandardButton.No)
                if answer != QMessageBox.StandardButton.Yes:
                    event.ignore()
                    return
            self.autosave_timer.stop()
            self.theme_timer.stop()
            QApplication.instance().removeEventFilter(self.drop_filter)
            event.accept()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(url.isLocalFile() and supported_image(url.toLocalFile()) for url in event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        self._enqueue_images([url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()])
        event.acceptProposedAction()


def run():
    app = QApplication([])
    app.setApplicationName("レシート家計簿")
    app.setApplicationVersion(APP_VERSION)
    app.setWindowIcon(QIcon(str(Path(__file__).resolve().parent / "assets/receipt-ledger.ico")))
    folder = app_data_dir()
    folder.mkdir(parents=True, exist_ok=True)
    lock = QLockFile(str(folder / "app.lock"))
    lock.setStaleLockTime(0)
    if not lock.tryLock(0):
        QMessageBox.information(None, "起動できません", "レシート家計簿が起動済みか、保存フォルダーを利用できません。起動中の画面をご確認ください。")
        return 1
    try:
        window = MainWindow()
    except (OSError, ValueError, LedgerError) as exc:
        QMessageBox.warning(None, "起動できません", f"保存フォルダまたは復元途中のデータを確認してください。\n{exc}")
        return 1
    window.setAcceptDrops(True)
    window.show()
    result = app.exec()
    lock.unlock()
    return result
