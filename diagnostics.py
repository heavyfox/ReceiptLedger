"""Bounded local diagnostics. Only explicitly allowed metadata is recorded."""
from __future__ import annotations

import json
import platform
import re
import threading
from datetime import datetime, timezone
from pathlib import Path

from config import app_data_dir
from version import APP_VERSION

MAX_BYTES = 512 * 1024
BACKUP_COUNT = 3
_lock = threading.RLock()
EVENTS = frozenset({"app_started", "task_failed", "extract_started", "extract_finished",
                    "extract_retry", "request_finished", "excel_saved", "excel_failed",
                    "batch_started", "batch_finished", "session_failed"})
ERRORS = frozenset({"unknown", "none", "image_decode", "timeout", "network", "auth",
                    "http_error", "redirect", "invalid_response", "invalid_receipt",
                    "output_limit", "refusal", "missing_model", "permission", "io",
                    "workbook_changed", "validation"})
FORMATS = frozenset({".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif", "other"})
NUMBERS = frozenset({"elapsed_ms", "http_status", "attempt", "count", "concurrency"})


def folder() -> Path:
    return app_data_dir() / "logs"


def _metadata(fields: dict) -> dict:
    """Reject unknown fields and free text, even when supplied by future callers."""
    clean = {}
    for key, value in fields.items():
        if key in NUMBERS and type(value) is int and 0 <= value <= 10**12:
            clean[key] = value
        elif key == "format" and isinstance(value, str) and value in FORMATS:
            clean[key] = value
        elif key == "error" and isinstance(value, str) and value in ERRORS:
            clean[key] = value
        elif key == "endpoint" and value in ("models", "extract"):
            clean[key] = value
        elif key == "outcome" and value in ("success", "failed", "stopped"):
            clean[key] = value
        elif key == "job" and isinstance(value, str) and len(value) == 32 and all(c in "0123456789abcdef" for c in value):
            clean[key] = value
    return clean


def error_code(exc: Exception) -> str:
    code = getattr(exc, "code", None)
    if isinstance(code, str) and code in ERRORS:
        return code
    if isinstance(exc, PermissionError):
        return "permission"
    if isinstance(exc, TimeoutError):
        return "timeout"
    if isinstance(exc, OSError):
        return "io"
    if isinstance(exc, ValueError):
        return "validation"
    # Avoid importing core (and Qt) into this small module.
    if type(exc).__name__ == "WorkbookChanged":
        return "workbook_changed"
    cause = getattr(exc, "__cause__", None)
    if isinstance(cause, Exception) and cause is not exc:
        if isinstance(cause, PermissionError):
            return "permission"
        if isinstance(cause, OSError):
            return "io"
        if isinstance(cause, ValueError):
            return "validation"
    return "unknown"


def record(event: str, **fields) -> None:
    """Logging failure must never interrupt receipt processing or saving."""
    if not isinstance(event, str) or event not in EVENTS:
        return
    row = {"time": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
           "event": event, **_metadata(fields)}
    data = (json.dumps(row, ensure_ascii=True) + "\n").encode("utf-8")
    try:
        with _lock:
            directory = folder()
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / "diagnostics.jsonl"
            if path.exists() and path.stat().st_size + len(data) > MAX_BYTES:
                oldest = directory / f"diagnostics.jsonl.{BACKUP_COUNT}"
                oldest.unlink(missing_ok=True)
                for index in range(BACKUP_COUNT - 1, 0, -1):
                    source = directory / f"diagnostics.jsonl.{index}"
                    if source.exists():
                        source.replace(directory / f"diagnostics.jsonl.{index + 1}")
                path.replace(directory / "diagnostics.jsonl.1")
            with path.open("ab") as stream:
                stream.write(data)
    except OSError:
        pass


def report(*, count: int = 0, concurrency: int = 1) -> str:
    """Copyable metadata plus the latest 50 events; no settings or receipt text."""
    header = (f"ReceiptLedger {APP_VERSION}\nPython {platform.python_version()} / "
              f"{platform.system()} {platform.machine()}\n"
              f"画像一覧: {max(0, int(count))} 件 / 同時解析: {int(concurrency)} 件\n"
              "URL・トークン・モデル名・画像・ファイル名・レシート本文は含みません。\n")
    rows = []
    try:
        with _lock:
            path = folder() / "diagnostics.jsonl"
            with path.open("rb") as stream:
                stream.seek(max(0, path.stat().st_size - 32 * 1024))
                lines = stream.read(32 * 1024).decode("utf-8", errors="replace").splitlines()
        for line in lines[-50:]:
            try:
                row = json.loads(line)
                if not isinstance(row, dict) or row.get("event") not in EVENTS:
                    continue
                # Re-sanitize even edited log files; accept only a canonical UTC timestamp.
                clean = {"event": row["event"], **_metadata(row)}
                timestamp = row.get("time")
                if isinstance(timestamp, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}\+00:00", timestamp):
                    clean["time"] = timestamp
                rows.append(json.dumps(clean, ensure_ascii=False))
            except (ValueError, TypeError):
                continue
    except OSError:
        pass
    legend = ("\n原因: timeout=通信時間切れ / network=接続失敗 / auth=認証失敗 / "
              "http_error=サーバーエラー / image_decode=画像変換失敗 / "
              "invalid_response=応答のJSON不正 / invalid_receipt=抽出結果の不正 / "
              "output_limit=モデル出力上限 / validation=保存内容の検証失敗\n")
    return header + legend + "\n直近の診断イベント:\n" + ("\n".join(rows) or "まだありません。")
