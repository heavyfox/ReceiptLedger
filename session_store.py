"""Atomic local recovery snapshots, separate from the Excel ledger."""
import json
import os
import shutil
import tempfile
from dataclasses import asdict
from datetime import datetime
from config import map_session_paths, portable_path, resolve_config_path


def receipt_signature(receipt):
    return {"date": receipt.purchased_on.isoformat(), "merchant": receipt.merchant,
            "amount": receipt.amount, "tax": receipt.tax, "discount": receipt.discount,
            "category": receipt.category, "payment": receipt.payment_method,
            "memo": receipt.memo, "items": [asdict(item) for item in receipt.items]}


class SessionStore:
    def __init__(self, folder):
        self.path = folder / "draft-session.json"
        self.last_text = None

    def load(self):
        if not self.path.exists():
            return None
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or data.get("version") != 1:
                raise ValueError("下書きの形式が対応していません")
            if not isinstance(data.get("entries"), list) or not isinstance(data.get("workbook"), str):
                raise ValueError("下書きの形式が正しくありません")
            if data.pop("paths_base", None) == "application":
                data = map_session_paths(data, resolve_config_path)
            return data
        except (ValueError, TypeError):
            raise ValueError("下書きの内容を解釈できませんでした。")

    def preserve_invalid(self):
        backup = self.path.with_name(f"draft-session.invalid-{datetime.now():%Y%m%d-%H%M%S-%f}.json")
        shutil.copy2(self.path, backup)

    def save(self, data):
        data = map_session_paths(data, portable_path)
        data["paths_base"] = "application"
        text = json.dumps(data, ensure_ascii=False, sort_keys=True)
        if text == self.last_text:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix=".draft-", suffix=".json", dir=self.path.parent)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                handle.write(text)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.path)
            self.last_text = text
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
