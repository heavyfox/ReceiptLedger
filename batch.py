"""Per-image receipt drafts and sequential local-server extraction."""

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from core import CATEGORIES, LedgerError, LineItem, Receipt, image_digest
from lm_client import ExtractedReceipt


def parse_yen(text: str, name: str, required=False) -> int | None:
    value = text.strip().replace(",", "").replace("¥", "").replace("￥", "")
    if not value:
        if required:
            raise LedgerError(f"{name}を入力してください。")
        return None
    try:
        return int(value)
    except ValueError as exc:
        raise LedgerError(f"{name}は円の整数で入力してください。") from exc


@dataclass
class ReceiptDraft:
    purchased_on: str = ""
    merchant: str = ""
    amount: str = ""
    tax: str = ""
    discount: str = ""
    category: str = "その他"
    payment: str = ""
    memo: str = ""
    items: list[list[str]] = field(default_factory=list)
    warning: str = "元画像と金額を確認してから登録してください。"

    @classmethod
    def from_extracted(cls, result: ExtractedReceipt):
        def text(value):
            return "" if value is None else str(value)

        warnings = ["元画像と抽出結果を確認してください。", *result.uncertain_fields]
        if not result.receipt_date or result.total_yen is None:
            warnings.append("日付または合計額が不明です。入力してから登録してください。")
        if result.tax_yen is None:
            warnings.append("消費税額が不明です。画像で確認できる場合は入力してください（空欄のままでも記録できます）。")
        return cls(
            purchased_on=result.receipt_date or "", merchant=result.merchant or "",
            amount=text(result.total_yen), tax=text(result.tax_yen), discount=text(result.discount_yen),
            category=result.category if result.category in CATEGORIES else "その他",
            payment=result.payment_method or "",
            items=[[text(v) for v in (i.name, i.quantity, i.unit_price_yen, i.line_total_yen, i.note)]
                   for i in result.items], warning="\n".join(warnings),
        )

    def has_content(self):
        return bool(self.purchased_on or self.merchant or self.amount or self.memo
                    or self.tax or self.discount or self.payment or self.category != "その他"
                    or any(any(row) for row in self.items))

    def to_receipt(self, image_path="", existing: Receipt | None = None) -> Receipt:
        try:
            purchased = date.fromisoformat(self.purchased_on.strip())
        except ValueError as exc:
            raise LedgerError("購入日を YYYY-MM-DD 形式で入力してください。") from exc
        receipt = Receipt(
            receipt_id=existing.receipt_id if existing else Receipt().receipt_id,
            purchased_on=purchased, merchant=self.merchant.strip(),
            amount=parse_yen(self.amount, "合計金額", required=True), category=self.category,
            tax=parse_yen(self.tax, "税額"), discount=parse_yen(self.discount, "値引き額"),
            payment_method=self.payment.strip(), memo=self.memo.strip(),
            created_at=existing.created_at if existing else datetime.now(), updated_at=datetime.now(),
            image_path=existing.image_path if existing else "",
            image_hash=existing.image_hash if existing else "",
        )
        if image_path and Path(image_path).is_file():
            receipt.image_hash = image_digest(image_path)
        for index, values in enumerate(self.items, 1):
            parts = [value.strip() for value in values]
            if not any(parts):
                continue
            try:
                quantity = Decimal(parts[1]) if parts[1] else None
                if quantity is not None and (not quantity.is_finite() or quantity < 0):
                    raise InvalidOperation
            except (InvalidOperation, ValueError) as exc:
                raise LedgerError(f"明細 {index} 行目の数量を確認してください。") from exc
            unit = parse_yen(parts[2], f"明細 {index} 行目の単価")
            amount = parse_yen(parts[3], f"明細 {index} 行目の金額")
            if amount is None and quantity is not None and unit is not None:
                amount = int((quantity * unit).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            receipt.items.append(LineItem(parts[0], float(quantity) if quantity is not None else None,
                                          unit, amount, receipt.category, parts[4]))
        return receipt


@dataclass
class QueueEntry:
    path: str
    state: str = "pending"
    draft: ReceiptDraft | None = None
    error: str = ""
    reviewed: bool = False
    checked: bool = False
    saved_to: str = ""
    stored_image: str = ""
    image_hash: str = ""


class BatchReadWorker(QThread):
    image_started = Signal(str, int, int)
    image_succeeded = Signal(str, object)
    image_failed = Signal(str, str)
    progress = Signal(int, int)

    def __init__(self, paths, client, model, parent=None, source_paths=None):
        super().__init__(parent)
        self.paths = tuple(paths)
        self.client = client
        self.model = model
        self.source_paths = source_paths or {}

    def run(self):
        for index, path in enumerate(self.paths, 1):
            if self.isInterruptionRequested():
                break
            self.image_started.emit(path, index, len(self.paths))
            try:
                result = self.client.extract(self.source_paths.get(path, path), self.model)
                self.image_succeeded.emit(path, result)
            except Exception as exc:
                self.image_failed.emit(path, str(exc))
            self.progress.emit(index, len(self.paths))
