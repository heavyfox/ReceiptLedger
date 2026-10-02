"""Verify the serialized Excel output independently of openpyxl's object model."""
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal
import posixpath
from pathlib import PurePosixPath
import zipfile
from xml.etree import ElementTree as ET

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


class ExportValidationError(ValueError):
    pass


def _require(condition, description):
    if not condition:
        raise ExportValidationError(description)


def _xml(archive, part):
    return ET.fromstring(archive.read(part))


def _relationships(archive, part):
    parent = PurePosixPath(part)
    rels = str(parent.parent / "_rels" / (parent.name + ".rels"))
    if rels not in archive.namelist():
        return {}
    return {rel.get("Id"): posixpath.normpath(posixpath.join(str(parent.parent), rel.get("Target"))).lstrip("/")
            for rel in _xml(archive, rels) if rel.get("TargetMode") != "External"}


def _cells(archive, part, shared):
    cells = {}
    for cell in _xml(archive, part).findall(".//s:sheetData/s:row/s:c", NS):
        kind = cell.get("t")
        if kind == "inlineStr":
            value = "".join(node.text or "" for node in cell.findall(".//s:t", NS))
        else:
            raw = cell.findtext("s:v", namespaces=NS)
            value = None if raw is None else (shared[int(raw)] if kind == "s" else Decimal(raw) if kind in (None, "n") else raw)
        cells[cell.get("r")] = value
    return cells


def _charts(archive, sheet_part):
    sheet = _xml(archive, sheet_part)
    links = _relationships(archive, sheet_part)
    charts = []
    for drawing in sheet.findall("s:drawing", NS):
        part = links[drawing.get("{" + NS["r"] + "}id")]
        drawing_links = _relationships(archive, part)
        for chart in _xml(archive, part).findall(".//c:chart", NS):
            charts.append(_xml(archive, drawing_links[chart.get("{" + NS["r"] + "}id")]))
    return charts


def validate_export(path, receipts, categories):
    """Raise before committing a workbook if totals, tax or chart labels are lost.

    This checks file contents, not native Excel rendering. User-created sheets
    and charts outside the managed monthly summary are deliberately ignored.
    """
    try:
        with zipfile.ZipFile(path) as archive:
            workbook = _xml(archive, "xl/workbook.xml")
            links = _relationships(archive, "xl/workbook.xml")
            sheets = {node.get("name"): links[node.get("{" + NS["r"] + "}id")]
                      for node in workbook.findall("s:sheets/s:sheet", NS)}
            _require(all(name in sheets for name in ("ダッシュボード", "レシート別", "支出一覧", "明細", "月別集計")), "家計簿のシートが不足しています。")
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                shared = ["".join(t.text or "" for t in si.findall(".//s:t", NS))
                          for si in _xml(archive, "xl/sharedStrings.xml")]
            expense = _cells(archive, sheets["支出一覧"], shared)
            rows = sorted(int(cell[1:]) for cell, value in expense.items() if cell.startswith("A") and cell != "A1" and value)
            _require(len(rows) == len(receipts), "保存したレシート件数が一致しません。")
            by_id = {r.receipt_id: r for r in receipts}
            _require(Counter(expense[f"A{row}"] for row in rows) == Counter(list(by_id)), "保存したレシートIDが重複しています。")
            for row in rows:
                receipt = by_id.get(expense[f"A{row}"])
                _require(receipt is not None, "保存したレシートIDが一致しません。")
                _require(expense.get(f"E{row}") == receipt.amount and expense.get(f"J{row}") == receipt.tax,
                         "税込合計または消費税額が一致しません。")

            summary_part = sheets["月別集計"]
            summary = _cells(archive, summary_part, shared)
            _require(summary.get("B6") == sum(r.amount for r in receipts), "累計支出が一致しません。")
            _require(summary.get("E6") == len(receipts), "累計件数が一致しません。")
            charts = _charts(archive, summary_part)
            if not receipts:
                _require(not charts, "空の家計簿に月別グラフがあります。")
                return
            first = min(r.purchased_on for r in receipts).replace(day=1)
            last = max(r.purchased_on for r in receipts).replace(day=1)
            months = []
            current = first
            while current <= last:
                months.append(current)
                current = date(current.year + (current.month == 12), current.month % 12 + 1, 1)
            totals = Counter()
            counts = Counter()
            breakdown = Counter()
            for receipt in receipts:
                month = receipt.purchased_on.replace(day=1)
                totals[month] += receipt.amount
                counts[month] += 1
                breakdown[month, receipt.category] += receipt.amount
            properties = workbook.find("s:workbookPr", NS)
            epoch = date(1904, 1, 1) if properties is not None and properties.get("date1904") in ("1", "true") else date(1899, 12, 30)
            for index, month in enumerate(months, 10):
                serial = summary.get(f"B{index}")
                _require(isinstance(serial, Decimal), "月別集計の年月が不正です。")
                day = int(serial) + (epoch == date(1899, 12, 30) and 0 < serial < 60)
                _require(epoch + timedelta(days=day) == month, "月別集計の年月が一致しません。")
                _require(summary.get(f"C{index}") == totals[month] and summary.get(f"D{index}") == counts[month], "月別の税込合計または件数が一致しません。")
                for column, category in enumerate(categories, 5):
                    _require(summary.get(f"{chr(64 + column)}9") == category, "費目の見出しが一致しません。")
                    _require(summary.get(f"{chr(64 + column)}{index}") == breakdown[month, category], "月別の費目合計が一致しません。")
            _require(summary.get(f"C{10 + len(months)}") == sum(totals.values()), "月別集計の合計行が一致しません。")
            years = sorted({m.year for m in months})
            _require(len(charts) == len(years), "年ごとのグラフ件数が一致しません。")
            for chart, year in zip(charts, years):
                year_months = [m for m in months if m.year == year]
                first_row = 10 + months.index(year_months[0])
                last_row = first_row + len(year_months) - 1
                title = "".join(t.text or "" for t in chart.findall(".//c:title//a:t", NS))
                _require(str(year) in title, "グラフの対象年がありません。")
                series = chart.findall(".//c:barChart/c:ser", NS)
                _require(len(series) == 1, "月別グラフの系列が不正です。")
                labels = [pt.findtext("c:v", namespaces=NS) for pt in series[0].findall("c:cat/c:strLit/c:pt", NS)]
                _require(labels == [f"{m.month}月" for m in year_months], "グラフの横軸の月名が一致しません。")
                references = {f"'月別集計'!$C${first_row}:$C${last_row}"}
                if first_row == last_row:
                    references.add(f"'月別集計'!$C${first_row}")
                _require(series[0].findtext("c:val/c:numRef/c:f", namespaces=NS) in references, "グラフの金額参照が不正です。")
                labels = chart.find(".//c:barChart/c:dLbls", NS)
                _require(labels is not None and labels.find("c:showVal", NS).get("val") in ("1", "true"), "グラフの金額ラベルが非表示です。")
                _require("円" in labels.find("c:numFmt", NS).get("formatCode", ""), "グラフの金額単位がありません。")
                for tag, position in (("catAx", "b"), ("valAx", "l")):
                    axis = chart.find(f".//c:{tag}", NS)
                    _require(axis is not None and axis.find("c:delete", NS).get("val") in ("0", "false"), "グラフの軸が非表示です。")
                    _require(axis.find("c:axPos", NS).get("val") == position, "グラフの軸位置が不正です。")
                    _require(axis.find("c:tickLblPos", NS).get("val") != "none", "グラフの目盛りが非表示です。")
    except ExportValidationError:
        raise
    except (KeyError, ValueError, TypeError, AttributeError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise ExportValidationError("Excel出力の構造を検証できませんでした。") from exc
