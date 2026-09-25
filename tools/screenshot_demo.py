"""Create documentation screenshots from the real UI using only synthetic data."""
import os
from pathlib import Path
import sys
import tempfile

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from PIL import Image, ImageDraw, ImageFont
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFontDatabase, QFont
from PySide6.QtTest import QTest
from app import MainWindow
from config import Settings
from lm_client import _parse_receipt


def main():
    output = ROOT / 'docs/screenshots'
    output.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp:
        folder = Path(temp).resolve()
        os.environ['RECEIPT_LEDGER_DATA_DIR'] = temp
        Settings(workbook_path=str(folder / '家計簿.xlsx'), theme='light').save()
        sample = folder / 'サンプル_文具店.jpg'
        image = Image.new('RGB', (700, 1040), '#fffdf7')
        draw = ImageDraw.Draw(image)
        font = ImageFont.truetype('C:/Windows/Fonts/meiryo.ttc', 29)
        title = ImageFont.truetype('C:/Windows/Fonts/meiryob.ttc', 42)
        draw.text((100, 45), 'サンプル文具店', font=title, fill='#152e38')
        lines = ['架空のレシート・操作例', '', '2026年9月25日 10:30', '------------------------------',
                 'ノート                 330円', 'ペン      2点 × 220円  440円', 'ファイル               330円',
                 '------------------------------', '小計                 1,100円', '消費税（10%）          110円',
                 '合計                 1,210円', '', 'お支払方法：現金', '', 'ご利用ありがとうございました']
        for i, line in enumerate(lines):
            draw.text((60, 135 + i * 48), line, font=font, fill='#152e38')
        image.save(sample, quality=95)
        app = QApplication([])
        QFontDatabase.addApplicationFont('C:/Windows/Fonts/meiryo.ttc')
        QFontDatabase.addApplicationFont('C:/Windows/Fonts/meiryob.ttc')
        app.setFont(QFont('Meiryo', 10))
        window = MainWindow()
        window.resize(1500, 1040)
        window.show()
        window._enqueue_images([str(sample)])
        window._batch_image_succeeded(str(sample), _parse_receipt({
            'receipt_date':'2026-09-25', 'merchant':'サンプル文具店', 'total_yen':1210, 'tax_yen':110,
            'category':'日用品', 'payment_method':'現金', 'items':[
                {'name':'ノート', 'quantity':1, 'unit_price_yen':330, 'line_total_yen':330},
                {'name':'ペン', 'quantity':2, 'unit_price_yen':220, 'line_total_yen':440},
                {'name':'ファイル', 'quantity':1, 'unit_price_yen':330, 'line_total_yen':330}]}))
        window.check_visible_btn.click()
        window.batch_progress.setRange(0, 1)
        window.batch_progress.setValue(1)
        window.batch_status.setText('読み取り完了。商品と合計を確認してExcelへ記録できます。')
        for mode in ('light', 'dark'):
            window.theme_selector.setCurrentIndex(window.theme_selector.findData(mode))
            app.processEvents()
            # Replace the temporary machine path in the documentation only.
            window.receipt_save_target.setText('記録先：家計簿.xlsx')
            window.statusBar().showMessage('操作例：表示されている店舗・商品・金額はすべて架空です。')
            window.preview.view.viewport().update()
            QTest.qWait(150)
            assert window.grab().save(str(output / f'{mode}.png'))
        window.close()


if __name__ == '__main__':
    main()
