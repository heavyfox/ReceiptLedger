# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

root = Path(SPECPATH)
datas, binaries, hiddenimports = collect_all('keyring')
datas += [(str(root / 'assets/receipt-ledger.ico'), 'assets'),
          (str(root / 'tests/fixtures/receipt.heic'), 'tests/fixtures')]
hiddenimports += ['PySide6.support.deprecated']
a = Analysis([str(root / 'main.py')], pathex=[str(root)], binaries=binaries,
             datas=datas, hiddenimports=hiddenimports, hookspath=[str(root / 'hooks')],
             excludes=['PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtPdf', 'PIL._avif'],
             noarchive=False, optimize=0)
# All application labels are Japanese source strings. No QTranslator is loaded;
# native Windows file dialogs use the OS language. Qt translation packs are unused.
a.datas = [entry for entry in a.datas if not entry[0].replace('\\', '/').startswith('PySide6/translations/')]
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='ReceiptLedger',
          debug=False, strip=False, upx=False, console=False,
          icon=str(root / 'assets/receipt-ledger.ico'),
          version=str(root / 'assets/windows-version.txt'))
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='ReceiptLedger')
