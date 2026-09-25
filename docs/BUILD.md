# 開発・ビルド

対象: Windows x64、Python 3.12。アプリ利用者はPythonの導入不要です。

## ソースから起動する

このリポジトリのルートをPowerShellで開きます。

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-lock.txt
.\.venv\Scripts\python.exe main.py
```

`vendor` のHEIC wheelは、pillow-heif 1.8.0を基にした読み取り専用のWindows x64 / CPython 3.12版です。通常のPyPI版に置き換えると、書き出し用x265が再び同梱され、配布条件も変わります。配布用ビルドは専用版の使用を確認してから進みます。

## テストとWindows版の作成

```powershell
$env:QT_QPA_PLATFORM = 'offscreen'
.\.venv\Scripts\python.exe -m unittest discover -p 'test_*.py'
Remove-Item Env:QT_QPA_PLATFORM
.\build.ps1 -Python .\.venv\Scripts\python.exe
```

出力は `dist/ReceiptLedger/ReceiptLedger.exe`。`_internal` を必ずセットで配布します。
ビルド後の自己テストは `ReceiptLedger.exe --self-test` です。認証トークンや本番データを使わず、一時フォルダ内で画像変換・Excel記録・下書き・一括処理・削除取消・バックアップ復元を検証し、成功時に終了コード0を返します。

## HEIC wheelの再ビルド

Visual Studio 2022 C++ Build Tools、Windows SDK、CMakeを導入して実行します。システムのライブラリは変更せず、`--work` 配下へビルドします。

```powershell
.\.venv\Scripts\python.exe -m pip install wheel ninja
.\.venv\Scripts\python.exe tools/build_heif.py --work build/native --archives build/sources --wheel-dir vendor
```

公式配布元の次のソースをSHA-256で確認して使用します。すでに `--archives` にファイルがあれば通信しません。

- libheif 1.23.4 / libde265 1.1.3: LGPL-3.0-or-later。共有DLLとしてビルド。
- pillow-heif 1.8.0: BSD-3-Clause。DLL探索パス、バージョン識別子、同梱DLLとライセンス案内を変更。
- HEVC読み取りはlibde265を使用。x265、x264などの圧縮エンコーダーは無効。libheif内部の非圧縮mask処理は残る場合があります。

変更内容は `tools/build_heif.py` に含まれ、元ソースと合わせて再現できます。生成wheelの再ビルドはタイムスタンプ等によりハッシュが変わるため、`vendor/SHA256SUMS.txt` も更新して検証してください。

## 配布構成

- `ReceiptLedger.spec` と `hooks/` で未使用の画像形式・Qtプラグインを収集対象から除外します。
- UIはQt Widgetsのラスタ描画。OpenGL/QMLの描画機能は使用しません。
- ソースとDLLは改変・差し替え可能です。Qt/PySide6の更新時はABIの組み合わせを揃え、HEIC実画像・画像回転・ドラッグ＆ドロップ・Excelの税額・復元処理を再確認してください。
- アイコンの原稿は `assets/receipt-ledger.svg`。`python tools/render_icon.py` でPNG/ICOを再生成します。
- `tools/collect_licenses.py --sources <ソースアーカイブのフォルダ>` で依存パッケージと元ソースのライセンスを収集できます。
- `tools/package_release.py --sources <ソースアーカイブのフォルダ> --output <出力先>` で公開用のソース・Windows ZIP・対応ソースZIPとチェックサムを作成します。アップロードは行いません。
- `python tools/verify_release.py <出力先>` で3つのZIPのSHA-256・CRC・内部パス、公開用ソースの必須ファイルと画像・家計簿の混入、Windows版の全ファイルのマニフェストを照合します。

## 対応ソースとDLLの差し替え

Qt / PySide6 / Shiboken6は公式6.11.2バイナリを使用しています。対応ソースZIPには同バージョンのQt Base、PySideセットアップ一式、HEICの元ソースとビルド手順が含まれます。Qtの公式手順は https://doc.qt.io/qt-6/windows-building.html 、PySideの手順は https://doc.qt.io/qtforpython-6/building_from_source/index.html を参照してください。

アプリを終了後、ABI互換の自作DLLを `_internal/PySide6/`、`_internal/shiboken6/`、`_internal/pillow_heif/` に配置できます。pillow-heifの拡張モジュールを再ビルドした場合は `_internal/_pillow_heif*.pyd` も対応する版に交換します。署名・ハッシュによるDLL交換の禁止はありません。改変箇所のデバッグを目的としたリバースエンジニアリングも禁止しません。
