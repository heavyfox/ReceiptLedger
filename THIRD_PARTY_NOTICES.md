# 第三者ライブラリ・対応ソース

## この配布物

ReceiptLedger 1.2.2の独自コード・アイコン・説明書・架空のテスト画像はMITライセンスです。第三者のコード、共有ライブラリ、フォントエンジン、画像コーデック等には、それぞれのライセンスが適用されます。著作権は各権利者に帰属します。

| 部品 | 使用版 | 主な適用条件・同梱文書 |
|---|---|---|
| Qt Base / PySide6 / Shiboken6 | 6.11.2 | LGPLv3を選択。`licenses/native-sources/` の各 `LICENSES` と著作権・第三者通知 |
| libheif | 1.23.4 | LGPL-3.0-or-later、共有DLL。`licenses/native-sources/libheif-1.23.4/COPYING` |
| libde265 | 1.1.3 | LGPL-3.0-or-later、共有DLL。`licenses/native-sources/libde265-1.1.3/COPYING` |
| pillow-heif | 1.8.0+receiptledger1 | Python拡張・DLL探索パス変更はBSD-3-Clause。同梱ネイティブDLLは上記LGPL |
| Python | 3.12.14 | PSF等。`licenses/Python/LICENSE.txt` |
| Pillow | 12.3.0 | MIT-CMUおよび同梱画像ライブラリの条件。`licenses/python-packages/pillow-12.3.0/licenses/LICENSE` |
| openpyxl / et-xmlfile | 3.1.5 / 2.0.0 | MIT。`licenses/python-packages/` |
| keyringと補助Pythonパッケージ | `DEPENDENCIES.json`参照 | 各MIT/BSD等の文書を `licenses/python-packages/` に同梱 |
| OpenSSL | 3.5.8 | Apache-2.0。`licenses/OpenSSL/` |
| libffi | Pythonに付属するWindows DLL | MIT。`licenses/libffi/LICENSE.txt` |
| PyInstallerブートローダー | 6.22.3 | GPL-2.0-or-laterとBootloader例外。独自アプリをMITで配布することを妨げません。`licenses/python-packages/pyinstaller-6.22.3/licenses/COPYING.txt` |
| Microsoft C/C++ runtime | 各DLLのファイル情報参照 | Microsoftの再頒布可能コード。`licenses/Microsoft/VC-Runtime.docx` |

依存一覧にはビルド用パッケージも含みます。元ソース由来のライセンス集には、上流プロジェクトの任意機能や開発ツールの文書も含みます。それらすべてがWindows版に組み込まれているという意味ではありません。

## HEICの構成と変更

この配布版はHEIC読み取り専用です。pillow-heifの一般公開wheelをそのまま再配布していません。libheif / libde265を公式ソースからMSVCで共有DLLとしてビルドし、x265・x264等の圧縮エンコーダーを無効にしました。x265とMinGWランタイムは配布物に含みません。

pillow-heif本体への変更は、DLL探索パスの登録、専用バージョン識別子、同梱DLL・ライセンス案内です。変更日: 2026-09-25。変更者: ReceiptLedger contributors。変更コードもBSD-3-Clauseで提供します。変更を適用する手順は `tools/build_heif.py` に含まれます。

元のpillow-heifソースアーカイブ内の `LICENSES_bundled.txt` は上流の一般公開wheelを説明しています。本配布版に実際に含まれる部品はこの文書および `licenses/python-packages/pillow_heif-1.8.0+receiptledger1/` の案内に従ってください。

## 対応ソースの提供

Windows版と同じ配布場所に、次を提供します。

- `ReceiptLedger-v1.2.2-source.zip`: アプリのソース、ビルドスクリプト、専用HEIC wheel、ライセンス。
- `ReceiptLedger-v1.2.2-third-party-sources.zip`: Qt Base / PySide / Shiboken / libheif / libde265 / pillow-heifの対応する元ソースと変更手順、取得元・ハッシュ一覧。

公開担当者はこの2つをWindows版と同じReleaseへ無償で添付してください。ローカル準備段階では、同じ出力フォルダにあります。対応ソースはインストールや通常利用には不要です。

Qt Base / PySide / Shibokenは上流の公式バイナリを使用し、コード変更はしていません。アプリは共有DLLを読み込みます。LGPL対象ライブラリの改変・ABI互換のDLLへの交換、およびそのデバッグに必要なリバースエンジニアリングを制限しません。[ビルドと差し替え手順](docs/BUILD.md) を参照してください。LGPL本文とそれが参照するGPL本文はQtソース由来の `LICENSES` に同梱しています。

LM Studioおよびモデルは同梱していません。取得先が示す利用条件をご確認ください。
