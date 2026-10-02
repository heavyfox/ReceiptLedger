# GitHubリリース手順

公開先: https://github.com/heavyfox/ReceiptLedger

この資料はリリースを作成・更新する際の手順です。ローカルでの準備だけではアップロードされません。

## リポジトリ

`ReceiptLedger-GitHub` フォルダの中身を公開用のリポジトリにします。開発作業フォルダ全体をアップロードしないでください。

含めるもの: ソース、テスト、架空のテスト画像、空のテンプレート、アイコン、README、MIT LICENSE、第三者ライセンス、ビルド手順、専用HEIC wheel。

含めないもの: `.venv`、`build`、`dist`、実際のレシート、家計簿、設定、下書き、認証情報、利用者のバックアップ。`.gitignore` と公開用パッケージ作成の許可リストで除外します。

## Releases

バージョン `v1.5.1` のソースにタグを付け、次のファイルを同じReleaseへ添付します。

1. `ReceiptLedger-v1.5.1-windows-x64.zip` — 利用者向け。
2. `ReceiptLedger-v1.5.1-source.zip` — この版のアプリソース。
3. `ReceiptLedger-v1.5.1-third-party-sources.zip` — LGPLライブラリの対応ソースとHEIC再ビルド手順。
4. `SHA256SUMS.txt` — 上記3ファイルのSHA-256。

**対応ソースZIPはWindows版と一緒に提供します。** 依存ライブラリの公式サイトへのリンクだけで代用せず、同じダウンロード場所から無償で取得できる状態にしてください。

説明文には、対応Windows、LM Studioと画像対応モデルの別途導入が必要なこと、Excel未導入でも保存できること、画像の送信先、変更点、更新手順へのリンクを記載します。公開用の画面例はすべて架空のデータです。

## 公開時の確認

- 公開するリポジトリ名・所有者を決定し、READMEのダウンロード案内をそのリポジトリのReleasesへリンクする。
- ZIPのチェックサムと検証レポートを確認する。
- `python tools/verify_release.py <リリースフォルダ>` が成功することを確認する。
- 既存データをバックアップして更新し、保存先と件数を確認する。
- 配布するソースとWindows版が同じバージョンであることを確認する。

コード署名はまだ付けていません。Windowsの警告を回避するためにセキュリティ機能を無効にする案内はしません。将来コード署名を導入する場合も、Qt等のライブラリを利用者が改変・交換できる条件を維持してください。

## 公開時に使用する資料

- Release本文は `docs/RELEASE_NOTES.md` の内容を使用できます。
- `.github/workflows/windows.yml` はpush・Pull Request・手動実行でWindows上のテスト、ビルド、配布実行ファイルの自己テストを行います。アップロードやRelease公開は行いません。GitHub上での初回実行結果は公開担当者が確認してください。
- ワークフローは [GitHubのPythonテスト手順](https://docs.github.com/en/actions/tutorials/build-and-test-code/python) と [checkout](https://github.com/actions/checkout)、[setup-python](https://github.com/actions/setup-python) の公式説明に基づいています。
- 不具合報告テンプレートを同梱しています。Issueで実際の認証情報・レシート・家計簿を収集しない運用にしてください。
