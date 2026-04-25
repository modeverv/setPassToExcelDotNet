# Manual Compatibility Checklist

このチェックリストは、Microsoft Excel / LibreOffice の GUI 実機確認をリリースごとに記録するためのものです。

## Scope

- 対象形式: `.xlsx`, `.xlsm`
- 対象アプリ: Excel for Windows, Excel for Mac, LibreOffice Calc
- 対象ビルド: リリース候補（RC）またはリリースタグの成果物

## Release Record

| Release Version | Check Date | Checker  | Environment                                                         | Result | Notes |
|-----------------|------------|----------|---------------------------------------------------------------------|--------|-------|
| 2.0.5           | 2026-04-25 | modeverv | mac - 16.105.1<br/>win - 16.0.19127.20302<br/>linux - calc 24.2.7.2 | OK     | -     |

Environment は最低でも `OS version - App Version` を記録する。

## Test Inputs

- 既定パスワード: `pass`
- 日本語パスワード: `パスワード`
- 参照ファイル:
  - `test-vectors/plain/simple.xlsx`
  - `test-vectors/plain/japanese.xlsx`
  - `test-vectors/xlsm/excel_sample.xlsm`
- 生成ファイル（`dotnet test` で作成）:
  - `test-manual-files/excel_image_en.xlsx`
  - `test-manual-files/simple_en.xlsx`
  - `test-manual-files/japanese_en.xlsx`
  - `test-manual-files/excel_en.xlsm`
  - `test-manual-files/excel_image_ja.xlsx`
  - `test-manual-files/simple_ja.xlsx`
  - `test-manual-files/japanese_ja.xlsx`
  - `test-manual-files/excel_ja.xlsm`

## Excel for Windows

- [x] 正しいパスワードで開ける
- [x] 間違ったパスワードで開けない
- [x] 日本語パスワードで開ける
- [x] 日本語シート名が崩れない
- [x] 再保存後に再オープンできる
- [x] 再保存後ファイルが破損しない

## Excel for Mac

- [x] 正しいパスワードで開ける
- [x] 間違ったパスワードで開けない
- [x] 日本語パスワードで開ける
- [x] 日本語シート名が崩れない
- [x] 再保存後に再オープンできる
- [x] 再保存後ファイルが破損しない

## LibreOffice Calc

- [x] 正しいパスワードで開ける
- [x] 間違ったパスワードで開けない
- [x] 日本語パスワードで開ける
- [x] 日本語シート名が崩れない
- [x] 再保存後に再オープンできる
- [x] 再保存後ファイルが破損しない

## .xlsm Specific Checks

- [x] `.xlsm` を正しいパスワードで開ける
- [x] VBA プロジェクトが消失していない
- [x] マクロを実行できる
- [x] 再保存後もマクロを実行できる
- [x] 再保存後も再オープンできる

## Release Gate Rule

- [x] リリース前にこのチェックリストを埋める
- [x] すべての必須項目が完了するまでリリースしない
- [x] 不合格項目がある場合は、Issue を起票して再確認日を記録する
- [x] 完了した記録は Pull Request またはリリースノートにリンクする

## DockerでのLinux環境の確認

```bash
cp .env.example .env
docker compose up -d
open http://localhost:6901/
```
