# Manual Compatibility Checklist

このチェックリストは、Microsoft Excel / LibreOffice の GUI 実機確認をリリースごとに記録するためのものです。

## Scope

- 対象形式: `.xlsx`, `.xlsm`
- 対象アプリ: Excel for Windows, Excel for Mac, LibreOffice Calc
- 対象ビルド: リリース候補（RC）またはリリースタグの成果物

## Release Record

| Release Version | Check Date | Checker | Environment | Result | Notes |
|---|---|---|---|---|---|
|  |  |  |  |  |  |

Environment は最低でも `OS version + App version` を記録する。

## Test Inputs

- 既定パスワード: `pass`
- 日本語パスワード: `日本語パスワード`
- 参照ファイル:
  - `test-vectors/plain/simple.xlsx`
  - `test-vectors/plain/japanese.xlsx`
  - `test-vectors/xlsm/excel_sample.xlsm`
- 生成ファイル（`dotnet test` で作成）:
  - `test-manual-files/simple_aes256_sha512.xlsx`
  - `test-manual-files/japanese_aes256_sha512.xlsx`
  - `test-manual-files/excel_sample_aes256_sha512.xlsm`

## Excel for Windows

- [ ] 正しいパスワードで開ける
- [ ] 間違ったパスワードで開けない
- [ ] 日本語パスワードで開ける
- [ ] 日本語シート名が崩れない
- [ ] 再保存後に再オープンできる
- [ ] 再保存後ファイルが破損しない

## Excel for Mac

- [ ] 正しいパスワードで開ける
- [ ] 間違ったパスワードで開けない
- [ ] 日本語パスワードで開ける
- [ ] 日本語シート名が崩れない
- [ ] 再保存後に再オープンできる
- [ ] 再保存後ファイルが破損しない

## LibreOffice Calc

- [ ] 正しいパスワードで開ける
- [ ] 間違ったパスワードで開けない
- [ ] 日本語パスワードで開ける
- [ ] 日本語シート名が崩れない
- [ ] 再保存後に再オープンできる
- [ ] 再保存後ファイルが破損しない

## .xlsm Specific Checks

- [ ] `.xlsm` を正しいパスワードで開ける
- [ ] VBA プロジェクトが消失していない
- [ ] マクロを実行できる
- [ ] 再保存後もマクロを実行できる
- [ ] 再保存後も再オープンできる

## Release Gate Rule

- [ ] リリース前にこのチェックリストを埋める
- [ ] すべての必須項目が完了するまでリリースしない
- [ ] 不合格項目がある場合は、Issue を起票して再確認日を記録する
- [ ] 完了した記録は Pull Request またはリリースノートにリンクする

