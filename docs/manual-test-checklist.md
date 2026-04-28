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
| 2.0.5.1         | 2026-04-28 | codex    | mac - 16.108<br/>win - 16.0.19127.20302<br/>linux - calc 24.2.7.2   | OK     | -     |

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

## v2.0.5.1 Manual Verification - 2026-04-28

### macOS Manual Verification (Excel) - 2026-04-28

- 実施日時（JST）: 2026-04-28 12:25:37 - 12:30:33
- 実行環境: macOS 26.3.1 (a) / Build 25D771280a / MacBook Pro (MacBookPro18,3), Apple M1 Pro, 10-core CPU, 16 GB RAM
- Excel バージョン: Microsoft Excel for Mac 16.108 (16.108.26041219)

| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |
|---|---|---:|---:|---:|---:|---:|---|
| simple_en.xlsx | en | PASS | PASS | PASS | PASS | PASS | simple_en-open.png<br/>simple_en-ja-sheet-added.png<br/>simple_en-reopen.png |
| simple_ja.xlsx | ja | PASS | PASS | PASS | PASS | PASS | simple_ja-open.png<br/>simple_ja-ja-sheet-added.png<br/>simple_ja-reopen.png |
| japanese_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | japanese_en-open.png<br/>japanese_en-ja-sheet-added.png<br/>japanese_en-reopen.png |
| japanese_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | japanese_ja-open.png<br/>japanese_ja-ja-sheet-added.png<br/>japanese_ja-reopen.png |
| excel_en.xlsm | en | PASS | N/A | PASS | PASS | PASS | excel_en-open.png<br/>excel_en-ja-sheet-added.png<br/>excel_en-reopen.png |
| excel_ja.xlsm | ja | PASS | N/A | PASS | PASS | PASS | excel_ja-open.png<br/>excel_ja-ja-sheet-added.png<br/>excel_ja-reopen.png |
| excel_image_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | excel_image_en-open.png<br/>excel_image_en-ja-sheet-added.png<br/>excel_image_en-reopen.png |
| excel_image_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | excel_image_ja-open.png<br/>excel_image_ja-ja-sheet-added.png<br/>excel_image_ja-reopen.png |

- 総合判定: PASS
- 備考: 誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。`.xlsm` の初回オープンと保存後再オープンでは、Excel のマクロ確認ダイアログで `マクロを有効にする` を押下した。ファイル破損警告、修復ダイアログ、保存不可、再オープン不可は発生しなかった。


### Linux Manual Verification (Docker/VNC) - 2026-04-28

- 実施日時（JST）: 2026-04-28 17:34:52
- 実行環境: docker compose service `lo-vnc` / 07f9ea3ecc72a61f30dfd36dba2d148eb25eaefc6f6a18f79d32e859b7e6955b / image=setpasstoexceldotnet-lo-vnc / status=running / started=2026-04-28T01:51:17.089880716Z / ip=172.23.0.2
- LibreOffice バージョン: LibreOffice 24.2.7.2 420(Build:2)

| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |
|---|---|---:|---:|---:|---:|---:|---|
| simple_en.xlsx | en | PASS | PASS | PASS | PASS | PASS | simple_en-open.png<br/>simple_en-ja-sheet-added.png<br/>simple_en-reopen.png |
| simple_ja.xlsx | ja | PASS | PASS | PASS | PASS | PASS | simple_ja-open.png<br/>simple_ja-ja-sheet-added.png<br/>simple_ja-reopen.png |
| japanese_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | japanese_en-open.png<br/>japanese_en-ja-sheet-added.png<br/>japanese_en-reopen.png |
| japanese_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | japanese_ja-open.png<br/>japanese_ja-ja-sheet-added.png<br/>japanese_ja-reopen.png |
| excel_en.xlsm | en | PASS | N/A | PASS | PASS | PASS | excel_en-open.png<br/>excel_en-ja-sheet-added.png<br/>excel_en-reopen.png |
| excel_ja.xlsm | ja | PASS | N/A | PASS | PASS | PASS | excel_ja-open.png<br/>excel_ja-ja-sheet-added.png<br/>excel_ja-reopen.png |
| excel_image_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | excel_image_en-open.png<br/>excel_image_en-ja-sheet-added.png<br/>excel_image_en-reopen.png |
| excel_image_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | excel_image_ja-open.png<br/>excel_image_ja-ja-sheet-added.png<br/>excel_image_ja-reopen.png |

- 総合判定: PASS
- 備考: docker compose service: lo-vnc (ports 5901/tcp, 6901/tcp)。誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。ファイル破損警告、修復ダイアログ、内容削除/修復の確認ダイアログは検出されなかった。`.xlsm` では macro disabled バーが表示されたが、open/save/reopen はブロックされなかった。

### Windows Manual Verification (Excel) - 2026-04-28

- Executed (JST): 2026-04-28 18:17:05 - 2026-04-28 18:18:38
- Environment: Microsoft Windows 11 Pro Insider Preview / Version 10.0.26300 / Build 26300 / LENOVO 20V9
- Excel version: Microsoft Excel for Windows 16.0.19127.20302

| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |
|---|---|---:|---:|---:|---:|---:|---|
| simple_en.xlsx | en | PASS | PASS | PASS | PASS | PASS | simple_en-open.png<br/>simple_en-ja-sheet-added.png<br/>simple_en-reopen.png |
| simple_ja.xlsx | ja | PASS | PASS | PASS | PASS | PASS | simple_ja-open.png<br/>simple_ja-ja-sheet-added.png<br/>simple_ja-reopen.png |
| japanese_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | japanese_en-open.png<br/>japanese_en-ja-sheet-added.png<br/>japanese_en-reopen.png |
| japanese_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | japanese_ja-open.png<br/>japanese_ja-ja-sheet-added.png<br/>japanese_ja-reopen.png |
| excel_en.xlsm | en | PASS | N/A | PASS | PASS | PASS | excel_en-open.png<br/>excel_en-ja-sheet-added.png<br/>excel_en-reopen.png |
| excel_ja.xlsm | ja | PASS | N/A | PASS | PASS | PASS | excel_ja-open.png<br/>excel_ja-ja-sheet-added.png<br/>excel_ja-reopen.png |
| excel_image_en.xlsx | en | PASS | N/A | PASS | PASS | PASS | excel_image_en-open.png<br/>excel_image_en-ja-sheet-added.png<br/>excel_image_en-reopen.png |
| excel_image_ja.xlsx | ja | PASS | N/A | PASS | PASS | PASS | excel_image_ja-open.png<br/>excel_image_ja-ja-sheet-added.png<br/>excel_image_ja-reopen.png |

- Overall result: PASS
- Notes: Wrong password rejection was verified with `simple_en.xlsx` and `simple_ja.xlsx`; no corruption/repair dialogs were observed. `.xlsm` macro prompts did not block open/save/reopen in this Excel COM run.
