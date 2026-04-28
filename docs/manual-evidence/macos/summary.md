# macOS Manual Verification Summary

- 実施日時（JST）: 2026-04-28 12:25:37 - 12:30:33
- 実行環境: macOS 26.3.1 (a) / Build 25D771280a / MacBook Pro (MacBookPro18,3), Apple M1 Pro, 10-core CPU, 16 GB RAM
- Excel バージョン: Microsoft Excel for Mac 16.108 (16.108.26041219)
- 総合判定: PASS

## Result

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

## Notes

- 誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。
- `.xlsm` の初回オープンと保存後再オープンでは、Excel のマクロ確認ダイアログで `マクロを有効にする` を押下して検証を継続した。
- ファイル破損警告、修復ダイアログ、保存不可、再オープン不可はいずれも発生しなかった。

## Commands

```bash
mkdir -p docs/manual-evidence/macos/workbooks
cp test-manual-files/simple_en.xlsx test-manual-files/simple_ja.xlsx test-manual-files/japanese_en.xlsx test-manual-files/japanese_ja.xlsx test-manual-files/excel_en.xlsm test-manual-files/excel_ja.xlsm test-manual-files/excel_image_en.xlsx test-manual-files/excel_image_ja.xlsx docs/manual-evidence/macos/workbooks/
system_profiler SPHardwareDataType SPSoftwareDataType
/usr/libexec/PlistBuddy -c Print:CFBundleShortVersionString /Applications/Microsoft\ Excel.app/Contents/Info.plist
/usr/libexec/PlistBuddy -c Print:CFBundleVersion /Applications/Microsoft\ Excel.app/Contents/Info.plist
osascript docs/manual-evidence/macos/run_excel_manual_verification.applescript
dotnet test
```
