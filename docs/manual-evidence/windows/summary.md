# Windows Manual Verification Summary

- Executed (JST): 2026-04-28 18:17:05 - 2026-04-28 18:18:38
- Environment: Microsoft Windows 11 Pro Insider Preview / Version 10.0.26300 / Build 26300 / LENOVO 20V9
- Excel version: Microsoft Excel for Windows 16.0.19127.20302
- Overall result: PASS

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

- Wrong password rejection was verified with `simple_en.xlsx` and `simple_ja.xlsx` as password-type representative cases.
- `.xlsm` files were opened with Excel COM `AutomationSecurity = 1` for this manual verification run; macro prompts did not block open/save/reopen.
- No file corruption warning, repair dialog, removed content prompt, save failure, or reopen failure was observed during the automated Excel COM run.

## Commands

```powershell
Get-ComputerInfo | Select-Object OsName, OsVersion, OsBuildNumber, CsManufacturer, CsModel
Get-Item "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
powershell -NoProfile -ExecutionPolicy Bypass -File docs\manual-evidence\windows\run_excel_manual_verification.ps1
dotnet test
```
