# Linux Manual Verification Summary

- 実施日時（JST）: 2026-04-28 17:34:52 - 2026-04-28 17:37:35
- 実行環境: docker compose service `lo-vnc` / 07f9ea3ecc72a61f30dfd36dba2d148eb25eaefc6f6a18f79d32e859b7e6955b / image=setpasstoexceldotnet-lo-vnc / status=running / started=2026-04-28T01:51:17.089880716Z / ip=172.23.0.2
- LibreOffice バージョン: LibreOffice 24.2.7.2 420(Build:2)
- noVNC: http://localhost:6901/ (`VNC_PASSWORD` from `.env`)
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

- docker compose service: lo-vnc (ports 5901/tcp, 6901/tcp)
- 誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。
- ファイル破損警告、修復ダイアログ、内容削除/修復の確認ダイアログは検出されなかった。
- excel_en.xlsm: macro prompt/bar was not blocking open/save/reopen in this run
- excel_ja.xlsm: macro prompt/bar was not blocking open/save/reopen in this run

## Commands

```bash
docker compose up -d
docker compose ps
docker inspect lo-vnc --format ...
docker compose exec -T lo-vnc bash -lc 'libreoffice --version; command -v xdotool; command -v import; command -v xclip'
mkdir -p docs/manual-evidence/linux/workbooks
cp test-manual-files/*.xlsx test-manual-files/*.xlsm docs/manual-evidence/linux/workbooks/
docker compose exec -T lo-vnc python3 /workspace/docs/manual-evidence/linux/run_libreoffice_manual_verification.py
```
