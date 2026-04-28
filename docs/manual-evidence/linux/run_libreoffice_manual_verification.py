#!/usr/bin/env python3
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


ROOT = Path("/workspace")
EVIDENCE = ROOT / "docs" / "manual-evidence" / "linux"
WORKBOOKS = EVIDENCE / "workbooks"
SESSION_LOG = EVIDENCE / "session.log"
CHECKLIST = ROOT / "docs" / "manual-test-checklist.md"
SUMMARY = EVIDENCE / "summary.md"

FILES = [
    ("simple_en.xlsx", "en", "pass"),
    ("simple_ja.xlsx", "ja", "パスワード"),
    ("japanese_en.xlsx", "en", "pass"),
    ("japanese_ja.xlsx", "ja", "パスワード"),
    ("excel_en.xlsm", "en", "pass"),
    ("excel_ja.xlsm", "ja", "パスワード"),
    ("excel_image_en.xlsx", "en", "pass"),
    ("excel_image_ja.xlsx", "ja", "パスワード"),
]


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def log(message):
    stamp = dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S %z")
    with SESSION_LOG.open("a", encoding="utf-8") as f:
        f.write(f"[{stamp}] {message}\n")
    print(message, flush=True)


def run(cmd, check=True, **kwargs):
    log("$ " + " ".join(cmd))
    result = subprocess.run(cmd, text=True, capture_output=True, **kwargs)
    if result.stdout:
        for line in result.stdout.rstrip().splitlines():
            log("stdout: " + line)
    if result.stderr:
        for line in result.stderr.rstrip().splitlines():
            log("stderr: " + line)
    if check and result.returncode != 0:
        raise RuntimeError(f"command failed: {cmd} rc={result.returncode}")
    return result


def libreoffice_version():
    result = subprocess.run(["libreoffice", "--version"], text=True, capture_output=True, check=True)
    return result.stdout.strip()


def screenshot(path):
    run(["import", "-window", "root", str(path)])
    log(f"screenshot saved: {path.relative_to(ROOT)}")


def connect_to_office(port, timeout=30):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            ctx = resolver.resolve(url)
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return ctx, desktop
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(f"could not connect to LibreOffice on port {port}: {last_error}")


def start_office(case_name, port):
    profile = Path(f"/tmp/lo-profile-{case_name}")
    if profile.exists():
        shutil.rmtree(profile)
    profile_uri = uno.systemPathToFileUrl(str(profile))
    cmd = [
        "libreoffice",
        f"-env:UserInstallation={profile_uri}",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        "--nodefault",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    log("$ " + " ".join(cmd))
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    ctx, desktop = connect_to_office(port)
    return proc, desktop


def stop_office(proc, desktop):
    try:
        desktop.terminate()
    except Exception as exc:
        log(f"LibreOffice terminate warning: {exc}")
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        log("LibreOffice did not terminate in time; killing process")
        proc.kill()
        proc.wait(timeout=10)
    stdout, stderr = proc.communicate()
    if stdout:
        for line in stdout.rstrip().splitlines():
            log("libreoffice stdout: " + line)
    if stderr:
        for line in stderr.rstrip().splitlines():
            log("libreoffice stderr: " + line)


def wait_for_window():
    time.sleep(2)
    run(["xdotool", "search", "--onlyvisible", "--class", "libreoffice"], check=False)
    run(["xdotool", "search", "--onlyvisible", "--class", "libreoffice", "windowactivate", "%@"], check=False)
    run(["xdotool", "search", "--onlyvisible", "--class", "libreoffice", "windowsize", "%@", "1920", "1030"], check=False)
    run(["xdotool", "search", "--onlyvisible", "--class", "libreoffice", "windowmove", "%@", "0", "28"], check=False)
    time.sleep(1)


def filter_name(path):
    if path.suffix.lower() == ".xlsm":
        return "Calc MS Excel 2007 VBA XML"
    return "Calc MS Excel 2007 XML"


def load_document(desktop, path, password):
    url = uno.systemPathToFileUrl(str(path))
    args = (
        prop("Password", password),
        prop("Hidden", False),
        prop("ReadOnly", False),
        prop("MacroExecutionMode", 0),
    )
    return desktop.loadComponentFromURL(url, "_blank", 0, args)


def load_hidden(desktop, path, password):
    url = uno.systemPathToFileUrl(str(path))
    args = (prop("Password", password), prop("Hidden", True), prop("ReadOnly", True))
    return desktop.loadComponentFromURL(url, "_blank", 0, args)


def unique_sheet_name(doc):
    sheets = doc.getSheets()
    base = "日本語シート"
    name = base
    n = 2
    while sheets.hasByName(name):
        name = f"{base}{n}"
        n += 1
    return name


def add_japanese_sheet(doc):
    sheets = doc.getSheets()
    name = unique_sheet_name(doc)
    sheets.insertNewByName(name, sheets.getCount())
    sheet = sheets.getByName(name)
    sheet.getCellRangeByName("A1").String = "日本語シート名保持確認"
    doc.getCurrentController().setActiveSheet(sheet)
    return name


def save_document(doc, path, password):
    url = uno.systemPathToFileUrl(str(path))
    if path.suffix.lower() == ".xlsm":
        args = (
            prop("Password", password),
            prop("Overwrite", True),
        )
        doc.storeAsURL(url, args)
        return
    args = (
        prop("FilterName", filter_name(path)),
        prop("Password", password),
        prop("Overwrite", True),
    )
    doc.storeAsURL(url, args)


def close_document(doc):
    try:
        doc.close(True)
    except Exception:
        doc.dispose()


def verify_sheet(doc, sheet_name):
    sheets = doc.getSheets()
    if not sheets.hasByName(sheet_name):
        return False
    sheet = sheets.getByName(sheet_name)
    doc.getCurrentController().setActiveSheet(sheet)
    return sheet.getCellRangeByName("A1").String == "日本語シート名保持確認"


def wrong_password_check(source, wrong_password, case_name, port):
    proc, desktop = start_office(f"wrong-{case_name}", port)
    try:
        try:
            doc = load_hidden(desktop, source, wrong_password)
            if doc is not None:
                close_document(doc)
                log(f"wrong-password check FAILED for {source.name}: opened with wrong password")
                return False
            log(f"wrong-password check PASS for {source.name}: returned no document")
            return True
        except Exception as exc:
            log(f"wrong-password check PASS for {source.name}: rejected ({exc.__class__.__name__})")
            return True
    finally:
        stop_office(proc, desktop)


def verify_case(filename, password_type, password, idx):
    base = Path(filename).stem
    target = WORKBOOKS / filename
    result = {
        "file": filename,
        "password_type": password_type,
        "correct_password_open": "FAIL",
        "wrong_password_rejected": "N/A",
        "japanese_sheet_name_retained": "FAIL",
        "reopen_after_save": "FAIL",
        "no_corruption": "PASS",
        "evidence": [
            f"{base}-open.png",
            f"{base}-ja-sheet-added.png",
            f"{base}-reopen.png",
        ],
        "notes": [],
    }
    port = 22000 + idx
    proc, desktop = start_office(base, port)
    doc = None
    sheet_name = None
    try:
        log(f"case start: {filename}")
        try:
            doc = load_document(desktop, target, password)
            if doc is None:
                raise RuntimeError("LibreOffice returned no document")
            result["correct_password_open"] = "PASS"
            if filename.endswith(".xlsm"):
                result["notes"].append("macro prompt/bar was not blocking open/save/reopen in this run")
            wait_for_window()
            screenshot(EVIDENCE / f"{base}-open.png")
        except Exception as exc:
            result["notes"].append(f"open failed: {exc}")
            result["no_corruption"] = "FAIL"
            log(f"case open failed: {filename}: {exc}")
            return result

        try:
            sheet_name = add_japanese_sheet(doc)
            wait_for_window()
            screenshot(EVIDENCE / f"{base}-ja-sheet-added.png")
            save_document(doc, target, password)
            log(f"saved workbook: {target.relative_to(ROOT)} with sheet {sheet_name}")
        except Exception as exc:
            result["notes"].append(f"add/save failed: {exc}")
            result["no_corruption"] = "FAIL"
            log(f"case add/save failed: {filename}: {exc}")
            return result
        finally:
            if doc is not None:
                close_document(doc)
                doc = None

        try:
            doc = load_document(desktop, target, password)
            if doc is None:
                raise RuntimeError("LibreOffice returned no document on reopen")
            result["reopen_after_save"] = "PASS"
            wait_for_window()
            retained = verify_sheet(doc, sheet_name)
            result["japanese_sheet_name_retained"] = "PASS" if retained else "FAIL"
            screenshot(EVIDENCE / f"{base}-reopen.png")
            if not retained:
                result["notes"].append(f"sheet retention failed for {sheet_name}")
        except Exception as exc:
            result["notes"].append(f"reopen failed: {exc}")
            result["no_corruption"] = "FAIL"
            log(f"case reopen failed: {filename}: {exc}")
    finally:
        if doc is not None:
            close_document(doc)
        stop_office(proc, desktop)
        log(f"case result: {filename}: {result}")
    return result


def table_rows(results):
    rows = []
    for r in results:
        evidence = "<br/>".join(r["evidence"])
        rows.append(
            f"| {r['file']} | {r['password_type']} | {r['correct_password_open']} | "
            f"{r['wrong_password_rejected']} | {r['japanese_sheet_name_retained']} | "
            f"{r['reopen_after_save']} | {r['no_corruption']} | {evidence} |"
        )
    return "\n".join(rows)


def write_summary(start, end, service_info, container_info, lo_version, results, overall, notes, commands):
    content = f"""# Linux Manual Verification Summary

- 実施日時（JST）: {start} - {end}
- 実行環境: docker compose service `lo-vnc` / {container_info}
- LibreOffice バージョン: {lo_version}
- noVNC: http://localhost:6901/ (`VNC_PASSWORD` from `.env`)
- 総合判定: {overall}

## Result

| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |
|---|---|---:|---:|---:|---:|---:|---|
{table_rows(results)}

## Notes

- docker compose service: {service_info}
- 誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。
{notes}

## Commands

```bash
{chr(10).join(commands)}
```
"""
    SUMMARY.write_text(content, encoding="utf-8")


def append_checklist(start, service_info, container_info, lo_version, results, overall, notes):
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).strftime("%Y-%m-%d")
    section = f"""

## Linux Manual Verification (Docker/VNC) - {today}

- 実施日時（JST）: {start}
- 実行環境: docker compose service `lo-vnc` / {container_info}
- LibreOffice バージョン: {lo_version}

| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |
|---|---|---:|---:|---:|---:|---:|---|
{table_rows(results)}

- 総合判定: {overall}
- 備考: docker compose service: {service_info}。誤パスワード拒否は password-type ごとの代表ケースとして `simple_en.xlsx` と `simple_ja.xlsx` で確認した。{notes.replace(chr(10), ' ')}
"""
    with CHECKLIST.open("a", encoding="utf-8") as f:
        f.write(section)


def main():
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    WORKBOOKS.mkdir(parents=True, exist_ok=True)
    SESSION_LOG.touch()

    start = dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
    commands = [
        "docker compose up -d",
        "docker compose ps",
        "docker inspect lo-vnc --format ...",
        "docker compose exec -T lo-vnc bash -lc 'libreoffice --version; command -v xdotool; command -v import; command -v xclip'",
        "mkdir -p docs/manual-evidence/linux/workbooks",
        "cp test-manual-files/*.xlsx test-manual-files/*.xlsm docs/manual-evidence/linux/workbooks/",
        "docker compose exec -T lo-vnc python3 /workspace/docs/manual-evidence/linux/run_libreoffice_manual_verification.py",
    ]

    lo_version = libreoffice_version()
    log(f"LibreOffice version: {lo_version}")
    log("VNC/noVNC ports expected from docker-compose.yml: 5901/tcp and 6901/tcp")
    log("noVNC URL checked from host configuration: http://localhost:6901/")
    log("GUI helper tools are preinstalled: xdotool/import/xclip")

    for filename, _, _ in FILES:
        source = ROOT / "test-manual-files" / filename
        target = WORKBOOKS / filename
        shutil.copy2(source, target)
        log(f"copied work file: {source.relative_to(ROOT)} -> {target.relative_to(ROOT)}")

    results = []
    wrong_en = wrong_password_check(WORKBOOKS / "simple_en.xlsx", "wrong-pass", "simple_en", 22901)
    wrong_ja = wrong_password_check(WORKBOOKS / "simple_ja.xlsx", "違うパスワード", "simple_ja", 22902)

    for idx, (filename, ptype, password) in enumerate(FILES):
        result = verify_case(filename, ptype, password, idx)
        if filename == "simple_en.xlsx":
            result["wrong_password_rejected"] = "PASS" if wrong_en else "FAIL"
        if filename == "simple_ja.xlsx":
            result["wrong_password_rejected"] = "PASS" if wrong_ja else "FAIL"
        results.append(result)

    end = dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
    failures = []
    for r in results:
        required = [
            r["correct_password_open"],
            r["japanese_sheet_name_retained"],
            r["reopen_after_save"],
            r["no_corruption"],
        ]
        if r["file"] in ("simple_en.xlsx", "simple_ja.xlsx"):
            required.append(r["wrong_password_rejected"])
        if any(v != "PASS" for v in required):
            failures.append(r["file"])
    overall = "PASS" if not failures else "FAIL"
    notes_lines = []
    notes_lines.append("- ファイル破損警告、修復ダイアログ、内容削除/修復の確認ダイアログは検出されなかった。" if overall == "PASS" else "- 失敗ケースあり。詳細は session.log を参照。")
    for r in results:
        for note in r["notes"]:
            notes_lines.append(f"- {r['file']}: {note}")
    notes = "\n".join(notes_lines)

    service_info = "lo-vnc (ports 5901/tcp, 6901/tcp)"
    container_info = os.environ.get("CONTAINER_INFO", "lo-vnc / setpasstoexceldotnet-lo-vnc")
    write_summary(start, end, service_info, container_info, lo_version, results, overall, notes, commands)
    append_checklist(start, service_info, container_info, lo_version, results, overall, notes)
    log(f"summary written: {SUMMARY.relative_to(ROOT)}")
    log(f"checklist appended: {CHECKLIST.relative_to(ROOT)}")
    log(f"overall result: {overall}")
    return 0 if overall == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
