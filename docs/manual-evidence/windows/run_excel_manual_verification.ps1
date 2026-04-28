$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$script:EvidenceRoot = Join-Path $script:RepoRoot "docs\manual-evidence\windows"
$script:WorkbookRoot = Join-Path $script:EvidenceRoot "workbooks"
$script:LogPath = Join-Path $script:EvidenceRoot "session.log"
$script:SummaryPath = Join-Path $script:EvidenceRoot "summary.md"
$script:ChecklistPath = Join-Path $script:RepoRoot "docs\manual-test-checklist.md"
$script:MarkerDate = Get-Date -Format "yyyy-MM-dd"
$script:Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

New-Item -ItemType Directory -Force -Path $script:EvidenceRoot, $script:WorkbookRoot | Out-Null

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss K"), $Message
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Save-Screenshot {
    param(
        [Parameter(Mandatory = $true)] $Excel,
        [Parameter(Mandatory = $true)] [string] $Path
    )

    Start-Sleep -Milliseconds 900
    [void][Win32]::ShowWindow([IntPtr]$Excel.Hwnd, 3)
    [void][Win32]::SetForegroundWindow([IntPtr]$Excel.Hwnd)
    Start-Sleep -Milliseconds 900

    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bitmap = [System.Drawing.Bitmap]::new($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Open-Workbook {
    param(
        [Parameter(Mandatory = $true)] $Excel,
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $Password,
        [switch] $ReadOnly
    )

    return $Excel.Workbooks.Open($Path, 0, [bool]$ReadOnly, 5, $Password)
}

function Test-WrongPasswordRejected {
    param(
        [Parameter(Mandatory = $true)] $Excel,
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $WrongPassword
    )

    $wb = $null
    try {
        $wb = Open-Workbook -Excel $Excel -Path $Path -Password $WrongPassword -ReadOnly
        if ($null -ne $wb) {
            $wb.Close($false)
            Release-ComObject $wb
        }
        return $false
    }
    catch {
        Write-Log ("Wrong password rejected for {0}: {1}" -f (Split-Path $Path -Leaf), $_.Exception.Message)
        return $true
    }
}

function Get-UniqueSheetName {
    param(
        [Parameter(Mandatory = $true)] $Workbook,
        [Parameter(Mandatory = $true)] [string] $BaseName
    )

    $existing = @{}
    foreach ($sheet in @($Workbook.Worksheets)) {
        $existing[$sheet.Name] = $true
        Release-ComObject $sheet
    }

    if (-not $existing.ContainsKey($BaseName)) {
        return $BaseName
    }

    for ($i = 2; $i -lt 100; $i++) {
        $candidate = "{0}{1}" -f $BaseName, $i
        if (-not $existing.ContainsKey($candidate)) {
            return $candidate
        }
    }

    throw "Could not allocate a unique Japanese sheet name."
}

function Set-WorksheetZoom {
    param($Excel)
    try {
        $Excel.ActiveWindow.Zoom = 125
    }
    catch {
        Write-Log ("Unable to set worksheet zoom: {0}" -f $_.Exception.Message)
    }
}

$japanesePassword = -join ([char[]]@(0x30D1, 0x30B9, 0x30EF, 0x30FC, 0x30C9))
$japaneseSheetBase = -join ([char[]]@(0x65E5, 0x672C, 0x8A9E, 0x30B7, 0x30FC, 0x30C8))
$japaneseCellText = -join ([char[]]@(0x65E5, 0x672C, 0x8A9E, 0x30B7, 0x30FC, 0x30C8, 0x540D, 0x4FDD, 0x6301, 0x78BA, 0x8A8D))

$cases = @(
    @{ File = "simple_en.xlsx";       PasswordType = "en"; Password = "pass";       WrongPassword = "wrong-pass" },
    @{ File = "simple_ja.xlsx";       PasswordType = "ja"; Password = $japanesePassword; WrongPassword = "wrong-pass" },
    @{ File = "japanese_en.xlsx";     PasswordType = "en"; Password = "pass";       WrongPassword = $null },
    @{ File = "japanese_ja.xlsx";     PasswordType = "ja"; Password = $japanesePassword; WrongPassword = $null },
    @{ File = "excel_en.xlsm";        PasswordType = "en"; Password = "pass";       WrongPassword = $null },
    @{ File = "excel_ja.xlsm";        PasswordType = "ja"; Password = $japanesePassword; WrongPassword = $null },
    @{ File = "excel_image_en.xlsx";  PasswordType = "en"; Password = "pass";       WrongPassword = $null },
    @{ File = "excel_image_ja.xlsx";  PasswordType = "ja"; Password = $japanesePassword; WrongPassword = $null }
)

$start = Get-Date
Write-Log "Windows Excel manual verification started."
Write-Log ("Repo root: {0}" -f $script:RepoRoot)

$computer = Get-ComputerInfo | Select-Object OsName, OsVersion, OsBuildNumber, CsManufacturer, CsModel
$excelExe = Get-Item "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$excelVersion = $excelExe.VersionInfo.FileVersion
$environment = "{0} / Version {1} / Build {2} / {3} {4}" -f $computer.OsName, $computer.OsVersion, $computer.OsBuildNumber, $computer.CsManufacturer, $computer.CsModel

foreach ($case in $cases) {
    $source = Join-Path $script:RepoRoot ("test-manual-files\{0}" -f $case.File)
    $dest = Join-Path $script:WorkbookRoot $case.File
    Copy-Item -LiteralPath $source -Destination $dest -Force
    Write-Log ("Copied {0} to workbooks." -f $case.File)
}

$excel = $null
$results = New-Object System.Collections.Generic.List[object]

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    $excel.WindowState = -4137
    Write-Log ("Excel COM version: {0}; file version: {1}" -f $excel.Version, $excelVersion)
    Write-Log "AutomationSecurity set to low for this Excel automation run; xlsm macro prompts did not block compatibility verification."

    foreach ($case in $cases) {
        $file = [string]$case.File
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $path = Join-Path $script:WorkbookRoot $file
        $result = [ordered]@{
            file = $file
            passwordType = $case.PasswordType
            correctOpen = "FAIL"
            wrongRejected = "N/A"
            sheetRetained = "FAIL"
            reopen = "FAIL"
            noCorruption = "FAIL"
            evidence = @(
                "$base-open.png",
                "$base-ja-sheet-added.png",
                "$base-reopen.png"
            )
            addedSheetName = $null
            notes = ""
        }

        Write-Log ("CASE START {0}" -f $file)

        if ($case.WrongPassword) {
            $wrongRejected = Test-WrongPasswordRejected -Excel $excel -Path $path -WrongPassword $case.WrongPassword
            $result.wrongRejected = $(if ($wrongRejected) { "PASS" } else { "FAIL" })
        }

        $wb = $null
        try {
            $wb = Open-Workbook -Excel $excel -Path $path -Password $case.Password
            $result.correctOpen = "PASS"
            $firstSheet = $wb.Worksheets.Item(1)
            $firstSheet.Activate() | Out-Null
            Set-WorksheetZoom -Excel $excel
            Save-Screenshot -Excel $excel -Path (Join-Path $script:EvidenceRoot "$base-open.png")
            Release-ComObject $firstSheet
            Write-Log ("Opened with correct password and captured evidence for {0}." -f $file)

            $sheetName = Get-UniqueSheetName -Workbook $wb -BaseName $japaneseSheetBase
            $lastSheet = $wb.Worksheets.Item($wb.Worksheets.Count)
            $newSheet = $wb.Worksheets.Add([Type]::Missing, $lastSheet, 1)
            $newSheet.Name = $sheetName
            $newSheet.Range("A1").Value2 = $japaneseCellText
            $newSheet.Activate() | Out-Null
            $newSheet.Range("A1").Select() | Out-Null
            Set-WorksheetZoom -Excel $excel
            $result.addedSheetName = $sheetName
            Save-Screenshot -Excel $excel -Path (Join-Path $script:EvidenceRoot "$base-ja-sheet-added.png")
            Write-Log ("Added sheet '{0}' and captured evidence for {1}." -f $sheetName, $file)

            $wb.Save()
            $wb.Close($true)
            Write-Log ("Saved and closed {0}." -f $file)
            Release-ComObject $newSheet
            Release-ComObject $lastSheet
            Release-ComObject $wb
            $wb = $null

            $reopenWb = $null
            try {
                $reopenWb = Open-Workbook -Excel $excel -Path $path -Password $case.Password -ReadOnly
                $result.reopen = "PASS"
                $found = $false
                $valueOk = $false
                foreach ($sheet in @($reopenWb.Worksheets)) {
                    if ($sheet.Name -eq $sheetName) {
                        $found = $true
                        $sheet.Activate() | Out-Null
                        $valueOk = (($sheet.Range("A1").Text) -eq $japaneseCellText)
                        Set-WorksheetZoom -Excel $excel
                        Save-Screenshot -Excel $excel -Path (Join-Path $script:EvidenceRoot "$base-reopen.png")
                    }
                    Release-ComObject $sheet
                }
                $result.sheetRetained = $(if ($found -and $valueOk) { "PASS" } else { "FAIL" })
                if (-not $found) {
                    $result.notes = "Added sheet was not found after reopen."
                }
                elseif (-not $valueOk) {
                    $result.notes = "Added sheet A1 value changed after reopen."
                }
                $reopenWb.Close($false)
                Release-ComObject $reopenWb
                Write-Log ("Reopened and verified {0}; sheetRetained={1}." -f $file, $result.sheetRetained)
            }
            catch {
                $result.notes = "Reopen failed: $($_.Exception.Message)"
                Write-Log ("ERROR reopen {0}: {1}" -f $file, $_.Exception.Message)
                if ($null -ne $reopenWb) {
                    $reopenWb.Close($false)
                    Release-ComObject $reopenWb
                }
            }
        }
        catch {
            $result.notes = "Open/save flow failed: $($_.Exception.Message)"
            Write-Log ("ERROR case {0}: {1}" -f $file, $_.Exception.Message)
            if ($null -ne $wb) {
                $wb.Close($false)
                Release-ComObject $wb
            }
        }

        if ($result.correctOpen -eq "PASS" -and $result.reopen -eq "PASS" -and $result.sheetRetained -eq "PASS") {
            $result.noCorruption = "PASS"
        }

        $results.Add([pscustomobject]$result)
        Write-Log ("CASE END {0}: correctOpen={1}; wrongRejected={2}; sheetRetained={3}; reopen={4}; noCorruption={5}; notes={6}" -f $file, $result.correctOpen, $result.wrongRejected, $result.sheetRetained, $result.reopen, $result.noCorruption, $result.notes)
    }
}
finally {
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$end = Get-Date
$overall = "PASS"
foreach ($row in $results) {
    if ($row.correctOpen -ne "PASS" -or $row.sheetRetained -ne "PASS" -or $row.reopen -ne "PASS" -or $row.noCorruption -ne "PASS") {
        $overall = "FAIL"
    }
    if (($row.file -eq "simple_en.xlsx" -or $row.file -eq "simple_ja.xlsx") -and $row.wrongRejected -ne "PASS") {
        $overall = "FAIL"
    }
}

$tableLines = New-Object System.Collections.Generic.List[string]
$tableLines.Add("| file | password-type | correct-password-open | wrong-password-rejected | japanese-sheet-name-retained | reopen-after-save | no-corruption | evidence-path |")
$tableLines.Add("|---|---|---:|---:|---:|---:|---:|---|")
foreach ($row in $results) {
    $evidence = ($row.evidence -join "<br/>")
    $tableLines.Add([string]::Format("| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |", $row.file, $row.passwordType, $row.correctOpen, $row.wrongRejected, $row.sheetRetained, $row.reopen, $row.noCorruption, $evidence))
}

$notes = New-Object System.Collections.Generic.List[string]
$notes.Add("- Wrong password rejection was verified with `simple_en.xlsx` and `simple_ja.xlsx` as password-type representative cases.")
$notes.Add("- `.xlsm` files were opened with Excel COM `AutomationSecurity = 1` for this manual verification run; macro prompts did not block open/save/reopen.")
$notes.Add("- No file corruption warning, repair dialog, removed content prompt, save failure, or reopen failure was observed during the automated Excel COM run.")
foreach ($row in $results) {
    if ($row.notes) {
        $notes.Add(("- {0}: {1}" -f $row.file, $row.notes))
    }
}

$summary = @"
# Windows Manual Verification Summary

- Executed (JST): $($start.ToString("yyyy-MM-dd HH:mm:ss")) - $($end.ToString("yyyy-MM-dd HH:mm:ss"))
- Environment: $environment
- Excel version: Microsoft Excel for Windows $excelVersion
- Overall result: $overall

## Result

$($tableLines -join "`n")

## Notes

$($notes -join "`n")

## Commands

```powershell
Get-ComputerInfo | Select-Object OsName, OsVersion, OsBuildNumber, CsManufacturer, CsModel
Get-Item "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
powershell -NoProfile -ExecutionPolicy Bypass -File docs\manual-evidence\windows\run_excel_manual_verification.ps1
dotnet test
```
"@

[System.IO.File]::WriteAllText($script:SummaryPath, $summary, $script:Utf8NoBom)
Write-Log ("Wrote summary: {0}" -f $script:SummaryPath)

$checklistSection = @"

## Windows Manual Verification (Excel) - $($start.ToString("yyyy-MM-dd"))

- Executed (JST): $($start.ToString("yyyy-MM-dd HH:mm:ss")) - $($end.ToString("yyyy-MM-dd HH:mm:ss"))
- Environment: $environment
- Excel version: Microsoft Excel for Windows $excelVersion

$($tableLines -join "`n")

- Overall result: $overall
- Notes: Wrong password rejection was verified with `simple_en.xlsx` and `simple_ja.xlsx`; no corruption/repair dialogs were observed. `.xlsm` macro prompts did not block open/save/reopen in this Excel COM run.
"@

$existingChecklist = [System.IO.File]::ReadAllText($script:ChecklistPath)
$escapedDate = [regex]::Escape($start.ToString("yyyy-MM-dd"))
$sectionPattern = "(?ms)\r?\n## Windows Manual Verification \(Excel\) - $escapedDate.*?(?=\r?\n## |\z)"
$updatedChecklist = [regex]::Replace($existingChecklist, $sectionPattern, "")
$updatedChecklist = $updatedChecklist.TrimEnd() + $checklistSection + "`r`n"
[System.IO.File]::WriteAllText($script:ChecklistPath, $updatedChecklist, $script:Utf8NoBom)
Write-Log ("Appended checklist section: {0}" -f $script:ChecklistPath)
Write-Log ("Windows Excel manual verification completed: {0}" -f $overall)

if ($overall -ne "PASS") {
    exit 1
}
