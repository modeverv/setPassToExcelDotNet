# Compatibility Notes

This document summarizes the compatibility coverage currently verified in this repository.

## Scope


- Target formats: `.xlsx`, `.xlsm`
- Verification methods:
  - Automated: test cases executed by `dotnet test`
  - Manual: GUI verification tracked in `docs/manual-test-checklist.md`

## Compatibility Matrix

| Environment | Result |
|---|---|
| Microsoft Excel for Mac | Passed |
| Microsoft Excel for Windows | Passed |
| LibreOffice Calc on Ubuntu Docker/VNC | Passed |
| Apache POI decrypt compatibility | Passed |
| Image embedded workbook | Passed |

## Format Coverage

| Format | Status | Evidence |
|---|---|---|
| `.xlsx` | Automated + Manual checklist available | `tests/ExcelEncryptor.Tests/RoundtripTests.cs`, `tests/ExcelEncryptor.Interop.PoiTests/PoiInteropTests.cs`, `docs/manual-test-checklist.md` |
| `.xlsm` | Automated (byte-to-byte roundtrip) + Manual checklist available | `tests/ExcelEncryptor.Tests/RoundtripTests.cs`, `tests/ExcelEncryptor.Tests/ManualFileGenerationTests.cs`, `docs/manual-test-checklist.md` |

## Automated Evidence

- Apache POI interoperability:
  - `tests/ExcelEncryptor.Interop.PoiTests/PoiInteropTests.cs`
  - `tests/poi-decrypt-checker/pom.xml` (`poi.version` = `5.2.5`)
- Generator interoperability:
  - `tests/ExcelEncryptor.Interop.Tests/GeneratorInteropTests.cs`
  - Versions from `tests/ExcelEncryptor.Interop.Tests/ExcelEncryptor.Interop.Tests.csproj`
- Roundtrip and wrong-password behavior:
  - `tests/ExcelEncryptor.Tests/RoundtripTests.cs`
  - `tests/ExcelEncryptor.Tests/InvalidInputTests.cs`

## Known Limitations

- GUI checks for Excel for Windows / Mac / LibreOffice are recorded per release in `docs/manual-test-checklist.md`. Environments without a record are not treated as verified.
- Apache POI interop tests depend on Java and Maven. If `POI_INTEROP_REQUIRED=1` is not set, tests may skip strict failure when the environment is missing.
- The supported workbook scope is OOXML (`.xlsx`, `.xlsm`). Other formats such as `.xls` and `.xlsb` are out of scope.

