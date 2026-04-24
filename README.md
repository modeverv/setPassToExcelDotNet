# ExcelEncryptor

Password-protect and decrypt `.xlsx` / `.xlsm` files in .NET — compatible with Apache POI and Microsoft Excel.

> Inspired by [Apache POI](https://github.com/apache/poi)'s Agile encryption implementation.

## Features

- AES-128 / 192 / 256 encryption
- Multiple hash algorithms: MD5, SHA-1, SHA-256, SHA-384, SHA-512
- Compatible with Microsoft Excel password-protected files
- Cross-compatible with Java POI encrypted files
- No dependency on NPOI or ClosedXML — works with any xlsx source

## Security defaults

- Recommended for new encryption: `Aes256` + `Sha512`
- Compatibility-only legacy options: `Sha1`, `Md5`
- `Md5` is supported only for compatibility and is not recommended for new files

## Repository structure

```
src/ExcelEncryptor/          — library (NuGet package)
tests/ExcelEncryptor.Tests/  — automated tests
samples/ProjectForTest/      — round-trip compatibility sample against POI
samples/WorkbookSizeBenchmark/ — large workbook benchmark runner
test-vectors/                — deterministic workbook fixtures
```

## Usage

See [src/ExcelEncryptor/README.md](src/ExcelEncryptor/README.md) for full API documentation and examples.

## Dependencies

- [OpenMcdf](https://github.com/ironfede/openmcdf) — Compound File Binary (CFB) read/write

## Testing

Run `samples/ProjectForTest/Program.cs` to verify encryption/decryption round-trip compatibility with POI.

```
dotnet run --project samples/ProjectForTest/ProjectForTest.csproj

=== 復号化テスト ===

dotnet版 を復号化中...
  復号化後サイズ: 6819 bytes
  ✓ 正常なZIPファイル（PKシグネチャ確認）

poi版 を復号化中...
  復号化後サイズ: 6819 bytes
  ✓ 正常なZIPファイル（PKシグネチャ確認）

dotnet版と元ファイル: ✓ 完全一致
poi版と元ファイル:    ✓ 完全一致
```

Validated workbook sizes in automated tests: `1 MB`, `10 MB`, `50 MB`, `100 MB`.

To measure local performance for large files, run:

```bash
dotnet run --project samples/WorkbookSizeBenchmark/WorkbookSizeBenchmark.csproj -- 1 10 50 100
```

## License

Apache-2.0