# ExcelEncryptor

Password-protect and decrypt `.xlsx` / `.xlsm` files in .NET — compatible with Apache POI and Microsoft Excel.

> Inspired by [Apache POI](https://github.com/apache/poi)'s Agile encryption implementation.

## Features

- AES-128 / 192 / 256 encryption
- Multiple hash algorithms: MD5, SHA-1, SHA-256, SHA-384, SHA-512
- Compatible with Microsoft Excel password-protected files
- Cross-compatible with Java POI encrypted files
- No dependency on NPOI or ClosedXML — works with any xlsx source

## Repository structure

```
ExcelEncryptor/   — library (NuGet package)
ProjectForTest/   — round-trip compatibility test against POI
```

## Usage

See [ExcelEncryptor/README.md](ExcelEncryptor/README.md) for full API documentation and examples.

## Dependencies

- [OpenMcdf](https://github.com/ironfede/openmcdf) — Compound File Binary (CFB) read/write

## Testing

Run `ProjectForTest/Program.cs` to verify encryption/decryption round-trip compatibility with POI.

```
dotnet run --project ProjectForTest

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

## License

Apache-2.0