# WorkbookSizeBenchmark

Simple benchmark runner for large workbook encryption/decryption.

## Run

```bash
dotnet run --project samples/WorkbookSizeBenchmark/WorkbookSizeBenchmark.csproj
```

Optional sizes (MB):

```bash
dotnet run --project samples/WorkbookSizeBenchmark/WorkbookSizeBenchmark.csproj -- 1 10 50 100
```

The benchmark prints encryption/decryption elapsed milliseconds and a byte-to-byte roundtrip match result.

