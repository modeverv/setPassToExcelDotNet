# ExcelEncryptor

ExcelEncryptor is a small, focused, free and commercially usable .NET library for applying Excel-compatible Agile Encryption to OOXML workbooks.

Password-protect and decrypt `.xlsx` / `.xlsm` files using MS-OFFCRYPTO Agile Encryption.

Tested for compatibility with Apache POI, Microsoft Excel, and LibreOffice Calc.

Works with any library that produces `.xlsx` bytes: ClosedXML, NPOI, DocumentFormat.OpenXml, or a plain
`File.ReadAllBytes`.

## Commercial use and license

- Commercial use is allowed under `Apache-2.0`.
- This project is intended to remain free for both personal and commercial use.
- There is no plan to introduce a paid commercial license for this library.
- See `LICENSE` and [security.md](https://github.com/modeverv/setPassToExcelDotNet/blob/master/docs/security.md) for scope and limitations.

## Installation

```
dotnet add package ExcelEncryptor
```

## Dependencies

- [OpenMcdf](https://github.com/ironfede/openmcdf) — Compound File Binary (CFB) read/write

## Encryption

### From byte array

```csharp
byte[] xlsxBytes = File.ReadAllBytes("input.xlsx");

ExcelEncryptor.Encrypt.FromBytesToFile(xlsxBytes, "output.xlsx", "password");
```

### From file

```csharp
ExcelEncryptor.Encrypt.FromFileToFile("input.xlsx", "output.xlsx", "password");
```

### Low-level API (custom algorithm)

```csharp
var encryptor = new ExcelEncryptor.Encrypt(
    keySize:       AesKeySize.Aes256,
    hashAlgorithm: HashAlgorithmType.Sha512
);

encryptor.EncryptFile("input.xlsx", "output.xlsx", "password");
```

Supported key sizes: `Aes128` (default), `Aes192`, `Aes256`  
Supported hash algorithms: `Sha1` (default), `Sha256`, `Sha384`, `Sha512`, `Md5` (legacy compatibility only)

### Recommended and legacy settings

- Recommended for new files: `Aes256` + `Sha512`
- Legacy compatibility: `Sha1` and `Md5` are kept for compatibility with older environments
- Security note: `Md5` is not recommended for new encryption

## Decryption

```csharp
// Decrypt to byte array
byte[] xlsxBytes = ExcelEncryptor.Encrypt.Decrypt("encrypted.xlsx", "password");

// Decrypt to file
ExcelEncryptor.Encrypt.DecryptToFile("encrypted.xlsx", "decrypted.xlsx", "password");
```

---

## For ClosedXML users

No helper class needed. Save the workbook to a `MemoryStream` and pass the bytes directly.

```csharp
var wb = new XLWorkbook();
wb.AddWorksheet("Sheet1").Cell("A1").Value = "Hello";

using var ms = new MemoryStream();
wb.SaveAs(ms);

ExcelEncryptor.Encrypt.FromBytesToFile(ms.ToArray(), "output.xlsx", "password");
```

---

## For NPOI users

The library itself has no dependency on NPOI. If you use NPOI, the following helper stream lets you pipe
`IWorkbook.Write()` output directly into encryption without a temporary file.

Copy this class into your project:

```csharp
using System.IO;

/// <summary>
/// Drop-in stream for NPOI's IWorkbook.Write() that encrypts on Close().
/// Copy into your project — not part of the ExcelEncryptor package itself.
/// </summary>
public class NpoiXlsxPasswordFileOutputStream : Stream
{
    private readonly MemoryStream _buffer = new();
    private readonly string _outputPath;
    private readonly string _password;

    public NpoiXlsxPasswordFileOutputStream(string outputPath, string password)
    {
        _outputPath = outputPath;
        _password   = password;
    }

    public override bool CanRead  => false;
    public override bool CanSeek  => true;
    public override bool CanWrite => true;
    public override long Length   => _buffer.Length;

    public override long Position
    {
        get => _buffer.Position;
        set => _buffer.Position = value;
    }

    public override void Write(byte[] buffer, int offset, int count)
        => _buffer.Write(buffer, offset, count);

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => 0;

    public override long Seek(long offset, SeekOrigin origin)
        => _buffer.Seek(offset, origin);

    public override void SetLength(long value)
        => _buffer.SetLength(value);

    public override void Close()
    {
        base.Close();
        ExcelEncryptor.Encrypt.FromBytesToFile(_buffer.ToArray(), _outputPath, _password);
    }
}
```

Usage:

```csharp
IWorkbook wb = new XSSFWorkbook();
wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0).SetCellValue("Hello");

using var stream = new NpoiXlsxPasswordFileOutputStream("output.xlsx", "password");
wb.Write(stream);
```

---

## Compatibility


ExcelEncryptor is tested with:

- Manual open checks with Microsoft Excel for Mac
- Manual open checks with Microsoft Excel for Windows
- Manual open checks with LibreOffice Calc on Ubuntu via Docker/VNC
- Automated Apache POI interoperability tests
- Automated tests for `.xlsx` files containing formulas, styles, Japanese sheet names, and embedded images

For details, see [compatibility.md](https://github.com/modeverv/setPassToExcelDotNet/blob/master/docs/compatibility.md).

---


## Changelog

### v2.0.5.1

- Added an automated thread-safety test for concurrent encryption using a shared `Encrypt` instance.
- Documented that encryption behavior is verified under parallel execution.

### v2.0.5

- Added automated compatibility tests
- Added Apache POI interoperability tests
- Added image-containing workbook tests
- Added manual verification artifacts for Excel / LibreOffice checks

For details, see [tests](https://github.com/modeverv/setPassToExcelDotNet/tree/master/tests).
 
### v2.0.0

- Library is now NPOI-free — depends only on OpenMcdf
- `NpoiXlsxPasswordFileOutputStream` removed from package; available as a copy-paste snippet above
- Full backward compatibility on the encryption/decryption API

### v1.5.0

- Removed OpenMcdf dependency (reverted in v2.0.0)

### v1.0.0

- Initial release

## License

Apache-2.0
