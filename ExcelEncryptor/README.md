# set password to xlsx

# usage

## from bytes to file

```csharp
var wb = new XSSFWorkbook();
wb.CreateSheet("s1");

var outputPath = "/path/to/output.xslx"

using var ms = new MemoryStream();
wb.Write(ms);
var bytes = ms.ToArray();

ExcelEncryptor.Encrypt.FromBytesToFile(bytes, outputPath, "password-string");

```
             
## from file to file

```csharp
var inputPath = "/path/to/input.xslx"
var outputPath = "/path/to/output.xslx"
ExcelEncryptor.Encrypt.FromFileToFile(inputPath, outputPath, "password-string");
```
            
# LICENSE

MPL-2.0
