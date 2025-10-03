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

## with npoi

```csharp
IWorkbook wb = new XSSFWorkbook();
var sheet = wb.CreateSheet("Sheet1");
sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");

var outPath = "/path/to/output.xlsx";

using var outStream = new NpoiXlsxPasswordFileOutputStream(outPath, "pa");
wb.Write(outStream); 
```

# LICENSE

Apache-2.0

## dependency

- OpenMcdf https://github.com/ironfede/openmcdf
    - License MPL-2.0 https://github.com/ironfede/openmcdf?tab=MPL-2.0-1-ov-file#readme 

