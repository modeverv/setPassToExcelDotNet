using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.XSSF.UserModel;
using Xunit;

namespace ExcelEncryptor.Interop.Tests;

public class ReadmeSamplesSmokeTests
{
    private const string Password = "pass";

    [Fact]
    public void MinimalSample_FromFileToFile_AndDecrypt_Works()
    {
        var root = FindRepositoryRoot();
        var inputPath = Path.Combine(root, "test-vectors", "plain", "simple.xlsx");
        var encryptedPath = CreateTempPath("readme-minimal-encrypted");

        try
        {
            Encrypt.FromFileToFile(inputPath, encryptedPath, Password);
            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);

            Assert.Equal(File.ReadAllBytes(inputPath), decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void ClosedXmlSample_FromBytesToFile_Works()
    {
        var encryptedPath = CreateTempPath("readme-closedxml-encrypted");

        try
        {
            byte[] workbookBytes;
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Sheet1").Cell("A1").Value = "Hello";
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                workbookBytes = stream.ToArray();
            }

            Encrypt.FromBytesToFile(workbookBytes, encryptedPath, Password);
            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);

            Assert.Equal(workbookBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void NpoiSample_FromBytesToFile_Works()
    {
        var encryptedPath = CreateTempPath("readme-npoi-encrypted");

        try
        {
            byte[] workbookBytes;
            using (var workbook = new XSSFWorkbook())
            {
                workbook.CreateSheet("Sheet1").CreateRow(0).CreateCell(0).SetCellValue("Hello");
                using var stream = new MemoryStream();
                workbook.Write(stream);
                workbookBytes = stream.ToArray();
            }

            Encrypt.FromBytesToFile(workbookBytes, encryptedPath, Password);
            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);

            Assert.Equal(workbookBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void OpenXmlSdkSample_FromBytesToFile_Works()
    {
        var encryptedPath = CreateTempPath("readme-openxml-encrypted");

        try
        {
            var workbookBytes = CreateOpenXmlWorkbookBytes();

            Encrypt.FromBytesToFile(workbookBytes, encryptedPath, Password);
            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);

            Assert.Equal(workbookBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    private static byte[] CreateOpenXmlWorkbookBytes()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(
                new SheetData(
                    new Row(
                        CreateInlineStringCell("A1", "OpenXML SDK"),
                        CreateNumberCell("B1", 123))));

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            });

            workbookPart.Workbook.Save();
        }

        return stream.ToArray();
    }

    private static Cell CreateInlineStringCell(string reference, string text)
    {
        return new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(text))
        };
    }

    private static Cell CreateNumberCell(string reference, int value)
    {
        return new Cell
        {
            CellReference = reference,
            DataType = CellValues.Number,
            CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture))
        };
    }

    private static string FindRepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);

        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SetPassToExceldotNet.sln")))
                return dir.FullName;

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Repository root could not be located from test runtime directory.");
    }

    private static string CreateTempPath(string prefix)
    {
        return Path.Combine(Path.GetTempPath(), $"{prefix}-{Guid.NewGuid():N}.xlsx");
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }
}


