using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.XSSF.UserModel;
using Xunit;

namespace ExcelEncryptor.Interop.Tests;

public class GeneratorInteropTests
{
    private const string Password = "pass";
    private static readonly object BuildLock = new();
    private static string? _checkerJarPath;
    private static string? _buildFailureReason;

    [Theory]
    [MemberData(nameof(WorkbookScenarios))]
    public void Encrypt_GeneratedWorkbook_CanBeDecryptedByApachePoi(WorkbookScenario scenario)
    {
        var root = FindRepositoryRoot();
        var plainPath = CreateTempPath($"excelencryptor-{scenario.Name}-plain");
        var encryptedPath = CreateTempPath($"excelencryptor-{scenario.Name}-encrypted");
        var decryptedPath = CreateTempPath($"excelencryptor-{scenario.Name}-decrypted");

        try
        {
            scenario.CreateWorkbook(plainPath);
            scenario.ValidateWorkbook(plainPath);

            var originalBytes = File.ReadAllBytes(plainPath);
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(plainPath, encryptedPath, Password);

            if (!TryDecryptWithPoi(root, encryptedPath, decryptedPath, Password, out var reason))
            {
                if (IsPoiInteropRequired())
                    Assert.Fail(reason);

                return;
            }

            Assert.Equal(originalBytes, File.ReadAllBytes(decryptedPath));
            scenario.ValidateWorkbook(decryptedPath);
        }
        finally
        {
            DeleteIfExists(plainPath);
            DeleteIfExists(encryptedPath);
            DeleteIfExists(decryptedPath);
        }
    }

    [Theory]
    [MemberData(nameof(WorkbookScenarios))]
    public void Encrypt_GeneratedWorkbook_WithWrongPassword_DoesNotSucceed(WorkbookScenario scenario)
    {
        var root = FindRepositoryRoot();
        var plainPath = CreateTempPath($"excelencryptor-{scenario.Name}-plain-wrong");
        var encryptedPath = CreateTempPath($"excelencryptor-{scenario.Name}-encrypted-wrong");
        var wrongDecryptedPath = CreateTempPath($"excelencryptor-{scenario.Name}-wrong-decrypted");

        try
        {
            scenario.CreateWorkbook(plainPath);
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(plainPath, encryptedPath, Password);

            var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
            Assert.Contains("Invalid password", ex.Message);

            if (TryDecryptWithPoi(root, encryptedPath, wrongDecryptedPath, "wrong_password", out var reason))
                Assert.Fail($"Apache POI unexpectedly decrypted the workbook with a wrong password. {reason}");
        }
        finally
        {
            DeleteIfExists(plainPath);
            DeleteIfExists(encryptedPath);
            DeleteIfExists(wrongDecryptedPath);
        }
    }

    public static IEnumerable<object[]> WorkbookScenarios()
    {
        yield return new object[] { new WorkbookScenario("closedxml", CreateClosedXmlWorkbook, ValidateClosedXmlWorkbook) };
        yield return new object[] { new WorkbookScenario("npoi", CreateNpoiWorkbook, ValidateNpoiWorkbook) };
        yield return new object[] { new WorkbookScenario("openxml", CreateOpenXmlWorkbook, ValidateOpenXmlWorkbook) };
    }

    public sealed record WorkbookScenario(string Name, Action<string> CreateWorkbook, Action<string> ValidateWorkbook);

    private static void CreateClosedXmlWorkbook(string path)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell("A1").Value = "ClosedXML";
        worksheet.Cell("B1").Value = 123;
        worksheet.Cell("A2").Value = "日本語";
        workbook.SaveAs(path);
    }

    private static void ValidateClosedXmlWorkbook(string path)
    {
        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal("ClosedXML", worksheet.Cell("A1").GetString());
        Assert.Equal(123, worksheet.Cell("B1").GetValue<int>());
        Assert.Equal("日本語", worksheet.Cell("A2").GetString());
    }

    private static void CreateNpoiWorkbook(string path)
    {
        var workbook = new XSSFWorkbook();
        var worksheet = workbook.CreateSheet("Sheet1");

        var row1 = worksheet.CreateRow(0);
        row1.CreateCell(0).SetCellValue("NPOI");
        row1.CreateCell(1).SetCellValue(123);
        var row2 = worksheet.CreateRow(1);
        row2.CreateCell(0).SetCellValue("日本語");

        using var stream = File.Create(path);
        workbook.Write(stream);
    }

    private static void ValidateNpoiWorkbook(string path)
    {
        using var stream = File.OpenRead(path);
        var workbook = new XSSFWorkbook(stream);
        var worksheet = workbook.GetSheetAt(0);

        Assert.Equal("NPOI", worksheet.GetRow(0)!.GetCell(0)!.StringCellValue);
        Assert.Equal(123, (int)worksheet.GetRow(0)!.GetCell(1)!.NumericCellValue);
        Assert.Equal("日本語", worksheet.GetRow(1)!.GetCell(0)!.StringCellValue);
    }

    private static void CreateOpenXmlWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetData(
                new Row(
                    CreateInlineStringCell("A1", "OpenXML SDK"),
                    CreateNumberCell("B1", 123)),
                new Row(
                    CreateInlineStringCell("A2", "日本語"))));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        workbookPart.Workbook.Save();
    }

    private static void ValidateOpenXmlWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        Assert.Equal("OpenXML SDK", GetOpenXmlCellText(document, "Sheet1", "A1"));
        Assert.Equal("123", GetOpenXmlCellText(document, "Sheet1", "B1"));
        Assert.Equal("日本語", GetOpenXmlCellText(document, "Sheet1", "A2"));
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

    private static string GetOpenXmlCellText(SpreadsheetDocument document, string sheetName, string cellReference)
    {
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
        var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .SingleOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var cell = worksheetPart.Worksheet.Descendants<Cell>()
            .Single(item => string.Equals(item.CellReference?.Value, cellReference, StringComparison.Ordinal));

        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString)
            return GetSharedStringText(workbookPart, cell);

        if (dataType == CellValues.InlineString)
            return cell.InnerText;

        return cell.CellValue?.Text ?? cell.InnerText;
    }

    private static string GetSharedStringText(WorkbookPart workbookPart, Cell cell)
    {
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable ?? throw new InvalidOperationException("Shared string table is missing.");
        if (!int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
            throw new InvalidOperationException("Shared string index is invalid.");

        return sharedStrings.ElementAt(index).InnerText;
    }

    private static bool TryDecryptWithPoi(string root, string encryptedPath, string outputPath, string password, out string reason)
    {
        if (!TryEnsurePoiChecker(root, out var jarPath, out reason))
            return false;

        var args = $"-jar \"{jarPath}\" decrypt \"{encryptedPath}\" \"{outputPath}\" \"{password}\"";
        return TryRunProcess("java", args, root, out reason);
    }

    private static bool TryEnsurePoiChecker(string root, out string jarPath, out string reason)
    {
        lock (BuildLock)
        {
            if (_checkerJarPath != null)
            {
                jarPath = _checkerJarPath;
                reason = string.Empty;
                return true;
            }

            if (_buildFailureReason != null)
            {
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            var checkerDir = Path.Combine(root, "tests", "poi-decrypt-checker");
            var pomPath = Path.Combine(checkerDir, "pom.xml");
            var buildArgs = $"-q -f \"{pomPath}\" -DskipTests package";

            if (!TryRunProcess("mvn", buildArgs, root, out reason))
            {
                _buildFailureReason = $"POI checker build skipped: {reason}";
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            var jarCandidate = Path.Combine(checkerDir, "target", "poi-decrypt-checker-1.0.0-jar-with-dependencies.jar");
            if (!File.Exists(jarCandidate))
            {
                _buildFailureReason = $"POI checker jar was not produced: {jarCandidate}";
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            _checkerJarPath = jarCandidate;
            jarPath = _checkerJarPath;
            reason = string.Empty;
            return true;
        }
    }

    private static bool TryRunProcess(string fileName, string arguments, string workingDirectory, out string reason)
    {
        try
        {
            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            
            process.Start();
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                reason = $"Command failed: {fileName} {arguments}\n{stdout}\n{stderr}";
                return false;
            }

            reason = string.Empty;
            return true;
        }
        catch (Win32Exception ex)
        {
            reason = $"Command not available: {fileName} ({ex.Message})";
            return false;
        }
        catch (Exception ex)
        {
            reason = $"Command execution error: {fileName} ({ex.Message})";
            return false;
        }
    }

    private static bool IsPoiInteropRequired()
    {
        return string.Equals(Environment.GetEnvironmentVariable("POI_INTEROP_REQUIRED"), "1", StringComparison.Ordinal);
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

