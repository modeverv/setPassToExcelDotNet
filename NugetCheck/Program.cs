using ExcelEncryptor;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NugetCheck;

internal abstract class Program
{
    static void Main(string[] args)
    {
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var inputPath = Path.Combine(projectDir, "a.xlsx");
        var outputPath = Path.Combine(projectDir, "b.xlsx");
        var testFilePath = Path.Combine(projectDir, "poi_b.xlsx");

        const string password = "pass";

        try
        {
            Encrypt.FromFileToFile(inputPath, outputPath, password);
            Console.WriteLine($"encryption: {outputPath}");
            var check = Test.Check(
                outputPath, password, testFilePath, inputPath
            );
            if (!check) throw new InvalidProgramException("process check error");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }


        IWorkbook wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");

        var outPath = Path.Combine(projectDir, "protected.xlsx");

        using var outStream = new NpoiXlsxPasswordFileOutputStream(outPath, "pa");
        wb.Write(outStream);
        Console.WriteLine("output end.");

    }
}
