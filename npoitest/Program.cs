using ExcelEncryptor;
using NPOI.XSSF.UserModel;

namespace npoitest;

internal static class Program
{
    private static void Main()
    {
        // ワークブックを作成
        var wb = new XSSFWorkbook();
        wb.CreateSheet("s1");

        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var outputPath = Path.Combine(projectDir, "sample_b.xlsx");

        using var ms = new MemoryStream();
        wb.Write(ms);
        var bytes = ms.ToArray();
        Encrypt.FromBytesToFile(bytes, outputPath, "pass");
    }
}
