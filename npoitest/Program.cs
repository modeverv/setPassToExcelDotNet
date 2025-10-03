using NPOI.XSSF.UserModel;

namespace npoitest;

internal class Program
{
    private static void Main(string[] args)
    {
        // ワークブックを作成
        var wb = new XSSFWorkbook();
        wb.CreateSheet("s1");

        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var outputPath = Path.Combine(projectDir, "sample_b.xlsx");

        using (var ms = new MemoryStream())
        {
            wb.Write(ms);
            var bytes = ms.ToArray();
            ExcelEncryptor.EncryptFromBytes(bytes, outputPath, "pass");
        }
    }
}
