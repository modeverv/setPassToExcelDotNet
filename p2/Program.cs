using ExcelEncryptor;

namespace p2;

public static class Program
{
    private static void Main()
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
    }
}
