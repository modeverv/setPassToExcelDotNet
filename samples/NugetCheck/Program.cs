using ExcelEncryptor;

namespace NugetCheck;

internal abstract class Program
{
    private static void Main(string[] args)
    {
        var projectDir = FindRepositoryRoot();
        var inputPath = Path.Combine(projectDir, "test-vectors", "plain", "a.xlsx");
        var outputPath = Path.Combine(projectDir, "b.xlsx");
        var testFilePath = Path.Combine(projectDir, "test-vectors", "encrypted-by-apache-poi", "poi_b.xlsx");
        
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

    private static string FindRepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);

        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SetPassToExceldotNet.sln")))
                return dir.FullName;

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Repository root could not be located from sample runtime directory.");
    }
}