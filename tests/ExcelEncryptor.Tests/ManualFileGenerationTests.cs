using Xunit;

namespace ExcelEncryptor.Tests;

public class ManualFileGenerationTests
{
    private const string Password = "pass";

    [Fact]
    public void Generate_ManualTestFiles_WritesEncryptedWorkbooks()
    {
        var root = FindRepositoryRoot();
        var outputDirectory = Path.Combine(root, "test-manual-files");
        var encryptor = new ExcelEncryptor.Encrypt(ExcelEncryptor.AesKeySize.Aes256, ExcelEncryptor.HashAlgorithmType.Sha512);

        Directory.CreateDirectory(outputDirectory);

        foreach (var scenario in GetScenarios())
        {
            var inputPath = Path.Combine(root, scenario.InputRelativePath);
            var outputPath = Path.Combine(outputDirectory, scenario.OutputFileName);

            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Test vector not found.", inputPath);

            encryptor.EncryptFile(inputPath, outputPath, Password);

            Assert.True(File.Exists(outputPath));
            Assert.Equal(File.ReadAllBytes(inputPath), ExcelEncryptor.Encrypt.Decrypt(outputPath, Password));
        }
    }

    private static IEnumerable<(string InputRelativePath, string OutputFileName)> GetScenarios()
    {
        yield return (Path.Combine("test-vectors", "plain", "simple.xlsx"), "simple_en.xlsx");
        yield return (Path.Combine("test-vectors", "plain", "japanese.xlsx"), "japanese_en.xlsx");
        yield return (Path.Combine("test-vectors", "xlsm", "excel_sample.xlsm"), "excel_en.xlsm");
        yield return (Path.Combine("test-vectors", "image", "image.xlsx"), "excel_image_en.xlsx");

        yield return (Path.Combine("test-vectors", "plain", "simple.xlsx"), "simple_ja.xlsx");
        yield return (Path.Combine("test-vectors", "plain", "japanese.xlsx"), "japanese_ja.xlsx");
        yield return (Path.Combine("test-vectors", "xlsm", "excel_sample.xlsm"), "excel_ja.xlsm");
        yield return (Path.Combine("test-vectors", "image", "image.xlsx"), "excel_image_ja.xlsx");
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
}


