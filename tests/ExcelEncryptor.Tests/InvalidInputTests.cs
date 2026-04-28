using Xunit;

namespace ExcelEncryptor.Tests;

public class InvalidInputTests
{
    private const string Password = "pass";

    [Fact]
    public void EncryptFile_WithNonXlsxInput_ThrowsInvalidOperationException()
    {
        var inputPath = CreateTextFile("not-an-xlsx");
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-invalid-input-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            var ex = Assert.Throws<InvalidOperationException>(() => encryptor.EncryptFile(inputPath, encryptedPath, Password));
            Assert.Contains("valid OOXML workbook", ex.Message);
        }
        finally
        {
            DeleteIfExists(inputPath);
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void EncryptFile_WithCorruptedXlsx_ThrowsInvalidOperationException()
    {
        var inputPath = CreateCorruptedWorkbook();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-corrupt-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            var ex = Assert.Throws<InvalidOperationException>(() => encryptor.EncryptFile(inputPath, encryptedPath, Password));
            Assert.Contains("valid OOXML workbook", ex.Message);
        }
        finally
        {
            DeleteIfExists(inputPath);
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void EncryptFile_WithAlreadyEncryptedInput_ThrowsInvalidOperationException()
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "plain", "a.xlsx");
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-already-encrypted-{Guid.NewGuid():N}.xlsx");
        var reencryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-reencrypted-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(plainPath, encryptedPath, Password);

            var ex = Assert.Throws<InvalidOperationException>(() => encryptor.EncryptFile(encryptedPath, reencryptedPath, Password));
            Assert.Contains("valid OOXML workbook", ex.Message);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
            DeleteIfExists(reencryptedPath);
        }
    }
    
    [Fact]
    public void DecryptToStream_WithNullInputStream_ThrowsArgumentNullException()
    {
        var output = new MemoryStream();

        try
        {
            Assert.Throws<ArgumentNullException>(() => Encrypt.DecryptToStream(null!, output, Password));
        }
        finally
        {
            output.Dispose();
        }
    }

    [Fact]
    public void DecryptToStream_WithUnreadableInputStream_ThrowsArgumentException()
    {
        var output = new MemoryStream();

        try
        {
            using var input = Stream.Null;
            var ex = Assert.Throws<ArgumentException>(() => Encrypt.DecryptToStream(input, output, Password));
            Assert.Equal("encryptedStream", ex.ParamName);
        }
        finally
        {
            output.Dispose();
        }
    }

    [Fact]
    public void DecryptToStream_WithUnwritableOutputStream_ThrowsArgumentException()
    {
        var encryptedBytes = CreateEncryptedBytes();
        using var input = new MemoryStream(encryptedBytes, writable: false);
        using var output = new MemoryStream(new byte[0], writable: false);

        var ex = Assert.Throws<ArgumentException>(() => Encrypt.DecryptToStream(input, output, Password));
        Assert.Equal("outputStream", ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidParameterCases))]
    public void Encrypt_WithInvalidParameters_ThrowsInvalidOperationException(AesKeySize keySize, HashAlgorithmType hashAlgorithm)
    {
        var ex = Assert.Throws<InvalidOperationException>(() => new Encrypt(keySize, hashAlgorithm));
        Assert.Contains("Invalid", ex.Message);
    }

    public static IEnumerable<object[]> InvalidParameterCases()
    {
        yield return new object[] { (AesKeySize)0, HashAlgorithmType.Sha1 };
        yield return new object[] { AesKeySize.Aes128, (HashAlgorithmType)0 };
    }

    private static byte[] CreateEncryptedBytes()
    {
        var inputPath = GetPlainPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-stream-encrypted-source-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(inputPath, encryptedPath, Password);
            return File.ReadAllBytes(encryptedPath);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    private static string CreateTextFile(string content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelencryptor-text-{Guid.NewGuid():N}.txt");
        File.WriteAllText(path, content);
        return path;
    }

    private static string CreateCorruptedWorkbook()
    {
        var sourcePath = GetPlainPath();
        var targetPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-corrupt-source-{Guid.NewGuid():N}.xlsx");
        var bytes = File.ReadAllBytes(sourcePath);
        var truncated = new byte[Math.Max(1, bytes.Length / 4)];
        Array.Copy(bytes, truncated, truncated.Length);
        File.WriteAllBytes(targetPath, truncated);
        return targetPath;
    }

    private static string GetPlainPath()
    {
        var root = FindRepositoryRoot();
        return Path.Combine(root, "test-vectors", "plain", "a.xlsx");
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

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

}


