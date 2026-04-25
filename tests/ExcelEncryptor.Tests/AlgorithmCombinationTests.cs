using Xunit;

namespace ExcelEncryptor.Tests;

public class AlgorithmCombinationTests
{
    private const string Password = "pass";

    [Fact]
    public void EncryptDecrypt_WithAes128Sha1_Succeeds()
    {
        AssertRoundtrip(AesKeySize.Aes128, HashAlgorithmType.Sha1, Password);
    }

    [Fact]
    public void EncryptDecrypt_WithAes256Sha512_Succeeds()
    {
        AssertRoundtrip(AesKeySize.Aes256, HashAlgorithmType.Sha512, Password);
    }

    [Theory]
    [MemberData(nameof(SupportedCombinations))]
    public void EncryptDecrypt_AllSupportedCipherHashCombinations_Succeed(AesKeySize keySize, HashAlgorithmType hashAlgorithm)
    {
        AssertRoundtrip(keySize, hashAlgorithm, Password);
    }

    [Theory]
    [MemberData(nameof(SupportedCombinations))]
    public void Decrypt_WithWrongPassword_ForAllSupportedCipherHashCombinations_ThrowsUnauthorizedAccessException(
        AesKeySize keySize,
        HashAlgorithmType hashAlgorithm)
    {
        var originalPath = GetTestVectorPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-combo-wrong-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(keySize, hashAlgorithm);
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
            Assert.Contains("Invalid password", ex.Message);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    public static IEnumerable<object[]> SupportedCombinations()
    {
        var keySizes = new[] { AesKeySize.Aes128, AesKeySize.Aes192, AesKeySize.Aes256 };
        var hashAlgorithms = new[]
        {
            HashAlgorithmType.Sha1,
            HashAlgorithmType.Sha256,
            HashAlgorithmType.Sha384,
            HashAlgorithmType.Sha512
        };

        foreach (var keySize in keySizes)
        {
            foreach (var hashAlgorithm in hashAlgorithms)
            {
                yield return new object[] { keySize, hashAlgorithm };
            }
        }
    }

    private static void AssertRoundtrip(AesKeySize keySize, HashAlgorithmType hashAlgorithm, string password)
    {
        var originalPath = GetTestVectorPath();
        var originalBytes = File.ReadAllBytes(originalPath);
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-combo-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(keySize, hashAlgorithm);
            encryptor.EncryptFile(originalPath, encryptedPath, password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, password);
            Assert.Equal(originalBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    private static string GetTestVectorPath()
    {
        var root = FindRepositoryRoot();
        var path = Path.Combine(root, "test-vectors", "plain", "a.xlsx");

        if (!File.Exists(path))
            throw new FileNotFoundException("Test vector not found.", path);

        return path;
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

