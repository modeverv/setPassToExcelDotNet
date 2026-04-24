using System;
using System.Collections.Generic;
using System.IO;
using ExcelEncryptor;
using Xunit;

namespace ExcelEncryptor.Tests;

public class RoundtripTests
{
    private const string Password = "pass";

    [Fact]
    public void EncryptThenDecrypt_ReturnsOriginalPackage()
    {
        var originalPath = GetTestVectorPath();
        var originalBytes = File.ReadAllBytes(originalPath);
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-roundtrip-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
            Assert.Equal(originalBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void Decrypt_WithCorrectPassword_Succeeds()
    {
        var originalPath = GetTestVectorPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-correct-pw-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
            Assert.NotEmpty(decryptedBytes);
            Assert.True(decryptedBytes.Length > 1 && decryptedBytes[0] == 0x50 && decryptedBytes[1] == 0x4B);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void Decrypt_WithWrongPassword_ThrowsUnauthorizedAccessException()
    {
        var originalPath = GetTestVectorPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-wrong-pw-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
            Assert.Contains("Invalid password", ex.Message);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Theory]
    [MemberData(nameof(BoundaryPasswordCases))]
    public void EncryptDecrypt_WithBoundaryPasswords_Succeeds(string _, string password)
    {
        var originalPath = GetTestVectorPath();
        var originalBytes = File.ReadAllBytes(originalPath);
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-boundary-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, password);
            Assert.Equal(originalBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Theory]
    [MemberData(nameof(BoundaryPasswordCases))]
    public void Decrypt_WithWrongPassword_ForBoundaryPasswords_ThrowsUnauthorizedAccessException(string _, string password)
    {
        var originalPath = GetTestVectorPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-boundary-wrong-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, password);

            var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, password + "_wrong"));
            Assert.Contains("Invalid password", ex.Message);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void Encrypt_WithEmptyPassword_BehavesAsDocumented()
    {
        var originalPath = GetTestVectorPath();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-empty-pw-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            var ex = Assert.Throws<ArgumentException>(() => encryptor.EncryptFile(originalPath, encryptedPath, string.Empty));
            Assert.Equal("password", ex.ParamName);

            encryptor.EncryptFile(originalPath, encryptedPath, Password);
            var decryptEx = Assert.Throws<ArgumentException>(() => Encrypt.Decrypt(encryptedPath, string.Empty));
            Assert.Equal("password", decryptEx.ParamName);
        }
        finally
        {
            DeleteIfExists(encryptedPath);
        }
    }

    public static IEnumerable<object[]> BoundaryPasswordCases()
    {
        yield return new object[] { "long-255", new string('a', 255) };
        yield return new object[] { "japanese", "\u65e5\u672c\u8a9e\u30d1\u30b9\u30ef\u30fc\u30c9" };
        yield return new object[] { "emoji", "\ud83d\ude03\ud83d\ude80\ud83d\udd12" };
        yield return new object[] { "symbols", "!@#$%^&*()_+-=[]{}|;':\",./<>?`~" };
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

