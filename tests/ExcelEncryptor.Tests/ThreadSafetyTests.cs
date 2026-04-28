using System.Collections.Concurrent;
using Xunit;

namespace ExcelEncryptor.Tests;

public class ThreadSafetyTests
{
    private const string Password = "pass";

    [Fact]
    public void EncryptFile_WithSharedEncryptorInstance_IsThreadSafe()
    {
        var originalPath = GetTestVectorPath();
        var originalBytes = File.ReadAllBytes(originalPath);
        var outputFiles = new ConcurrentBag<string>();
        var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
        const int concurrentJobs = 16;

        try
        {
            Parallel.For(0, concurrentJobs, i =>
            {
                var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-thread-safe-{i}-{Guid.NewGuid():N}.xlsx");
                outputFiles.Add(encryptedPath);
                encryptor.EncryptFile(originalPath, encryptedPath, Password);
            });

            Assert.Equal(concurrentJobs, outputFiles.Count);

            foreach (var encryptedPath in outputFiles)
            {
                var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
                Assert.Equal(originalBytes, decryptedBytes);
            }
        }
        finally
        {
            foreach (var path in outputFiles)
                DeleteIfExists(path);
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
