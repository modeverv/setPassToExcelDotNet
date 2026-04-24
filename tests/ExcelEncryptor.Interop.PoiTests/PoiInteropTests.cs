using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using ExcelEncryptor;
using Xunit;

namespace ExcelEncryptor.Interop.PoiTests;

public class PoiInteropTests
{
    private const string Password = "pass";
    private static readonly object BuildLock = new();
    private static string? _checkerJarPath;
    private static string? _buildFailureReason;

    [Fact]
    public void Encrypt_WithAes256Sha512_CanBeDecryptedByApachePoi()
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "plain", "a.xlsx");
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-poi-encrypted-{Guid.NewGuid():N}.xlsx");
        var poiDecryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-poi-decrypted-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(plainPath, encryptedPath, Password);

            if (!TryDecryptWithPoi(root, encryptedPath, poiDecryptedPath, Password, out var reason))
            {
                if (IsPoiInteropRequired())
                    Assert.Fail(reason);

                return;
            }

            Assert.Equal(File.ReadAllBytes(plainPath), File.ReadAllBytes(poiDecryptedPath));
        }
        finally
        {
            DeleteIfExists(encryptedPath);
            DeleteIfExists(poiDecryptedPath);
        }
    }

    [Fact]
    public void Decrypt_FileEncryptedByApachePoi_ReturnsOriginalPackage()
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "plain", "a.xlsx");
        var poiEncryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", "poi_b.xlsx");

        var decrypted = Encrypt.Decrypt(poiEncryptedPath, Password);
        Assert.Equal(File.ReadAllBytes(plainPath), decrypted);
    }

    [Fact]
    public void Decrypt_FileEncryptedByApachePoi_WithWrongPassword_ThrowsUnauthorizedAccessException()
    {
        var root = FindRepositoryRoot();
        var poiEncryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", "poi_b.xlsx");

        var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(poiEncryptedPath, "wrong_password"));
        Assert.Contains("Invalid password", ex.Message);
    }

    [Theory]
    [InlineData("simple")]
    [InlineData("formulas")]
    [InlineData("styles")]
    [InlineData("japanese")]
    public void Decrypt_PoiAes256Sha512TestVectors_ReturnsOriginalBytes(string name)
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "plain", $"{name}.xlsx");
        var encryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", $"{name}_aes256_sha512.xlsx");

        var decrypted = Encrypt.Decrypt(encryptedPath, Password);
        Assert.Equal(File.ReadAllBytes(plainPath), decrypted);
    }

    [Theory]
    [InlineData("simple")]
    [InlineData("formulas")]
    [InlineData("styles")]
    [InlineData("japanese")]
    public void Decrypt_PoiAes256Sha512TestVectors_WithWrongPassword_ThrowsUnauthorizedAccessException(string name)
    {
        var root = FindRepositoryRoot();
        var encryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", $"{name}_aes256_sha512.xlsx");

        var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
        Assert.Contains("Invalid password", ex.Message);
    }

    [Fact]
    public void Encrypt_ImageXlsx_CanBeDecryptedByApachePoi()
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "image", "image.xlsx");
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-poi-image-enc-{Guid.NewGuid():N}.xlsx");
        var poiDecryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-poi-image-dec-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(plainPath, encryptedPath, Password);

            if (!TryDecryptWithPoi(root, encryptedPath, poiDecryptedPath, Password, out var reason))
            {
                if (IsPoiInteropRequired())
                    Assert.Fail(reason);
                return;
            }

            Assert.Equal(File.ReadAllBytes(plainPath), File.ReadAllBytes(poiDecryptedPath));
        }
        finally
        {
            DeleteIfExists(encryptedPath);
            DeleteIfExists(poiDecryptedPath);
        }
    }

    [Fact]
    public void Decrypt_PoiEncryptedImageXlsx_PreservesAllBytes()
    {
        var root = FindRepositoryRoot();
        var plainPath = Path.Combine(root, "test-vectors", "image", "image.xlsx");
        var encryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", "image_aes256_sha512.xlsx");

        var decrypted = Encrypt.Decrypt(encryptedPath, Password);
        Assert.Equal(File.ReadAllBytes(plainPath), decrypted);
    }

    [Fact]
    public void Decrypt_PoiEncryptedImageXlsx_WithWrongPassword_ThrowsUnauthorizedAccessException()
    {
        var root = FindRepositoryRoot();
        var encryptedPath = Path.Combine(root, "test-vectors", "encrypted-by-apache-poi", "image_aes256_sha512.xlsx");

        var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
        Assert.Contains("Invalid password", ex.Message);
    }

    private static bool TryDecryptWithPoi(string root, string encryptedPath, string outputPath, string password, out string reason)
    {
        if (!TryEnsurePoiChecker(root, out var jarPath, out reason))
            return false;

        var args = $"-jar \"{jarPath}\" decrypt \"{encryptedPath}\" \"{outputPath}\" \"{password}\"";
        return TryRunProcess("java", args, root, out reason);
    }

    private static bool TryEnsurePoiChecker(string root, out string jarPath, out string reason)
    {
        lock (BuildLock)
        {
            if (_checkerJarPath != null)
            {
                jarPath = _checkerJarPath;
                reason = string.Empty;
                return true;
            }

            if (_buildFailureReason != null)
            {
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            var checkerDir = Path.Combine(root, "tests", "poi-decrypt-checker");
            var pomPath = Path.Combine(checkerDir, "pom.xml");
            var buildArgs = $"-q -f \"{pomPath}\" -DskipTests package";

            if (!TryRunProcess("mvn", buildArgs, root, out reason))
            {
                _buildFailureReason = $"POI checker build skipped: {reason}";
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            var jarCandidate = Path.Combine(checkerDir, "target", "poi-decrypt-checker-1.0.0-jar-with-dependencies.jar");
            if (!File.Exists(jarCandidate))
            {
                _buildFailureReason = $"POI checker jar was not produced: {jarCandidate}";
                jarPath = string.Empty;
                reason = _buildFailureReason;
                return false;
            }

            _checkerJarPath = jarCandidate;
            jarPath = _checkerJarPath;
            reason = string.Empty;
            return true;
        }
    }

    private static bool TryRunProcess(string fileName, string arguments, string workingDirectory, out string reason)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    WorkingDirectory = workingDirectory,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                reason = $"Command failed: {fileName} {arguments}\n{stdout}\n{stderr}";
                return false;
            }

            reason = string.Empty;
            return true;
        }
        catch (Win32Exception ex)
        {
            reason = $"Command not available: {fileName} ({ex.Message})";
            return false;
        }
        catch (Exception ex)
        {
            reason = $"Command execution error: {fileName} ({ex.Message})";
            return false;
        }
    }

    private static bool IsPoiInteropRequired()
    {
        return string.Equals(Environment.GetEnvironmentVariable("POI_INTEROP_REQUIRED"), "1", StringComparison.Ordinal);
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


