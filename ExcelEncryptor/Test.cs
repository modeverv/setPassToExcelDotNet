using System;
using System.IO;
using System.Linq;

namespace ExcelEncryptor;

public static class Test
{
    public static bool Check(
        string outputPath,
        string password,
        string testFilePath,
        string originalPath
    )
    {
        Console.WriteLine("\n=== decryption test ===");
        var checkResult = TestDecryption(outputPath, password, "dotnet version");
        if (!checkResult) return checkResult;
        checkResult = TestDecryption(testFilePath, password, "poi version");
        if (!checkResult) return checkResult;
        Console.WriteLine("\n=== compare with original and poi version ===");
        checkResult = CompareWithOriginal(originalPath, outputPath, testFilePath, password);
        return checkResult;
    }


    private static bool TestDecryption(string encryptedPath, string password, string label)
    {
        try
        {
            Console.WriteLine($"\n{label} decrypt...");
            var decrypted = Encrypt.Decrypt(encryptedPath, password);

            Console.WriteLine($"  size: {decrypted.Length} bytes");
            Console.WriteLine(
                $"  first 16bytes: {BitConverter.ToString(decrypted, 0, Math.Min(16, decrypted.Length)).Replace("-", " ")}");

            if (decrypted.Length >= 2 && decrypted[0] == 0x50 && decrypted[1] == 0x4B)
                Console.WriteLine("  ✓ ok ZIP file（check PK signature）");
            else
                Console.WriteLine("  ✗ no ZIP signature");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ✗ error: {ex.Message}");
            return false;
        }

        return true;
    }

    private static bool CompareWithOriginal(string originalPath, string dotnetPath, string poiPath, string password)
    {
        try
        {
            var original = File.ReadAllBytes(originalPath);
            var dotnetDecrypted = Encrypt.Decrypt(dotnetPath, password);
            var poiDecrypted = Encrypt.Decrypt(poiPath, password);

            Console.WriteLine($"original: {original.Length} bytes");
            Console.WriteLine($"dotnet: {dotnetDecrypted.Length} bytes");
            Console.WriteLine($"poi: {poiDecrypted.Length} bytes");

            var dotnetMatch = CompareBytes(original, dotnetDecrypted);
            var poiMatch = CompareBytes(original, poiDecrypted);

            Console.WriteLine($"\ndotnet and original: {(dotnetMatch ? "✓ same" : "✗ not same")}");
            Console.WriteLine($"poi and original: {(poiMatch ? "✓ same" : "✗ not same")}");
            if (!(dotnetMatch && poiMatch)) return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"compare error: {ex.Message}");
            return false;
        }

        return true;
    }

    private static bool CompareBytes(byte[] a, byte[] b)
    {
        if (a.Length != b.Length) return false;
        return !a.Where((t, i) => t != b[i]).Any();
    }
}
