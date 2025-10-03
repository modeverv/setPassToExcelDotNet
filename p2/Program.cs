namespace p2;

public static class Program
{
    private static void Main()
    {
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var inputPath = Path.Combine(projectDir, "a.xlsx");
        var outputPath = Path.Combine(projectDir, "b.xlsx");
        var testFilePath = Path.Combine(projectDir, "poi_b.xlsx");

        var password = "pass";

        try
        {
            ExcelEncryptor.Encrypt.FromFileToFile(inputPath, outputPath, password);
            Console.WriteLine($"暗号化完了: {outputPath}");

            Console.WriteLine("\n=== 復号化テスト ===");
            TestDecryption(outputPath, password, "dotnet版");
            TestDecryption(testFilePath, password, "poi版");

            Console.WriteLine("\n=== 元ファイルとの比較 ===");
            CompareWithOriginal(inputPath, outputPath, testFilePath, password);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static void TestDecryption(string encryptedPath, string password, string label)
    {
        try
        {
            Console.WriteLine($"\n{label} を復号化中...");
            var decrypted = ExcelEncryptor.Encrypt.Decrypt(encryptedPath, password);

            Console.WriteLine($"  復号化後サイズ: {decrypted.Length} bytes");
            Console.WriteLine(
                $"  最初の16バイト: {BitConverter.ToString(decrypted, 0, Math.Min(16, decrypted.Length)).Replace("-", " ")}");

            if (decrypted.Length >= 2 && decrypted[0] == 0x50 && decrypted[1] == 0x4B)
                Console.WriteLine("  ✓ 正常なZIPファイル（PKシグネチャ確認）");
            else
                Console.WriteLine("  ✗ ZIPシグネチャがありません！");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ✗ エラー: {ex.Message}");
        }
    }

    private static void CompareWithOriginal(string originalPath, string dotnetPath, string poiPath, string password)
    {
        try
        {
            var original = File.ReadAllBytes(originalPath);
            var dotnetDecrypted = ExcelEncryptor.Encrypt.Decrypt(dotnetPath, password);
            var poiDecrypted = ExcelEncryptor.Encrypt.Decrypt(poiPath, password);

            Console.WriteLine($"元ファイル: {original.Length} bytes");
            Console.WriteLine($"dotnet復号化: {dotnetDecrypted.Length} bytes");
            Console.WriteLine($"poi復号化: {poiDecrypted.Length} bytes");

            var dotnetMatch = CompareBytes(original, dotnetDecrypted);
            var poiMatch = CompareBytes(original, poiDecrypted);

            Console.WriteLine($"\ndotnet版と元ファイル: {(dotnetMatch ? "✓ 完全一致" : "✗ 不一致")}");
            Console.WriteLine($"poi版と元ファイル: {(poiMatch ? "✓ 完全一致" : "✗ 不一致")}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"比較エラー: {ex.Message}");
        }
    }

    private static bool CompareBytes(byte[] a, byte[] b)
    {
        if (a.Length != b.Length) return false;
        for (var i = 0; i < a.Length; i++)
            if (a[i] != b[i])
                return false;
        return true;
    }
}
