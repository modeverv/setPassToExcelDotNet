using ExcelEncryptor;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace p2;

public static class Program
{
    private static void Main()
    {
        TestEncrypt();
        TestNPoi();        
    }
    
    private static void TestEncrypt()
    {
        var password = "pass";
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var inputPath = Path.Combine(projectDir, "a.xlsx");
        var outputPath = "";
        var decryptedFile = Path.Combine(projectDir, "decrypted.xlsx");
        var testFilePath = Path.Combine(projectDir, "poi_b.xlsx");
        
        // ========================================
        // pattern1: AES-128 + SHA-1
        // ========================================
        var encryptor1 = new Encrypt();
        outputPath = Path.Combine(projectDir, "encrypted_aes128_sha1.xlsx");
        encryptor1.EncryptFile(
            inputPath,
            outputPath,
            password
        );
        var check = Test.Check(
            outputPath, password, testFilePath, inputPath
        );
        if (!check) throw new InvalidProgramException("process check error");
        
        // ========================================
        // pattern2: AES-256 + SHA-256
        // ========================================
        var encryptor2 = new Encrypt(
            AesKeySize.Aes256,
            HashAlgorithmType.Sha256
        );
        outputPath = Path.Combine(projectDir, "encrypted_aes256_sha256.xlsx");
        encryptor2.EncryptFile(
            inputPath,
            outputPath,
            password
        );
        check = Test.Check(
            outputPath, password, testFilePath, inputPath
        );
        if (!check) throw new InvalidProgramException("process check error");
        Encrypt.DecryptToFile(outputPath, decryptedFile, password);
        
        // ========================================
        // pattern3: AES-192 + SHA-384
        // ========================================
        var encryptor3 = new Encrypt(
            AesKeySize.Aes192,
            HashAlgorithmType.Sha384
        );
        outputPath = Path.Combine(projectDir, "encrypted_aes384_sha384.xlsx");
        encryptor3.EncryptFile(
            inputPath,
            outputPath,
            password
        );
        check = Test.Check(
            outputPath, password, testFilePath, inputPath
        );
        if (!check) throw new InvalidProgramException("process check error");
        
        // ========================================
        // pattern4: AES-256 + SHA-512
        // ========================================
        var encryptor4 = new Encrypt(
            AesKeySize.Aes256,
            HashAlgorithmType.Sha512
        );
        outputPath = Path.Combine(projectDir, "encrypted_aes256_sha512.xlsx");
        encryptor4.EncryptFile(
            inputPath,
            outputPath,
            password
        );
        check = Test.Check(
            outputPath, password, testFilePath, inputPath
        );
        if (!check) throw new InvalidProgramException("process check error");
        
        // ========================================
        // pattern5: MD5
        // ========================================
        var encryptor5 = new Encrypt(
            AesKeySize.Aes128,
            HashAlgorithmType.Md5
        );
        outputPath = Path.Combine(projectDir, "encrypted_aes128_md5.xlsx");
        encryptor5.EncryptFile(
            inputPath,
            outputPath,
            password
        );
        check = Test.Check(
            outputPath, password, testFilePath, inputPath
        );
        if (!check) throw new InvalidProgramException("process check error");
        
        try
        {
            Console.WriteLine("例1: DecryptToFile - ファイルに直接復号化");
            Encrypt.DecryptToFile(outputPath, decryptedFile, password);
            Console.WriteLine($"  ✓ 復号化完了: {decryptedFile}");
            Console.WriteLine($"  ファイルサイズ: {new FileInfo(decryptedFile).Length} bytes\n");
            
            // ========================================
            // 例2: メモリに復号化してから処理
            // ========================================
            Console.WriteLine("例2: Decrypt - メモリに復号化");
            byte[] decryptedData = Encrypt.Decrypt(outputPath, password);
            Console.WriteLine($"  ✓ 復号化完了");
            Console.WriteLine($"  データサイズ: {decryptedData.Length} bytes");
            
            // 必要に応じてファイルに保存
            string outputFile = "decrypted_from_memory.xlsx";
            File.WriteAllBytes(outputFile, decryptedData);
            Console.WriteLine($"  ✓ 保存完了: {outputFile}\n");
            
            // ========================================
            // 例3: エラーハンドリング
            // ========================================
            Console.WriteLine("例3: エラーハンドリング");
            
            // 間違ったパスワード
            try
            {
                Console.Write("  間違ったパスワードでテスト... ");
                Encrypt.Decrypt(outputPath, "wrong_password");
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("✓ 正しくUnauthorizedAccessExceptionが発生");
            }
            
            // 存在しないファイル
            try
            {
                Console.Write("  存在しないファイルでテスト... ");
                Encrypt.Decrypt("nonexistent.xlsx", password);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("✓ 正しくFileNotFoundExceptionが発生");
            }
            
            Console.WriteLine("\n=== すべてのテストが成功しました！ ===");
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine($"\nエラー: パスワードが間違っています");
            Console.WriteLine($"詳細: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"\nエラー: ファイルが見つかりません");
            Console.WriteLine($"詳細: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"\nエラー: 復号化に失敗しました");
            Console.WriteLine($"詳細: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n予期しないエラー: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
    
    
    private static void TestNPoi()
    {
        IWorkbook wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");
        
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var outPath = Path.Combine(projectDir, "protected.xlsx");
        
        using var outStream = new NpoiXlsxPasswordFileOutputStream(outPath, "pa");
        wb.Write(outStream);
        Console.WriteLine("output end.");
    }
}
