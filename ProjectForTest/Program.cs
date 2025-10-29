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
        TestApi();
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
            Console.WriteLine("1: DecryptToFile");
            Encrypt.DecryptToFile(outputPath, decryptedFile, password);
            Console.WriteLine($"  ✓ decrypt complete: {decryptedFile}");
            Console.WriteLine($"  data size: {new FileInfo(decryptedFile).Length} bytes\n");
            
            Console.WriteLine("2: Decrypt");
            byte[] decryptedData = Encrypt.Decrypt(outputPath, password);
            Console.WriteLine($"  ✓ decrypt complete");
            Console.WriteLine($"  data size: {decryptedData.Length} bytes");
            
            // 必要に応じてファイルに保存
            string outputFile = "decrypted_from_memory.xlsx";
            File.WriteAllBytes(outputFile, decryptedData);
            Console.WriteLine($"  ✓ save complete: {outputFile}\n");
            
            Console.WriteLine("3: error handling");
            
            try
            {
                Console.Write("  test wrong password ");
                Encrypt.Decrypt(outputPath, "wrong_password");
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("✓ OK UnauthorizedAccessException");
            }
            
            try
            {
                Console.Write("  file not exist");
                Encrypt.Decrypt("nonexistent.xlsx", password);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("✓ OK FileNotFoundException");
            }
            
            Console.WriteLine("\n=== test OK ===");
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine($"\nerror : wrong password");
            Console.WriteLine($"detail: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"\nerror: file not found");
            Console.WriteLine($"detail: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"\nerror: decrypt fail");
            Console.WriteLine($"detail: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nerror: {ex.Message}");
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
    
    private static void TestApi()
    {
        IWorkbook wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");
        
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var outputPath = Path.Combine(projectDir, "protected.xlsx");
        var inputPath = Path.Combine(projectDir, "a.xlsx");        

        using var ms = new MemoryStream();
        wb.Write(ms);
        var bytes = ms.ToArray();
        ExcelEncryptor.Encrypt.FromBytesToFile(bytes, outputPath, "password-string");
        ExcelEncryptor.Encrypt.FromFileToFile(inputPath, outputPath, "password-string");

        using var outStream = new NpoiXlsxPasswordFileOutputStream(outputPath, "password-string");
        wb.Write(outStream); 
    }
}
