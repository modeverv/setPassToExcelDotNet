using System.IO.Compression;
using System.Text;
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

    [Fact]
    public void EncryptDecrypt_Xlsm_PreservesVbaProjectBin()
    {
        AssertXlsmRoundtrip("xl/vbaProject.bin");
    }

    [Fact]
    public void EncryptDecrypt_Xlsm_PreservesContentTypes()
    {
        AssertXlsmRoundtrip("[Content_Types].xml");
    }

    [Fact]
    public void EncryptDecrypt_Xlsm_PreservesRelationships()
    {
        AssertXlsmRoundtrip("_rels/.rels", "xl/_rels/workbook.xml.rels");
    }

    [Fact]
    public void Decrypt_Xlsm_WithWrongPassword_ThrowsUnauthorizedAccessException()
    {
        var originalPath = CreateXlsmTestVector();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-xlsm-wrong-pw-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt();
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var ex = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
            Assert.Contains("Invalid password", ex.Message);
        }
        finally
        {
            DeleteIfExists(originalPath);
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

    private static void AssertXlsmRoundtrip(params string[] entryNames)
    {
        var originalPath = CreateXlsmTestVector();
        var originalBytes = File.ReadAllBytes(originalPath);
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-xlsm-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
            Assert.Equal(originalBytes, decryptedBytes);

            foreach (var entryName in entryNames)
                Assert.Equal(ReadZipEntry(originalBytes, entryName), ReadZipEntry(decryptedBytes, entryName));
        }
        finally
        {
            DeleteIfExists(originalPath);
            DeleteIfExists(encryptedPath);
        }
    }

    private static string CreateXlsmTestVector()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelencryptor-xlsm-vector-{Guid.NewGuid():N}.xlsm");

        using var fileStream = File.Create(path);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, leaveOpen: false);

        WriteZipEntry(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
            "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
            "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
            "  <Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>\n" +
            "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main+xml\"/>\n" +
            "  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" +
            "</Types>\n");

        WriteZipEntry(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
            "</Relationships>\n");

        WriteZipEntry(archive, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
            "  <sheets>\n" +
            "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
            "  </sheets>\n" +
            "</workbook>\n");

        WriteZipEntry(archive, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
            "  <Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>\n" +
            "</Relationships>\n");

        WriteZipEntry(archive, "xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
            "  <sheetData>\n" +
            "    <row r=\"1\">\n" +
            "      <c r=\"A1\" t=\"inlineStr\"><is><t>Macro fixture</t></is></c>\n" +
            "    </row>\n" +
            "  </sheetData>\n" +
            "</worksheet>\n");

        WriteZipEntry(archive, "xl/vbaProject.bin", new byte[]
        {
            (byte)'F', (byte)'A', (byte)'K', (byte)'E', (byte)'_', (byte)'V', (byte)'B', (byte)'A',
            (byte)'_', (byte)'P', (byte)'R', (byte)'O', (byte)'J', (byte)'E', (byte)'C', (byte)'T',
            0x00, 0x01, 0x02, 0x03,
            (byte)'m', (byte)'a', (byte)'c', (byte)'r', (byte)'o', (byte)'-', (byte)'f', (byte)'i',
            (byte)'x', (byte)'t', (byte)'u', (byte)'r', (byte)'e', 0x00
        });
        return path;
    }

    private static void WriteZipEntry(ZipArchive archive, string entryName, string content)
    {
        WriteZipEntry(archive, entryName, Encoding.UTF8.GetBytes(content));
    }

    private static void WriteZipEntry(ZipArchive archive, string entryName, byte[] content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        using var stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static byte[] ReadZipEntry(byte[] packageBytes, string entryName)
    {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry(entryName);

        if (entry == null)
            throw new InvalidOperationException($"ZIP entry not found: {entryName}");

        using var entryStream = entry.Open();
        using var memoryStream = new MemoryStream();
        entryStream.CopyTo(memoryStream);
        return memoryStream.ToArray();
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

