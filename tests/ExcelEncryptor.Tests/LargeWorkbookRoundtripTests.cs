using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace ExcelEncryptor.Tests;

[CollectionDefinition("LargeWorkbookSerial", DisableParallelization = true)]
public sealed class LargeWorkbookSerialCollection
{
}

[Collection("LargeWorkbookSerial")]
public class LargeWorkbookRoundtripTests(ITestOutputHelper output)
{
    private const string Password = "pass";

    [Theory]
    [MemberData(nameof(LargeWorkbookSizesMb))]
    public void EncryptDecrypt_LargeWorkbook_RoundtripMatchesAndWrongPasswordFails(int sizeMb)
    {
        var originalPath = CreateDeterministicLargeWorkbook(sizeMb);
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-large-{sizeMb}mb-{Guid.NewGuid():N}.xlsx");

        try
        {
            var originalBytes = File.ReadAllBytes(originalPath);
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);

            var encryptWatch = Stopwatch.StartNew();
            encryptor.EncryptFile(originalPath, encryptedPath, Password);
            encryptWatch.Stop();

            var decryptWatch = Stopwatch.StartNew();
            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
            decryptWatch.Stop();

            output.WriteLine($"{sizeMb,4} MB | encrypt: {encryptWatch.ElapsedMilliseconds,6} ms | decrypt: {decryptWatch.ElapsedMilliseconds,6} ms");

            Assert.Equal(originalBytes, decryptedBytes);

            var wrongPwEx = Assert.Throws<UnauthorizedAccessException>(() => Encrypt.Decrypt(encryptedPath, "wrong_password"));
            Assert.Contains("Invalid password", wrongPwEx.Message);
        }
        finally
        {
            DeleteIfExists(originalPath);
            DeleteIfExists(encryptedPath);
        }
    }

    public static IEnumerable<object[]> LargeWorkbookSizesMb()
    {
        yield return new object[] { 1 };
        yield return new object[] { 10 };
        yield return new object[] { 50 };
        yield return new object[] { 100 };
    }

    private static string CreateDeterministicLargeWorkbook(int sizeMb)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelencryptor-large-vector-{sizeMb}mb-{Guid.NewGuid():N}.xlsx");
        var targetBytes = sizeMb * 1024L * 1024L;

        using (var fileStream = File.Create(path))
        using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, leaveOpen: false))
        {
            WriteZipEntry(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                "  <Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>\n" +
                "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
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
                "</Relationships>\n");

            WriteZipEntry(archive, "xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                "  <sheetData>\n" +
                "    <row r=\"1\">\n" +
                "      <c r=\"A1\" t=\"inlineStr\"><is><t>LargeWorkbookFixture</t></is></c>\n" +
                "    </row>\n" +
                "  </sheetData>\n" +
                "</worksheet>\n");

            var payloadBytes = Math.Max(0, targetBytes - fileStream.Position - 4096);
            WriteDeterministicPayload(archive, "xl/media/payload.bin", payloadBytes);
        }

        var finalLength = new FileInfo(path).Length;
        if (finalLength < targetBytes)
        {
            var deficit = targetBytes - finalLength + 4096;
            using var fileStream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var archive = new ZipArchive(fileStream, ZipArchiveMode.Update, leaveOpen: false);
            WriteDeterministicPayload(archive, "xl/media/payload2.bin", deficit);
        }

        return path;
    }

    private static void WriteZipEntry(ZipArchive archive, string entryName, string content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Encoding.UTF8, leaveOpen: false);
        writer.Write(content);
    }

    private static void WriteDeterministicPayload(ZipArchive archive, string entryName, long byteCount)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        var block = new byte[8192];
        for (var i = 0; i < block.Length; i++)
            block[i] = (byte)(i % 251);

        using var stream = entry.Open();
        long written = 0;

        while (written < byteCount)
        {
            var toWrite = (int)Math.Min(block.Length, byteCount - written);
            stream.Write(block, 0, toWrite);
            written += toWrite;
        }
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }
}


