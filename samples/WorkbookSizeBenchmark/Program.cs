using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using ExcelEncryptor;

const string password = "pass";
var sizesMb = ParseSizes(args);

Console.WriteLine("ExcelEncryptor workbook size benchmark");
Console.WriteLine($"Sizes (MB): {string.Join(", ", sizesMb)}");
Console.WriteLine();

foreach (var sizeMb in sizesMb)
{
    var plainPath = CreateDeterministicWorkbook(sizeMb);
    var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-bench-{sizeMb}mb-{Guid.NewGuid():N}.xlsx");

    try
    {
        var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);

        var encryptWatch = Stopwatch.StartNew();
        encryptor.EncryptFile(plainPath, encryptedPath, password);
        encryptWatch.Stop();

        var decryptWatch = Stopwatch.StartNew();
        var decryptedBytes = Encrypt.Decrypt(encryptedPath, password);
        decryptWatch.Stop();

        var originalBytes = File.ReadAllBytes(plainPath);
        var matches = originalBytes.SequenceEqual(decryptedBytes);

        Console.WriteLine($"{sizeMb,4} MB | encrypt: {encryptWatch.ElapsedMilliseconds,6} ms | decrypt: {decryptWatch.ElapsedMilliseconds,6} ms | match: {matches}");
    }
    finally
    {
        DeleteIfExists(plainPath);
        DeleteIfExists(encryptedPath);
    }
}

static int[] ParseSizes(string[] args)
{
    if (args.Length == 0)
        return [1, 10, 50, 100];

    return args.Select(int.Parse).Where(v => v > 0).Distinct().OrderBy(v => v).ToArray();
}

static string CreateDeterministicWorkbook(int sizeMb)
{
    var path = Path.Combine(Path.GetTempPath(), $"excelencryptor-bench-vector-{sizeMb}mb-{Guid.NewGuid():N}.xlsx");
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
            "  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\n" +
            "</workbook>\n");

        WriteZipEntry(archive, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
            "</Relationships>\n");

        WriteZipEntry(archive, "xl/worksheets/sheet1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Benchmark</t></is></c></row></sheetData></worksheet>\n");

        var payloadBytes = Math.Max(0, targetBytes - fileStream.Position - 4096);
        WritePayload(archive, "xl/media/payload.bin", payloadBytes);
    }

    var finalLength = new FileInfo(path).Length;
    if (finalLength < targetBytes)
    {
        var deficit = targetBytes - finalLength + 4096;
        using var fileStream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Update, leaveOpen: false);
        WritePayload(archive, "xl/media/payload2.bin", deficit);
    }

    return path;
}

static void WriteZipEntry(ZipArchive archive, string name, string content)
{
    var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
    entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    using var stream = entry.Open();
    using var writer = new StreamWriter(stream, Encoding.UTF8, leaveOpen: false);
    writer.Write(content);
}

static void WritePayload(ZipArchive archive, string name, long count)
{
    var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
    entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    var block = new byte[8192];
    for (var i = 0; i < block.Length; i++)
        block[i] = (byte)(i % 251);

    using var stream = entry.Open();
    long written = 0;
    while (written < count)
    {
        var toWrite = (int)Math.Min(block.Length, count - written);
        stream.Write(block, 0, toWrite);
        written += toWrite;
    }
}

static void DeleteIfExists(string path)
{
    if (File.Exists(path))
        File.Delete(path);
}

