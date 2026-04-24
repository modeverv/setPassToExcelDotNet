using System.IO.Compression;
using System.Text;
using Xunit;

namespace ExcelEncryptor.Tests;

public class ContentRetentionTests
{
    private const string Password = "pass";

    [Fact]
    public void EncryptDecrypt_ContentRichWorkbook_PreservesWorksheetContentAndPackageParts()
    {
        var originalPath = CreateContentRichWorkbook();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-content-{Guid.NewGuid():N}.xlsx");

        try
        {
            var originalBytes = File.ReadAllBytes(originalPath);
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
            encryptor.EncryptFile(originalPath, encryptedPath, Password);

            var decryptedBytes = Encrypt.Decrypt(encryptedPath, Password);
            Assert.Equal(originalBytes, decryptedBytes);

            AssertZipEntryListEquals(originalBytes, decryptedBytes);
            AssertPackagePartsPreserved(originalBytes, decryptedBytes);
        }
        finally
        {
            DeleteIfExists(originalPath);
            DeleteIfExists(encryptedPath);
        }
    }

    [Fact]
    public void Decrypt_ContentRichWorkbook_WithWrongPassword_ThrowsUnauthorizedAccessException()
    {
        var originalPath = CreateContentRichWorkbook();
        var encryptedPath = Path.Combine(Path.GetTempPath(), $"excelencryptor-content-wrong-{Guid.NewGuid():N}.xlsx");

        try
        {
            var encryptor = new Encrypt(AesKeySize.Aes256, HashAlgorithmType.Sha512);
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

    private static string CreateContentRichWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelencryptor-content-vector-{Guid.NewGuid():N}.xlsx");
        using var fileStream = File.Create(path);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, leaveOpen: false);

        WriteZipEntry(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Default Extension=""png"" ContentType=""image/png""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/worksheets/sheet2.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/comments1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml""/>
  <Override PartName=""/xl/drawings/drawing1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.drawing+xml""/>
</Types>
");

        WriteZipEntry(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>
");

        WriteZipEntry(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets>
  <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/>
  <sheet name=""日本語"" sheetId=""2"" state=""hidden"" r:id=""rId2""/>
</sheets></workbook>
");

        WriteZipEntry(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet2.xml""/>
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>
");

        WriteZipEntry(archive, "xl/styles.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><fonts count=""1""><font><sz val=""11""/><color theme=""1""/><name val=""Calibri""/><family val=""2""/><scheme val=""minor""/></font></fonts><fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills><borders count=""2""><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style=""thin""><color rgb=""FF000000""/></bottom><diagonal/></border></borders><cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs><cellXfs count=""2""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyBorder=""1""/></cellXfs></styleSheet>
");

        WriteZipEntry(archive, "xl/worksheets/sheet1.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheetData>
  <row r=""1""><c r=""A1"" t=""inlineStr""><is><t>セル値</t></is></c><c r=""B1"" s=""1""><f>SUM(B2:B3)</f><v>3</v></c><c r=""C1"" s=""1""><f>2+2</f><v>4</v></c></row>
  <row r=""2""><c r=""B2""><v>1</v></c><c r=""B3""><v>2</v></c></row>
  <row r=""4""><c r=""A4"" t=""inlineStr""><is><t>結合</t></is></c></row>
  <row r=""5""><c r=""A5"" t=""inlineStr""><is><t>リンク</t></is></c></row>
  <row r=""6""><c r=""A6"" t=""inlineStr""><is><t>コメント</t></is></c></row>
  <row r=""7""><c r=""A7"" s=""1"" t=""inlineStr""><is><t>罫線</t></is></c></row>
  <row r=""8""><c r=""D1"" t=""inlineStr""><is><t>画像</t></is></c></row>
</sheetData><mergeCells count=""1""><mergeCell ref=""A4:B4""/></mergeCells><hyperlinks><hyperlink ref=""A5"" r:id=""rIdHyperlink1""/></hyperlinks><drawing r:id=""rIdDrawing1""/>
</worksheet>
");

        WriteZipEntry(archive, "xl/worksheets/_rels/sheet1.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rIdComments1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"" Target=""../comments1.xml""/><Relationship Id=""rIdHyperlink1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"" Target=""https://example.com/"" TargetMode=""External""/><Relationship Id=""rIdDrawing1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing1.xml""/></Relationships>
");

        WriteZipEntry(archive, "xl/comments1.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<comments xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><authors><author>Tester</author></authors><commentList><comment ref=""A6"" authorId=""0""><text><r><t>コメント保持</t></r></text></comment></commentList></comments>
");

        WriteZipEntry(archive, "xl/drawings/drawing1.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=""1"" name=""Picture 1""/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=""rIdImage1"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""9525"" cy=""9525""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>
");

        WriteZipEntry(archive, "xl/drawings/_rels/drawing1.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rIdImage1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image1.png""/></Relationships>
");

        WriteZipEntry(archive, "xl/worksheets/sheet2.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>非表示シート</t></is></c></row></sheetData></worksheet>
");

        WriteZipEntry(archive, "xl/media/image1.png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+7O0AAAAASUVORK5CYII="));
        return path;
    }

    private static void AssertPackagePartsPreserved(byte[] originalBytes, byte[] decryptedBytes)
    {
        Assert.Equal(ReadZipEntry(originalBytes, "[Content_Types].xml"), ReadZipEntry(decryptedBytes, "[Content_Types].xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/workbook.xml"), ReadZipEntry(decryptedBytes, "xl/workbook.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/_rels/workbook.xml.rels"), ReadZipEntry(decryptedBytes, "xl/_rels/workbook.xml.rels"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/styles.xml"), ReadZipEntry(decryptedBytes, "xl/styles.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/worksheets/sheet1.xml"), ReadZipEntry(decryptedBytes, "xl/worksheets/sheet1.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/worksheets/_rels/sheet1.xml.rels"), ReadZipEntry(decryptedBytes, "xl/worksheets/_rels/sheet1.xml.rels"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/comments1.xml"), ReadZipEntry(decryptedBytes, "xl/comments1.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/drawings/drawing1.xml"), ReadZipEntry(decryptedBytes, "xl/drawings/drawing1.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/drawings/_rels/drawing1.xml.rels"), ReadZipEntry(decryptedBytes, "xl/drawings/_rels/drawing1.xml.rels"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/worksheets/sheet2.xml"), ReadZipEntry(decryptedBytes, "xl/worksheets/sheet2.xml"));
        Assert.Equal(ReadZipEntry(originalBytes, "xl/media/image1.png"), ReadZipEntry(decryptedBytes, "xl/media/image1.png"));

        var workbookXml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/workbook.xml"));
        var sheet1Xml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/worksheets/sheet1.xml"));
        var sheet1RelsXml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/worksheets/_rels/sheet1.xml.rels"));
        var stylesXml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/styles.xml"));
        var commentsXml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/comments1.xml"));
        var drawingRelXml = Encoding.UTF8.GetString(ReadZipEntry(originalBytes, "xl/drawings/_rels/drawing1.xml.rels"));

        Assert.Contains("name=\"日本語\"", workbookXml);
        Assert.Contains("state=\"hidden\"", workbookXml);
        Assert.Contains("<c r=\"A1\" t=\"inlineStr\">", sheet1Xml);
        Assert.Contains("<f>SUM(B2:B3)</f>", sheet1Xml);
        Assert.Contains("<mergeCell ref=\"A4:B4\"/>", sheet1Xml);
        Assert.Contains("rIdHyperlink1", sheet1Xml);
        Assert.Contains("rIdComments1", sheet1RelsXml);
        Assert.Contains("border", stylesXml);
        Assert.Contains("コメント保持", commentsXml);
        Assert.Contains("image1.png", drawingRelXml);
    }

    private static void AssertZipEntryListEquals(byte[] originalBytes, byte[] decryptedBytes)
    {
        using var originalStream = new MemoryStream(originalBytes, writable: false);
        using var decryptedStream = new MemoryStream(decryptedBytes, writable: false);
        using var originalArchive = new ZipArchive(originalStream, ZipArchiveMode.Read, leaveOpen: false);
        using var decryptedArchive = new ZipArchive(decryptedStream, ZipArchiveMode.Read, leaveOpen: false);

        var originalEntries = originalArchive.Entries.Select(entry => entry.FullName).ToArray();
        var decryptedEntries = decryptedArchive.Entries.Select(entry => entry.FullName).ToArray();

        Assert.Equal(originalEntries, decryptedEntries);
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

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }
}


