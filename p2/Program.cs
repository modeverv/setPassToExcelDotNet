using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenMcdf;

namespace p2;

public static class Program
{
    private static void Main(string[] args)
    {
        var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var inputPath = Path.Combine(projectDir, "a.xlsx");
        var outputPath = Path.Combine(projectDir, "b.xlsx");
        var testFilePath = Path.Combine(projectDir, "poi_b.xlsx");

        var password = "pass";

        try
        {
            ExcelEncryptor.Encrypt(inputPath, outputPath, password);
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
            var decrypted = ExcelEncryptor.Decrypt(encryptedPath, password);

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
            var dotnetDecrypted = ExcelEncryptor.Decrypt(dotnetPath, password);
            var poiDecrypted = ExcelEncryptor.Decrypt(poiPath, password);

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

public static class ExcelEncryptor
{
    private const int KEY_SIZE = 128; // AES-128
    private const int BLOCK_SIZE = 16; // 16 bytes
    private const int SALT_SIZE = 16; // 16 bytes
    private const int SPIN_COUNT = 100000; // Agile 既定
    private const int SEGMENT_LENGTH = 4096; // パッケージ暗号化のセグメント長
    private const int HASH_SIZE = 20; // SHA1 = 20 bytes

    public static void Encrypt(string inputPath, string outputPath, string password)
    {
        var packageData = File.ReadAllBytes(inputPath);
        var (xmlDoc, encryptionKey, keySalt, integritySalt) = GenerateEncryptionInfo(password);
        var encryptedPackage = EncryptPackage(packageData, encryptionKey, keySalt);
        UpdateIntegrityHMAC(encryptedPackage, packageData.Length, encryptionKey, keySalt, integritySalt, xmlDoc);
        CreateEncryptedFile(outputPath, xmlDoc, encryptedPackage);
    }

    // === EncryptionInfo 生成（HMAC鍵の下準備まで） ===
    private static (XDocument, byte[], byte[], byte[]) GenerateEncryptionInfo(string password)
    {
        var keySalt = RandomBytes(SALT_SIZE); // keyData.saltValue
        var verifierSalt = RandomBytes(SALT_SIZE); // p:encryptedKey.saltValue
        var pwHash = HashPassword(password, verifierSalt, SPIN_COUNT);

        // 検証データ（verifier / verifierHash）
        var verifier = RandomBytes(SALT_SIZE);
        var keySpec = RandomBytes(KEY_SIZE / 8); // 実際の AES 鍵 (encryptedKey の中身)
        var encryptionKey = keySpec;

        // POI 互換ブロックキー
        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };

        var encryptedVerifier = HashInput(pwHash, verifierSalt, kVerifierInputBlock, verifier, KEY_SIZE / 8);
        var verifierHash = SHA1.Create().ComputeHash(verifier);
        var encryptedVerifierHash = HashInput(pwHash, verifierSalt, kHashedVerifierBlock, verifierHash, KEY_SIZE / 8);
        var encryptedKey = HashInput(pwHash, verifierSalt, kCryptoKeyBlock, keySpec, KEY_SIZE / 8);

        // dataIntegrity: encryptedHmacKey だけ先に作る
        var integritySalt = RandomBytes(HASH_SIZE); // HMAC の生鍵（20B）
        byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
        var ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, BLOCK_SIZE);
        var hmacKeyPadded = PadBlock(integritySalt); // 16 境界に 0 詰め
        var encryptedHmacKey = EncryptWithAES(hmacKeyPadded, encryptionKey, ivKey);

        // EncryptionInfo XML の組み立て（encryptedHmacValue は後で埋める）
        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        var keyDataElement = new XElement(ns + "keyData",
            new XAttribute("blockSize", BLOCK_SIZE),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HASH_SIZE),
            new XAttribute("keyBits", KEY_SIZE),
            new XAttribute("saltSize", SALT_SIZE),
            new XAttribute("saltValue", Convert.ToBase64String(keySalt))
        );

        var dataIntegrityElement = new XElement(ns + "dataIntegrity",
            new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
            new XAttribute("encryptedHmacValue", "") // 後で UpdateIntegrityHMAC で上書き
        );

        var encryptedKeyElement = new XElement(p + "encryptedKey",
            new XAttribute("blockSize", BLOCK_SIZE),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("encryptedKeyValue", Convert.ToBase64String(encryptedKey)),
            new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(encryptedVerifier)),
            new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(encryptedVerifierHash)),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HASH_SIZE),
            new XAttribute("keyBits", KEY_SIZE),
            new XAttribute("saltSize", SALT_SIZE),
            new XAttribute("saltValue", Convert.ToBase64String(verifierSalt)),
            new XAttribute("spinCount", SPIN_COUNT)
        );

        var xmlDoc = new XDocument(
            new XElement(ns + "encryption",
                new XAttribute(XNamespace.Xmlns + "p", p.NamespaceName),
                keyDataElement,
                dataIntegrityElement,
                new XElement(ns + "keyEncryptors",
                    new XElement(ns + "keyEncryptor",
                        new XAttribute("uri", p.NamespaceName),
                        encryptedKeyElement
                    )
                )
            )
        );

        return (xmlDoc, encryptionKey, keySalt, integritySalt);
    }

    // === Integrity HMAC を計算して XML を更新 ===
    private static void UpdateIntegrityHMAC(byte[] encryptedPackage, int oleStreamSize, byte[] encryptionKey,
        byte[] keySalt, byte[] integritySalt, XDocument xmlDoc)
    {
        using (var hmac = new HMACSHA1(integritySalt))
        {
            // 先頭の StreamSize(8B, little-endian) を HMAC に供給
            var sizeBytes = BitConverter.GetBytes((long)oleStreamSize);
            hmac.TransformBlock(sizeBytes, 0, 8, null, 0);

            // EncryptedPackage 本体（サイズ 8B を除く）
            var body = new byte[encryptedPackage.Length - 8];
            Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
            hmac.TransformFinalBlock(body, 0, body.Length);

            // HMAC を 16 バイト境界に 0 パディング → AES-CBC で暗号化
            var hmacValPadded = PadBlock(hmac.Hash);
            byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
            var ivVal = GenerateIv(keySalt, kIntegrityValueBlock, BLOCK_SIZE);
            var encryptedHmacValue = EncryptWithAES(hmacValPadded, encryptionKey, ivVal);

            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
            xmlDoc.Root.Element(ns + "dataIntegrity")
                .SetAttributeValue("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue));
        }
    }

    // === ここまでが前半。後半で Decrypt/EncryptPackage/CFB 書き出し/補助関数 を続けます ===
    // === Decrypt 実装 ===
    public static byte[] Decrypt(string encryptedPath, string password)
    {
        using (var root = RootStorage.OpenRead(encryptedPath))
        {
            using (var encInfoStream = root.OpenStream("EncryptionInfo"))
            using (var reader = new BinaryReader(encInfoStream))
            {

                var xmlBytes = reader.ReadBytes((int)(encInfoStream.Length - 8));
                var xmlString = Encoding.UTF8.GetString(xmlBytes);

                // XMLからsalt等を抽出（簡易パーサー）
                var keySaltMatch = Regex.Match(xmlString, @"<keyData[^>]*saltValue=""([^""]+)""");
                var verifierSaltMatch = Regex.Match(xmlString, @"<p:encryptedKey[^>]*saltValue=""([^""]+)""");
                var spinCountMatch = Regex.Match(xmlString, @"spinCount=""(\d+)""");
                var encryptedKeyMatch = Regex.Match(xmlString, @"encryptedKeyValue=""([^""]+)""");


                if (!keySaltMatch.Success || !verifierSaltMatch.Success || !spinCountMatch.Success ||
                    !encryptedKeyMatch.Success)
                    throw new InvalidProgramException("暗号化情報の解析に失敗");

                var keySalt = Convert.FromBase64String(keySaltMatch.Groups[1].Value);
                var verifierSalt = Convert.FromBase64String(verifierSaltMatch.Groups[1].Value);
                var spinCount = int.Parse(spinCountMatch.Groups[1].Value);
                var encryptedKey = Convert.FromBase64String(encryptedKeyMatch.Groups[1].Value);

                var pwHash = HashPassword(password, verifierSalt, spinCount);
                byte[] kCryptoKeyBlock = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };
                var intermedKey = GenerateKey(pwHash, kCryptoKeyBlock, KEY_SIZE / 8);
                var iv = GenerateIv(verifierSalt, null, BLOCK_SIZE);

                byte[] decryptionKey;
                using (var aes = Aes.Create())
                {
                    aes.Key = intermedKey;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.None;
                    using (var dec = aes.CreateDecryptor())
                    {
                        decryptionKey = dec.TransformFinalBlock(encryptedKey, 0, encryptedKey.Length);
                    }
                }

                var actualKey = new byte[KEY_SIZE / 8];
                Array.Copy(decryptionKey, actualKey, actualKey.Length);

                using (var encPackageStream = root.OpenStream("EncryptedPackage"))
                using (var br = new BinaryReader(encPackageStream))
                {
                    var streamSize = br.ReadInt64();
                    using (var ms = new MemoryStream())
                    {
                        var block = 0;
                        var remaining = streamSize;
                        while (remaining > 0)
                        {
                            var segSize = (int)Math.Min(SEGMENT_LENGTH, remaining);
                            var isLast = remaining <= SEGMENT_LENGTH;

                            var encryptedSeg =
                                isLast ? br.ReadBytes(PadLen((int)remaining)) : br.ReadBytes(SEGMENT_LENGTH);
                            var blockKey = BitConverter.GetBytes(block);
                            var segIv = GenerateIv(keySalt, blockKey, BLOCK_SIZE);

                            using (var aes = Aes.Create())
                            {
                                aes.Key = actualKey;
                                aes.IV = segIv;
                                aes.Mode = CipherMode.CBC;
                                aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                                using (var dec = aes.CreateDecryptor())
                                {
                                    var decSeg = dec.TransformFinalBlock(encryptedSeg, 0, encryptedSeg.Length);
                                    ms.Write(decSeg, 0, Math.Min(segSize, decSeg.Length));
                                }
                            }

                            remaining -= segSize;
                            block++;
                        }

                        return ms.ToArray();
                    }
                }
            }
        }
    }

    // === パッケージ暗号化 ===
    private static byte[] EncryptPackage(byte[] packageData, byte[] encryptionKey, byte[] keySalt)
    {
        using (var ms = new MemoryStream())
        using (var bw = new BinaryWriter(ms))
        {
            bw.Write((long)packageData.Length);
            var offset = 0;
            var block = 0;
            while (offset < packageData.Length)
            {
                var segSize = Math.Min(SEGMENT_LENGTH, packageData.Length - offset);
                var isLast = offset + segSize >= packageData.Length;
                var seg = new byte[segSize];
                Array.Copy(packageData, offset, seg, 0, segSize);
                var blockKey = BitConverter.GetBytes(block);
                var iv = GenerateIv(keySalt, blockKey, BLOCK_SIZE);
                using (var aes = Aes.Create())
                {
                    aes.Key = encryptionKey;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                    if (!isLast && segSize < SEGMENT_LENGTH)
                    {
                        var padSeg = new byte[SEGMENT_LENGTH];
                        Array.Copy(seg, padSeg, segSize);
                        seg = padSeg;
                    }

                    using (var enc = aes.CreateEncryptor())
                    {
                        bw.Write(enc.TransformFinalBlock(seg, 0, seg.Length));
                    }
                }

                offset += segSize;
                block++;
            }

            return ms.ToArray();
        }
    }

    // === CFBファイル書き込み (DataSpaces含む) ===
    private static void CreateEncryptedFile(string outputPath, XDocument encryptionInfo, byte[] encryptedData)
    {
        using (var root = RootStorage.Create(outputPath))
        {
            using (var s = root.CreateStream("EncryptedPackage"))
            {
                s.Write(encryptedData, 0, encryptedData.Length);
            }

            CreateDataSpacesStructure(root);
            using (var s2 = root.CreateStream("EncryptionInfo"))
            using (var bw = new BinaryWriter(s2))
            {
                bw.Write((ushort)4);
                bw.Write((ushort)4);
                bw.Write((uint)0x40);
                var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                          encryptionInfo.Root.ToString(SaveOptions.DisableFormatting);
                xml = xml.Replace(" />", "/>");
                bw.Write(Encoding.UTF8.GetBytes(xml));
            }
        }
    }

    private static void CreateDataSpacesStructure(RootStorage root)
    {
        var ds = root.CreateStorage("\u0006DataSpaces");
        using (var v = ds.CreateStream("Version"))
        using (var bw = new BinaryWriter(v))
        {
            WriteUnicodeLPP4(bw, "Microsoft.Container.DataSpaces");
            bw.Write((ushort)1);
            bw.Write((ushort)0);
            bw.Write((ushort)1);
            bw.Write((ushort)0);
            bw.Write((ushort)1);
            bw.Write((ushort)0);
        }

        using (var m = ds.CreateStream("DataSpaceMap"))
        using (var bw = new BinaryWriter(m))
        {
            bw.Write((uint)8);
            bw.Write((uint)1);
            var pos = m.Position;
            bw.Write((uint)0);
            bw.Write((uint)1);
            bw.Write((uint)0);
            WriteUnicodeLPP4(bw, "EncryptedPackage");
            WriteUnicodeLPP4(bw, "StrongEncryptionDataSpace");
            var end = m.Position;
            m.Seek(pos, SeekOrigin.Begin);
            bw.Write((uint)(end - pos));
            m.Seek(end, SeekOrigin.Begin);
        }

        var dsi = ds.CreateStorage("DataSpaceInfo");
        using (var s = dsi.CreateStream("StrongEncryptionDataSpace"))
        using (var bw = new BinaryWriter(s))
        {
            bw.Write((uint)8);
            bw.Write((uint)1);
            WriteUnicodeLPP4(bw, "StrongEncryptionTransform");
        }

        var ti = ds.CreateStorage("TransformInfo");
        var st = ti.CreateStorage("StrongEncryptionTransform");
        using (var p = st.CreateStream("\u0006Primary"))
        using (var bw = new BinaryWriter(p))
        {
            var hdr = p.Position;
            bw.Write((uint)0);
            bw.Write((uint)1);
            WriteUnicodeLPP4(bw, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
            var hdrEnd = p.Position;
            p.Seek(hdr, SeekOrigin.Begin);
            bw.Write((uint)(hdrEnd - hdr));
            p.Seek(hdrEnd, SeekOrigin.Begin);
            WriteUnicodeLPP4(bw, "Microsoft.Container.EncryptionTransform");
            bw.Write((ushort)1);
            bw.Write((ushort)0);
            bw.Write((ushort)1);
            bw.Write((ushort)0);
            bw.Write((ushort)1);
            bw.Write((ushort)0);
            bw.Write((uint)0);
            bw.Write((uint)0);
            bw.Write((uint)0);
            bw.Write((uint)4);
        }
    }

    private static void WriteUnicodeLPP4(BinaryWriter bw, string s)
    {
        var b = Encoding.Unicode.GetBytes(s);
        bw.Write((uint)b.Length);
        bw.Write(b);
        var pad = (4 - b.Length % 4) % 4;
        for (var i = 0; i < pad; i++) bw.Write((byte)0);
    }

    // === 汎用ヘルパー ===
    private static byte[] RandomBytes(int n)
    {
        var b = new byte[n];
        RandomNumberGenerator.Fill(b);
        return b;
    }

    private static byte[] PadBlock(byte[] d)
    {
        var s = PadLen(d.Length);
        var r = new byte[s];
        Buffer.BlockCopy(d, 0, r, 0, d.Length);
        return r;
    }

    private static int PadLen(int len)
    {
        return (len + 15) / 16 * 16;
    }

    private static byte[] HashPassword(string pw, byte[] salt, int spin)
    {
        var pwb = Encoding.Unicode.GetBytes(pw);
        using (var sha = SHA1.Create())
        {
            sha.TransformBlock(salt, 0, salt.Length, null, 0);
            sha.TransformFinalBlock(pwb, 0, pwb.Length);
            var h = sha.Hash;
            for (var i = 0; i < spin; i++)
            {
                sha.Initialize();
                var iter = BitConverter.GetBytes(i);
                sha.TransformBlock(iter, 0, 4, null, 0);
                sha.TransformFinalBlock(h, 0, h.Length);
                h = sha.Hash;
            }

            return h;
        }
    }

    private static byte[] HashInput(byte[] pwHash, byte[] salt, byte[] blk, byte[] input, int keySize)
    {
        var k = GenerateKey(pwHash, blk, keySize);
        var iv = GenerateIv(salt, null, BLOCK_SIZE);
        var pad = PadBlock(input);
        return EncryptWithAES(pad, k, iv);
    }

    private static byte[] GenerateKey(byte[] h, byte[] blk, int ks)
    {
        using (var sha = SHA1.Create())
        {
            sha.TransformBlock(h, 0, h.Length, null, 0);
            sha.TransformFinalBlock(blk, 0, blk.Length);
            var d = sha.Hash;
            var k = new byte[ks];
            Array.Copy(d, k, Math.Min(d.Length, ks));
            return k;
        }
    }

    private static byte[] GenerateIv(byte[] salt, byte[] blk, int bs)
    {
        if (blk == null)
        {
            var iv = new byte[bs];
            Array.Copy(salt, iv, Math.Min(salt.Length, bs));
            return iv;
        }

        using (var sha = SHA1.Create())
        {
            sha.TransformBlock(salt, 0, salt.Length, null, 0);
            sha.TransformFinalBlock(blk, 0, blk.Length);
            var d = sha.Hash;
            var iv = new byte[bs];
            Array.Copy(d, iv, Math.Min(d.Length, bs));
            return iv;
        }
    }

    private static byte[] EncryptWithAES(byte[] d, byte[] k, byte[] iv)
    {
        using (var aes = Aes.Create())
        {
            aes.Key = k;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            return aes.CreateEncryptor().TransformFinalBlock(d, 0, d.Length);
        }
    }
}
