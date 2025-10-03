using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenMcdf;

namespace ExcelEncryptor;

public static class Encrypt
{
    private const int KeySize = 128; // AES-128
    private const int BlockSize = 16; // 16 bytes
    private const int SaltSize = 16; // 16 bytes
    private const int SpinCount = 100000; // Agile 既定
    private const int SegmentLength = 4096; // パッケージ暗号化のセグメント長
    private const int HashSize = 20; // SHA1 = 20 bytes

    public static void FromBytesToFile(byte[] wbByte, string outputPath, string password)
    {
        ValidateEncryptionParameters();

        if (wbByte == null || wbByte.Length == 0)
            throw new ArgumentException("Input data cannot be null or empty", nameof(wbByte));

        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        if (password.Length > 255)
            throw new ArgumentException("Password is too long (max 255 characters)", nameof(password));

        try
        {
            var (xmlDoc, encryptionKey, keySalt, integritySalt) = GenerateEncryptionInfo(password);
            var encryptedPackage = EncryptPackage(wbByte, encryptionKey, keySalt);
            UpdateIntegrityHmac(encryptedPackage, wbByte.Length, encryptionKey, keySalt, integritySalt, xmlDoc);
            CreateEncryptedFile(outputPath, xmlDoc, encryptedPackage);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to encrypt file", ex);
        }
    }

    public static void FromFileToFile(string inputPath, string outputPath, string password)
    {
        var packageData = File.ReadAllBytes(inputPath);
        FromBytesToFile(packageData, outputPath, password);
    }

    private static (XDocument, byte[], byte[], byte[]) GenerateEncryptionInfo(string password)
    {
        var keySalt = RandomBytes(SaltSize); // keyData.saltValue
        var verifierSalt = RandomBytes(SaltSize); // p:encryptedKey.saltValue
        var pwHash = HashPassword(password, verifierSalt, SpinCount);

        // 検証データ（verifier / verifierHash）
        var verifier = RandomBytes(SaltSize);
        var keySpec = RandomBytes(KeySize / 8); // 実際の AES 鍵 (encryptedKey の中身)
        var encryptionKey = keySpec;

        // POI 互換ブロックキー
        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };

        var encryptedVerifier = HashInput(pwHash, verifierSalt, kVerifierInputBlock, verifier, KeySize / 8);

        byte[] verifierHash;
        using (var sha = SHA1.Create())
        {
            verifierHash = sha.ComputeHash(verifier);
        }

        var encryptedVerifierHash = HashInput(pwHash, verifierSalt, kHashedVerifierBlock, verifierHash, KeySize / 8);


        var encryptedKey = HashInput(pwHash, verifierSalt, kCryptoKeyBlock, keySpec, KeySize / 8);

        // dataIntegrity: encryptedHmacKey だけ先に作る
        var integritySalt = RandomBytes(HashSize); // HMAC の生鍵（20B）
        byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
        var ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, BlockSize);
        var hmacKeyPadded = PadBlock(integritySalt); // 16 境界に 0 詰め
        var encryptedHmacKey = EncryptWithAes(hmacKeyPadded, encryptionKey, ivKey);

        // EncryptionInfo XML の組み立て（encryptedHmacValue は後で埋める）
        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        var keyDataElement = new XElement(ns + "keyData",
            new XAttribute("blockSize", BlockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HashSize),
            new XAttribute("keyBits", KeySize),
            new XAttribute("saltSize", SaltSize),
            new XAttribute("saltValue", Convert.ToBase64String(keySalt))
        );

        var dataIntegrityElement = new XElement(ns + "dataIntegrity",
            new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
            new XAttribute("encryptedHmacValue", "") // 後で UpdateIntegrityHMAC で上書き
        );

        var encryptedKeyElement = new XElement(p + "encryptedKey",
            new XAttribute("blockSize", BlockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("encryptedKeyValue", Convert.ToBase64String(encryptedKey)),
            new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(encryptedVerifier)),
            new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(encryptedVerifierHash)),
            new XAttribute("hashAlgorithm", "SHA1"),
            new XAttribute("hashSize", HashSize),
            new XAttribute("keyBits", KeySize),
            new XAttribute("saltSize", SaltSize),
            new XAttribute("saltValue", Convert.ToBase64String(verifierSalt)),
            new XAttribute("spinCount", SpinCount)
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

    private static void UpdateIntegrityHmac(byte[] encryptedPackage, int oleStreamSize, byte[] encryptionKey,
        byte[] keySalt, byte[] integritySalt, XDocument xmlDoc)
    {
        using var hmac = new HMACSHA1(integritySalt);
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
        var ivVal = GenerateIv(keySalt, kIntegrityValueBlock, BlockSize);
        var encryptedHmacValue = EncryptWithAes(hmacValPadded, encryptionKey, ivVal);

        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        if (xmlDoc.Root != null)
            xmlDoc.Root.Element(ns + "dataIntegrity")
                ?.SetAttributeValue("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue));
    }

    // === Decrypt 実装 ===
    public static byte[] Decrypt(string encryptedPath, string password)
    {
        if (!File.Exists(encryptedPath))
            throw new FileNotFoundException("Encrypted file not found", encryptedPath);

        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        using var root = RootStorage.OpenRead(encryptedPath);

        // EncryptionInfo の読み取り
        CfbStream encInfoStream;
        try
        {
            encInfoStream = root.OpenStream("EncryptionInfo");
        }
        catch
        {
            throw new InvalidOperationException("File is not encrypted (EncryptionInfo missing)");
        }

        using (encInfoStream)
        using (var reader = new BinaryReader(encInfoStream))
        {
            // バージョン情報とフラグの読み取り
            var versionMajor = reader.ReadUInt16();
            var versionMinor = reader.ReadUInt16();
            reader.ReadUInt32();

            if (versionMajor != 4 || versionMinor != 4)
                throw new NotSupportedException($"Unsupported encryption version: {versionMajor}.{versionMinor}");

            // XML部分の読み取りとパース
            var xmlBytes = reader.ReadBytes((int)(encInfoStream.Length - 8));
            var xmlString = Encoding.UTF8.GetString(xmlBytes);

            // XMLから必要な情報を抽出
            var keySaltMatch = Regex.Match(xmlString, @"<keyData[^>]*saltValue=""([^""]+)""");
            var verifierSaltMatch = Regex.Match(xmlString, @"<p:encryptedKey[^>]*saltValue=""([^""]+)""");
            var spinCountMatch = Regex.Match(xmlString, @"spinCount=""(\d+)""");
            var encryptedKeyMatch = Regex.Match(xmlString, @"encryptedKeyValue=""([^""]+)""");

            if (!keySaltMatch.Success || !verifierSaltMatch.Success || !spinCountMatch.Success ||
                !encryptedKeyMatch.Success)
                throw new InvalidOperationException("fail: check encrypted info");

            var keySalt = Convert.FromBase64String(keySaltMatch.Groups[1].Value);
            var verifierSalt = Convert.FromBase64String(verifierSaltMatch.Groups[1].Value);
            var spinCount = int.Parse(spinCountMatch.Groups[1].Value);
            var encryptedKey = Convert.FromBase64String(encryptedKeyMatch.Groups[1].Value);

            var xmlDoc = XDocument.Parse(xmlString);

            // パスワード検証
            if (!VerifyPassword(password, xmlDoc, xmlString))
                throw new UnauthorizedAccessException("Invalid password");

            // 暗号化キーの復号化
            var pwHash = HashPassword(password, verifierSalt, spinCount);
            byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
            var keyIntermedKey = GenerateKey(pwHash, kCryptoKeyBlock, KeySize / 8);
            var keyIv = GenerateIv(verifierSalt, null, BlockSize);

            byte[] actualKey;
            using (var aes = Aes.Create())
            {
                aes.Key = keyIntermedKey;
                aes.IV = keyIv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                using var dec = aes.CreateDecryptor();
                var decryptionKey = dec.TransformFinalBlock(encryptedKey, 0, encryptedKey.Length);
                actualKey = new byte[KeySize / 8];
                Array.Copy(decryptionKey, actualKey, actualKey.Length);
            }

            // EncryptedPackage の読み取り（一度だけ）
            CfbStream encPackageStream;
            try
            {
                encPackageStream = root.OpenStream("EncryptedPackage");
            }
            catch
            {
                throw new InvalidOperationException("EncryptedPackage stream not found");
            }

            byte[] encryptedPackageData;
            using (encPackageStream)
            {
                encryptedPackageData = new byte[encPackageStream.Length];
                _ = encPackageStream.Read(encryptedPackageData, 0, encryptedPackageData.Length);
            }

            // EncryptedPackage の復号化
            byte[] decryptedData;
            long streamSize;

            using (var ms = new MemoryStream(encryptedPackageData))
            using (var br = new BinaryReader(ms))
            {
                streamSize = br.ReadInt64();

                using var outMs = new MemoryStream();
                var block = 0;
                var remaining = streamSize;

                while (remaining > 0)
                {
                    var segSize = (int)Math.Min(SegmentLength, remaining);
                    var isLast = remaining <= SegmentLength;

                    // 暗号化されたセグメントの読み取り
                    var encryptedSeg = isLast
                        ? br.ReadBytes(PadLen((int)remaining))
                        : br.ReadBytes(SegmentLength);

                    // ブロック番号からIVを生成
                    var blockKey = BitConverter.GetBytes(block);
                    var segIv = GenerateIv(keySalt, blockKey, BlockSize);

                    using (var aes = Aes.Create())
                    {
                        aes.Key = actualKey;
                        aes.IV = segIv;
                        aes.Mode = CipherMode.CBC;
                        aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                        using var dec = aes.CreateDecryptor();
                        var decSeg = dec.TransformFinalBlock(encryptedSeg, 0, encryptedSeg.Length);
                        outMs.Write(decSeg, 0, Math.Min(segSize, decSeg.Length));
                    }

                    remaining -= segSize;
                    block++;
                }

                decryptedData = outMs.ToArray();
            }

            // HMAC整合性検証
            if (!VerifyIntegrity(encryptedPackageData, (int)streamSize, actualKey, keySalt, xmlString))
                throw new InvalidOperationException("Data integrity check failed - file may be corrupted or tampered");

            return decryptedData;
        }
    }


    // === CFBファイル書き込み (DataSpaces含む) ===
    private static void CreateEncryptedFile(string outputPath, XDocument encryptionInfo, byte[] encryptedData)
    {
        using var root = RootStorage.Create(outputPath);
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
            if (encryptionInfo.Root != null)
            {
                var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                          encryptionInfo.Root.ToString(SaveOptions.DisableFormatting);
                xml = xml.Replace(" />", "/>");
                bw.Write(Encoding.UTF8.GetBytes(xml));
            }
        }
    }

    private static byte[] EncryptPackage(byte[] packageData, byte[] encryptionKey, byte[] keySalt)
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);
        bw.Write((long)packageData.Length);
        var offset = 0;
        var block = 0;
        while (offset < packageData.Length)
        {
            var segSize = Math.Min(SegmentLength, packageData.Length - offset);
            var isLast = offset + segSize >= packageData.Length;
            var seg = new byte[segSize];
            Array.Copy(packageData, offset, seg, 0, segSize);
            var blockKey = BitConverter.GetBytes(block);
            var iv = GenerateIv(keySalt, blockKey, BlockSize);
            using (var aes = Aes.Create())
            {
                aes.Key = encryptionKey;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
                if (!isLast && segSize < SegmentLength)
                {
                    var padSeg = new byte[SegmentLength];
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


    private static void CreateDataSpacesStructure(RootStorage root)
    {
        var ds = root.CreateStorage("\u0006DataSpaces");
        using (var v = ds.CreateStream("Version"))
        using (var bw = new BinaryWriter(v))
        {
            WriteUnicodeLpp4(bw, "Microsoft.Container.DataSpaces");
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
            WriteUnicodeLpp4(bw, "EncryptedPackage");
            WriteUnicodeLpp4(bw, "StrongEncryptionDataSpace");
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
            WriteUnicodeLpp4(bw, "StrongEncryptionTransform");
        }

        var ti = ds.CreateStorage("TransformInfo");
        var st = ti.CreateStorage("StrongEncryptionTransform");
        using (var p = st.CreateStream("\u0006Primary"))
        using (var bw = new BinaryWriter(p))
        {
            var hdr = p.Position;
            bw.Write((uint)0);
            bw.Write((uint)1);
            WriteUnicodeLpp4(bw, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
            var hdrEnd = p.Position;
            p.Seek(hdr, SeekOrigin.Begin);
            bw.Write((uint)(hdrEnd - hdr));
            p.Seek(hdrEnd, SeekOrigin.Begin);
            WriteUnicodeLpp4(bw, "Microsoft.Container.EncryptionTransform");
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

    private static void WriteUnicodeLpp4(BinaryWriter bw, string s)
    {
        var b = Encoding.Unicode.GetBytes(s);
        bw.Write((uint)b.Length);
        bw.Write(b);
        var pad = (4 - b.Length % 4) % 4;
        for (var i = 0; i < pad; i++) bw.Write((byte)0);
    }

    private static byte[] RandomBytes(int n)
    {
        var b = new byte[n];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(b);

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

    private static void ValidateEncryptionParameters()
    {
        if (KeySize != 128 && KeySize != 192 && KeySize != 256)
#pragma warning disable CS0162 // 到達できないコードが検出されました
            throw new InvalidOperationException($"Invalid key size: {KeySize}");
#pragma warning restore CS0162 // 到達できないコードが検出されました

        if (BlockSize != 16)
#pragma warning disable CS0162 // 到達できないコードが検出されました
            throw new InvalidOperationException($"Invalid block size: {BlockSize}");
#pragma warning restore CS0162 // 到達できないコードが検出されました

        if (SpinCount < 1)
#pragma warning disable CS0162 // 到達できないコードが検出されました
            throw new InvalidOperationException($"Invalid spin count: {SpinCount}");
#pragma warning restore CS0162 // 到達できないコードが検出されました
    }

    private static byte[] HashPassword(string pw, byte[] salt, int spin)
    {
        var pwb = Encoding.Unicode.GetBytes(pw);
        using var sha = SHA1.Create();

        try
        {
            sha.TransformBlock(salt, 0, salt.Length, null, 0);
            sha.TransformFinalBlock(pwb, 0, pwb.Length);
            var h = (byte[])sha.Hash.Clone(); // Clone to avoid issues

            for (var i = 0; i < spin; i++)
            {
                sha.Initialize();
                var iter = BitConverter.GetBytes(i);
                sha.TransformBlock(iter, 0, 4, null, 0);
                sha.TransformFinalBlock(h, 0, h.Length);
                h = (byte[])sha.Hash.Clone();
            }

            return h;
        }
        finally
        {
            Array.Clear(pwb, 0, pwb.Length);
        }
    }

    private static byte[] HashInput(byte[] pwHash, byte[] salt, byte[] blk, byte[] input, int keySize)
    {
        var k = GenerateKey(pwHash, blk, keySize);
        var iv = GenerateIv(salt, null, BlockSize);
        var pad = PadBlock(input);
        return EncryptWithAes(pad, k, iv);
    }

    private static byte[] GenerateKey(byte[] h, byte[] blk, int ks)
    {
        using var sha = SHA1.Create();
        sha.TransformBlock(h, 0, h.Length, null, 0);
        sha.TransformFinalBlock(blk, 0, blk.Length);
        var d = sha.Hash;
        var k = new byte[ks];
        Array.Copy(d, k, Math.Min(d.Length, ks));
        return k;
    }

    private static byte[] GenerateIv(byte[] salt, byte[]? blk, int bs)
    {
        if (blk == null)
        {
            var iv1 = new byte[bs];
            Array.Copy(salt, iv1, Math.Min(salt.Length, bs));
            return iv1;
        }

        using var sha = SHA1.Create();

        sha.TransformBlock(salt, 0, salt.Length, null, 0);
        sha.TransformFinalBlock(blk, 0, blk.Length);
        var d = sha.Hash;
        var iv = new byte[bs];
        Array.Copy(d, iv, Math.Min(d.Length, bs));
        return iv;
    }

    private static byte[] EncryptWithAes(byte[] d, byte[] k, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = k;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        return aes.CreateEncryptor().TransformFinalBlock(d, 0, d.Length);
    }

    private static bool VerifyPassword(string password, XDocument xmlDoc, string xmlString)
    {
        var encVerifierMatch = Regex.Match(xmlString, @"encryptedVerifierHashInput=""([^""]+)""");
        var encVerifierHashMatch = Regex.Match(xmlString, @"encryptedVerifierHashValue=""([^""]+)""");
        var verifierSaltMatch = Regex.Match(xmlString, @"<p:encryptedKey[^>]*saltValue=""([^""]+)""");
        var spinCountMatch = Regex.Match(xmlString, @"spinCount=""(\d+)""");

        if (!encVerifierMatch.Success || !encVerifierHashMatch.Success) return false;

        var encryptedVerifier = Convert.FromBase64String(encVerifierMatch.Groups[1].Value);
        var encryptedVerifierHash = Convert.FromBase64String(encVerifierHashMatch.Groups[1].Value);
        var verifierSalt = Convert.FromBase64String(verifierSaltMatch.Groups[1].Value);
        var spinCount = int.Parse(spinCountMatch.Groups[1].Value);

        var pwHash = HashPassword(password, verifierSalt, spinCount);

        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        var intermedKey = GenerateKey(pwHash, kVerifierInputBlock, KeySize / 8);
        var iv = GenerateIv(verifierSalt, null, BlockSize);

        byte[] decryptedVerifier;
        using (var aes = Aes.Create())
        {
            aes.Key = intermedKey;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using var dec = aes.CreateDecryptor();
            decryptedVerifier = dec.TransformFinalBlock(encryptedVerifier, 0, encryptedVerifier.Length);
        }

        // 修正：usingを追加
        byte[] verifierHash;
        using (var sha = SHA1.Create())
        {
            verifierHash = sha.ComputeHash(decryptedVerifier, 0, SaltSize);
        }

        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        intermedKey = GenerateKey(pwHash, kHashedVerifierBlock, KeySize / 8);

        byte[] decryptedVerifierHash;
        using (var aes = Aes.Create())
        {
            aes.Key = intermedKey;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            using var dec = aes.CreateDecryptor();
            decryptedVerifierHash = dec.TransformFinalBlock(encryptedVerifierHash, 0, encryptedVerifierHash.Length);
        }

        return verifierHash.Take(HashSize).SequenceEqual(decryptedVerifierHash.Take(HashSize));
    }

    private static bool VerifyIntegrity(byte[] encryptedPackage, int oleStreamSize,
        byte[] encryptionKey, byte[] keySalt, string xmlString)
    {
        var encHmacKeyMatch = Regex.Match(xmlString, @"encryptedHmacKey=""([^""]+)""");
        var encHmacValueMatch = Regex.Match(xmlString, @"encryptedHmacValue=""([^""]+)""");

        if (!encHmacKeyMatch.Success || !encHmacValueMatch.Success) return false;

        var encryptedHmacKey = Convert.FromBase64String(encHmacKeyMatch.Groups[1].Value);
        var encryptedHmacValue = Convert.FromBase64String(encHmacValueMatch.Groups[1].Value);

        // Decrypt HMAC key
        byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
        var ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, BlockSize);
        byte[] hmacKey;
        using (var aes = Aes.Create())
        {
            aes.Key = encryptionKey;
            aes.IV = ivKey;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            hmacKey = aes.CreateDecryptor().TransformFinalBlock(encryptedHmacKey, 0, encryptedHmacKey.Length);
        }

        hmacKey = hmacKey.Take(HashSize).ToArray();

        // Calculate HMAC
        using var hmac = new HMACSHA1(hmacKey);
        var sizeBytes = BitConverter.GetBytes((long)oleStreamSize);
        hmac.TransformBlock(sizeBytes, 0, 8, null, 0);
        var body = new byte[encryptedPackage.Length - 8];
        Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
        hmac.TransformFinalBlock(body, 0, body.Length);

        // Decrypt expected HMAC value
        byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
        var ivVal = GenerateIv(keySalt, kIntegrityValueBlock, BlockSize);
        byte[] expectedHmac;
        using (var aes = Aes.Create())
        {
            aes.Key = encryptionKey;
            aes.IV = ivVal;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            expectedHmac = aes.CreateDecryptor().TransformFinalBlock(encryptedHmacValue, 0, encryptedHmacValue.Length);
        }

        expectedHmac = expectedHmac.Take(HashSize).ToArray();

        return hmac.Hash.SequenceEqual(expectedHmac);
    }
}
