using System;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OpenMcdf;

namespace ExcelEncryptor;

public enum AesKeySize
{
    Aes128 = 128,
    Aes192 = 192,
    Aes256 = 256
}

public enum HashAlgorithmType
{
    Sha1 = 20, // 20 bytes
    Sha256 = 32, // 32 bytes
    Sha384 = 48, // 48 bytes
    Sha512 = 64, // 64 bytes
    Md5 = 16 // 16 bytes (非推奨だが互換性のため)
}

public partial class Encrypt
{
    private readonly int _blockSize = 16;
    private readonly HashAlgorithmType _hashAlgorithm;
    private readonly int _hashSize;
    private readonly int _keySize;
    private readonly int _saltSize = 16;
    private readonly int _segmentLength = 4096;
    private readonly int _spinCount = 100000;
    
    /// <summary>
    /// Provides functionalities for encrypting and decrypting files and byte arrays
    /// using AES encryption, with configurable key sizes and hash algorithms.
    /// </summary>
    public Encrypt(
        AesKeySize keySize = AesKeySize.Aes128,
        HashAlgorithmType hashAlgorithm = HashAlgorithmType.Sha1)
    {
        _keySize = (int)keySize;
        _hashAlgorithm = hashAlgorithm;
        _hashSize = (int)hashAlgorithm;
        ValidateParameters();
    }
    
    /// <summary>
    /// Converts a byte array to an encrypted file using the specified password and saves it to the provided output path.
    /// </summary>
    /// <param name="bytes">The byte array representing the data to be encrypted and saved to a file.</param>
    /// <param name="outputPath">The path where the encrypted file will be saved.</param>
    /// <param name="passwordString">The password used to encrypt the file.</param>
    public static void FromBytesToFile(byte[] bytes, string outputPath, string passwordString)
    {
        var encryptor = new Encrypt();
        encryptor.EncryptToFile(bytes, outputPath, passwordString);
    }
    
    /// <summary>
    /// Encrypts the content of a file specified by the input path and saves the encrypted data to another file at the specified output path using the given password.
    /// </summary>
    /// <param name="inputPath">The path of the file to be encrypted.</param>
    /// <param name="outputPath">The path where the encrypted file will be saved.</param>
    /// <param name="passwordString">The password used to encrypt the file.</param>
    public static void FromFileToFile(string inputPath, string outputPath, string passwordString)
    {
        var encryptor = new Encrypt();
        encryptor.EncryptFile(inputPath, outputPath, passwordString);
    }
    
    /// <summary>
    /// Encrypts the file at the specified input path with the given password and saves the encrypted output to the specified output path.
    /// </summary>
    /// <param name="inputPath">The path to the input file that will be encrypted.</param>
    /// <param name="outputPath">The path where the encrypted file will be saved.</param>
    /// <param name="password">The password used to encrypt the file.</param>
    public void EncryptFile(string inputPath, string outputPath, string password)
    {
        var packageData = File.ReadAllBytes(inputPath);
        EncryptToFile(packageData, outputPath, password);
    }
    
    
    /// <summary>
    /// Validates the configured encryption parameters, including the key size and hash size,
    /// ensuring they conform to supported values. Throws an exception if any parameter is invalid.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the key size or hash size does not match one of the accepted values.
    /// </exception>
    private void ValidateParameters()
    {
        if (_keySize != 128 && _keySize != 192 && _keySize != 256)
            throw new InvalidOperationException($"Invalid key size: {_keySize}");
        
        if (_hashSize != 16 && _hashSize != 20 && _hashSize != 32 && _hashSize != 48 && _hashSize != 64)
            throw new InvalidOperationException($"Invalid hash size: {_hashSize}");
    }
    
    /// <summary>
    /// Creates and returns an instance of a hash algorithm based on the specified type.
    /// Supports various algorithms including MD5, SHA-1, SHA-256, SHA-384, and SHA-512.
    /// </summary>
    /// <returns>An instance of the selected hash algorithm.</returns>
    /// <exception cref="NotSupportedException">Thrown when an unsupported hash algorithm type is specified.</exception>
    private HashAlgorithm CreateHashAlgorithm()
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => MD5.Create(),
            HashAlgorithmType.Sha1 => SHA1.Create(),
            HashAlgorithmType.Sha256 => SHA256.Create(),
            HashAlgorithmType.Sha384 => SHA384.Create(),
            HashAlgorithmType.Sha512 => SHA512.Create(),
            _ => throw new NotSupportedException($"Hash algorithm not supported: {_hashAlgorithm}")
        };
    }
    
    /// <summary>
    /// Retrieves the name of the hash algorithm as a string based on the configured
    /// hash algorithm type.
    /// </summary>
    /// <returns>The name of the hash algorithm.</returns>
    /// <exception cref="NotSupportedException">Thrown when the configured hash algorithm type is unsupported.</exception>
    private string GetHashAlgorithmName()
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => "MD5",
            HashAlgorithmType.Sha1 => "SHA1",
            HashAlgorithmType.Sha256 => "SHA256",
            HashAlgorithmType.Sha384 => "SHA384",
            HashAlgorithmType.Sha512 => "SHA512",
            _ => throw new NotSupportedException($"Hash algorithm not supported: {_hashAlgorithm}")
        };
    }
    
    /// <summary>
    /// Creates an HMAC (Hash-Based Message Authentication Code) instance using the specified cryptographic hash
    /// algorithm and key, enabling the computation of a message authentication code for data integrity and authenticity validation.
    /// </summary>
    /// <param name="key">The secret key used for HMAC generation. The key must be compatible with the selected hash algorithm.</param>
    /// <returns>An HMAC instance that uses the specified hash algorithm and key.</returns>
    /// <exception cref="NotSupportedException">Thrown when the specified hash algorithm is not supported.</exception>
    private HMAC CreateHmac(byte[] key)
    {
        return _hashAlgorithm switch
        {
            HashAlgorithmType.Md5 => new HMACMD5(key),
            HashAlgorithmType.Sha1 => new HMACSHA1(key),
            HashAlgorithmType.Sha256 => new HMACSHA256(key),
            HashAlgorithmType.Sha384 => new HMACSHA384(key),
            HashAlgorithmType.Sha512 => new HMACSHA512(key),
            _ => throw new NotSupportedException($"HMAC algorithm not supported: {_hashAlgorithm}")
        };
    }
    
    /// <summary>
    /// Encrypts a byte array and saves the resulting encrypted data to a specified file path.
    /// </summary>
    /// <param name="wbByte">The byte array to be encrypted. Cannot be null or empty.</param>
    /// <param name="outputPath">The path to the output file where the encrypted data will be saved. The file will be created or overwritten.</param>
    /// <param name="password">The password used for encryption. Cannot be null, empty, or exceed 255 characters.</param>
    /// <exception cref="ArgumentException">Thrown if the input data is null, empty, or if the password criteria are not met.</exception>
    /// <exception cref="InvalidOperationException">Thrown if the encryption process fails due to an internal error.</exception>
    private void EncryptToFile(byte[] wbByte, string outputPath, string password)
    {
        if (wbByte == null || wbByte.Length == 0)
            throw new ArgumentException("Input data cannot be null or empty", nameof(wbByte));
        
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));
        
        if (password.Length > 255)
            throw new ArgumentException("Password is too long (max 255 characters)", nameof(password));

        ValidateWorkbookPackage(wbByte);
        
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
    
    /// <summary>
    /// Generates encryption information, including an XML document, an encryption key, key salt,
    /// and integrity salt, based on the provided password. This information is used to securely encrypt
    /// and validate data.
    /// </summary>
    /// <param name="password">The password used to derive cryptographic keys and hashes.</param>
    /// <returns>
    /// A tuple containing:
    /// 1. An <c>XDocument</c> representing the encryption structure for storing key-related metadata.
    /// 2. A byte array representing the derived encryption key.
    /// 3. A byte array representing the salt used for key derivation.
    /// 4. A byte array representing the salt used for data integrity validation.
    /// </returns>
    private (XDocument, byte[], byte[], byte[]) GenerateEncryptionInfo(string password)
    {
        var keySalt = RandomBytes(_saltSize);
        var verifierSalt = RandomBytes(_saltSize);
        var pwHash = HashPassword(password, verifierSalt, _spinCount);
        
        var verifier = RandomBytes(_saltSize);
        var keySpec = RandomBytes(_keySize / 8);
        var encryptionKey = keySpec;
        
        byte[] kVerifierInputBlock = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
        byte[] kHashedVerifierBlock = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
        byte[] kCryptoKeyBlock = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
        
        var encryptedVerifier = HashInput(pwHash, verifierSalt, kVerifierInputBlock, verifier, _keySize / 8);
        
        byte[] verifierHash;
        using (var hashAlg = CreateHashAlgorithm())
        {
            verifierHash = hashAlg.ComputeHash(verifier);
        }
        
        var encryptedVerifierHash = HashInput(pwHash, verifierSalt, kHashedVerifierBlock, verifierHash, _keySize / 8);
        var encryptedKey = HashInput(pwHash, verifierSalt, kCryptoKeyBlock, keySpec, _keySize / 8);
        
        var integritySalt = RandomBytes(_hashSize);
        byte[] kIntegrityKeyBlock = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
        var ivKey = GenerateIv(keySalt, kIntegrityKeyBlock, _blockSize);
        var hmacKeyPadded = PadBlock(integritySalt);
        var encryptedHmacKey = EncryptWithAes(hmacKeyPadded, encryptionKey, ivKey, false);
        
        XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
        
        var keyDataElement = new XElement(ns + "keyData",
            new XAttribute("blockSize", _blockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("hashAlgorithm", GetHashAlgorithmName()),
            new XAttribute("hashSize", _hashSize),
            new XAttribute("keyBits", _keySize),
            new XAttribute("saltSize", _saltSize),
            new XAttribute("saltValue", Convert.ToBase64String(keySalt))
        );
        
        var dataIntegrityElement = new XElement(ns + "dataIntegrity",
            new XAttribute("encryptedHmacKey", Convert.ToBase64String(encryptedHmacKey)),
            new XAttribute("encryptedHmacValue", "")
        );
        
        var encryptedKeyElement = new XElement(p + "encryptedKey",
            new XAttribute("blockSize", _blockSize),
            new XAttribute("cipherAlgorithm", "AES"),
            new XAttribute("cipherChaining", "ChainingModeCBC"),
            new XAttribute("encryptedKeyValue", Convert.ToBase64String(encryptedKey)),
            new XAttribute("encryptedVerifierHashInput", Convert.ToBase64String(encryptedVerifier)),
            new XAttribute("encryptedVerifierHashValue", Convert.ToBase64String(encryptedVerifierHash)),
            new XAttribute("hashAlgorithm", GetHashAlgorithmName()),
            new XAttribute("hashSize", _hashSize),
            new XAttribute("keyBits", _keySize),
            new XAttribute("saltSize", _saltSize),
            new XAttribute("saltValue", Convert.ToBase64String(verifierSalt)),
            new XAttribute("spinCount", _spinCount)
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
    
    /// <summary>
    /// Encrypts the provided byte array data using AES encryption with the specified key and key salt.
    /// The encryption is performed in multiple segments, ensuring data is securely processed in blocks.
    /// </summary>
    /// <param name="data">The byte array containing the data to encrypt.</param>
    /// <param name="key">The encryption key used to encrypt the data.</param>
    /// <param name="keySalt">The salt used to derive the initialization vector for block encryption.</param>
    /// <returns>A byte array containing the encrypted data.</returns>
    private byte[] EncryptPackage(byte[] data, byte[] key, byte[] keySalt)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);
        
        writer.Write((long)data.Length);
        
        var offset = 0;
        uint blockIndex = 0;
        
        while (offset < data.Length)
        {
            var blockSize = Math.Min(_segmentLength, data.Length - offset);
            var isLast = offset + blockSize >= data.Length;
            
            var iv = GenerateBlockIv(keySalt, blockIndex, _blockSize);
            
            byte[] block;
            if (isLast)
            {
                block = new byte[blockSize];
                Buffer.BlockCopy(data, offset, block, 0, blockSize);
            }
            else
            {
                block = new byte[_segmentLength];
                Buffer.BlockCopy(data, offset, block, 0, blockSize);
            }
            
            var encrypted = EncryptWithAes(block, key, iv, isLast);
            writer.Write(encrypted);
            
            offset += blockSize;
            blockIndex++;
        }
        
        return ms.ToArray();
    }
    
    /// <summary>
    /// Updates the integrity HMAC in the encrypted package and modifies the XML document
    /// to include the encrypted HMAC value for data integrity checking.
    /// </summary>
    /// <param name="encryptedPackage">The byte array representing the encrypted package.</param>
    /// <param name="oleStreamSize">The size of the OLE stream in the original data.</param>
    /// <param name="encryptionKey">The encryption key used for encrypting the package.</param>
    /// <param name="keySalt">The salt used for key derivation.</param>
    /// <param name="integritySalt">The salt used for HMAC generation.</param>
    /// <param name="xmlDoc">The XML document to which the encrypted HMAC value will be added.</param>
    private void UpdateIntegrityHmac(byte[] encryptedPackage, int oleStreamSize, byte[] encryptionKey,
        byte[] keySalt, byte[] integritySalt, XDocument xmlDoc)
    {
        using var hmac = CreateHmac(integritySalt);
        var sizeBytes = BitConverter.GetBytes((long)oleStreamSize);
        hmac.TransformBlock(sizeBytes, 0, 8, null, 0);
        
        var body = new byte[encryptedPackage.Length - 8];
        Buffer.BlockCopy(encryptedPackage, 8, body, 0, body.Length);
        hmac.TransformFinalBlock(body, 0, body.Length);
        
        if (hmac.Hash != null)
        {
            var hmacValPadded = PadBlock(hmac.Hash);
            byte[] kIntegrityValueBlock = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };
            var ivVal = GenerateIv(keySalt, kIntegrityValueBlock, _blockSize);
            var encryptedHmacValue = EncryptWithAes(hmacValPadded, encryptionKey, ivVal, false);
            
            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";
            if (xmlDoc.Root != null)
                xmlDoc.Root.Element(ns + "dataIntegrity")
                    ?.SetAttributeValue("encryptedHmacValue", Convert.ToBase64String(encryptedHmacValue));
        }
        
    }
    
    /// <summary>
    /// Creates an encrypted file with the specified output path, encryption information,
    /// and encrypted package data.
    /// </summary>
    /// <param name="outputPath">The file path where the encrypted file will be created.</param>
    /// <param name="xmlDoc">The XML document containing the encryption metadata.</param>
    /// <param name="encryptedPackage">The byte array containing the encrypted package data.</param>
    private static void CreateEncryptedFile(string outputPath, XDocument xmlDoc, byte[] encryptedPackage)
    {
        using var root = RootStorage.Create(outputPath);
        
        using (var encInfoStream = root.CreateStream("EncryptionInfo"))
        using (var writer = new BinaryWriter(encInfoStream))
        {
            writer.Write((ushort)4);
            writer.Write((ushort)4);
            writer.Write((uint)0x40);
            
            var xmlString = xmlDoc.ToString(SaveOptions.DisableFormatting);
            var xmlBytes = Encoding.UTF8.GetBytes(xmlString);
            writer.Write(xmlBytes);
        }
        
        using (var encPackageStream = root.CreateStream("EncryptedPackage"))
        {
            encPackageStream.Write(encryptedPackage, 0, encryptedPackage.Length);
        }
    }
    
    /// <summary>
    /// Validates if the provided byte array represents a valid Open Office XML (OOXML)
    /// workbook package by checking for the existence of required entries within the archive.
    /// </summary>
    /// <param name="wbByte">The byte array of the workbook to validate.</param>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the byte array does not contain a valid OOXML workbook structure
    /// or when the data is not a valid zip archive.
    /// </exception>
    private static void ValidateWorkbookPackage(byte[] wbByte)
    {
        try
        {
            using var buffer = new MemoryStream(wbByte, writable: false);
            using var archive = new ZipArchive(buffer, ZipArchiveMode.Read, leaveOpen: false);

            if (archive.GetEntry("[Content_Types].xml") == null || archive.GetEntry("xl/workbook.xml") == null)
                throw new InvalidOperationException("Input file is not a valid OOXML workbook");
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidOperationException("Input file is not a valid OOXML workbook", ex);
        }
    }
    
    /// <summary>
    /// Generates a random byte array of the specified length using a cryptographic random number generator.
    /// </summary>
    /// <param name="length">The length of the byte array to generate.</param>
    /// <returns>A byte array containing cryptographically secure random values.</returns>
    private static byte[] RandomBytes(int length)
    {
        var bytes = new byte[length];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(bytes);
        return bytes;
    }
    
    /// <summary>
    /// Pads the input byte array to the nearest multiple of the block size (16 bytes)
    /// by appending zeroes to the end of the array.
    /// </summary>
    /// <param name="data">The input byte array to be padded.</param>
    /// <returns>A new byte array padded to the nearest multiple of 16 bytes, with the original data preserved at the beginning of the array.</returns>
    private static byte[] PadBlock(byte[] data)
    {
        var padded = (data.Length + 15) / 16 * 16;
        var result = new byte[padded];
        Buffer.BlockCopy(data, 0, result, 0, data.Length);
        return result;
    }
    
    /// <summary>
    /// Generates a hashed representation of a password combined with a salt, using
    /// a specified number of iterations for added computational complexity.
    /// </summary>
    /// <param name="pw">The password to hash, provided as a string.</param>
    /// <param name="salt">A byte array representing the salt to be incorporated into the hash.</param>
    /// <param name="spin">The number of iterations to perform for the hash computation.</param>
    /// <returns>A byte array containing the resulting hash.</returns>
    private byte[] HashPassword(string pw, byte[] salt, int spin)
    {
        var pwb = Encoding.Unicode.GetBytes(pw);
        using var hashAlg = CreateHashAlgorithm();
        
        try
        {
            hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
            hashAlg.TransformFinalBlock(pwb, 0, pwb.Length);
            var h = (byte[])hashAlg.Hash?.Clone()!;
            
            for (var i = 0; i < spin; i++)
            {
                hashAlg.Initialize();
                var iter = BitConverter.GetBytes(i);
                hashAlg.TransformBlock(iter, 0, 4, null, 0);
                hashAlg.TransformFinalBlock(h, 0, h.Length);
                h = (byte[])hashAlg.Hash?.Clone()!;
            }
            
            return h;
        }
        finally
        {
            Array.Clear(pwb, 0, pwb.Length);
        }
    }
    
    /// <summary>
    /// Generates a hashed input by combining the provided password hash, salt, block information,
    /// and input data, utilizing AES encryption with a specified key size.
    /// </summary>
    /// <param name="pwHash">The password hash used to derive the encryption key.</param>
    /// <param name="salt">The salt value used to derive the initialization vector (IV).</param>
    /// <param name="blk">The block information used for key derivation.</param>
    /// <param name="input">The input data to be hashed and encrypted.</param>
    /// <param name="keySize">The size of the encryption key in bits.</param>
    /// <returns>A byte array representing the hashed and encrypted result of the input data.</returns>
    private byte[] HashInput(byte[] pwHash, byte[] salt, byte[] blk, byte[] input, int keySize)
    {
        var k = GenerateKey(pwHash, blk, keySize);
        var iv = GenerateIv(salt, null, _blockSize);
        var pad = PadBlock(input);
        return EncryptWithAes(pad, k, iv, false);
    }
    
    /// <summary>
    /// Generates an encryption key by hashing the provided inputs using the configured hash algorithm.
    /// Combines the primary hash (h) and additional block data (blk) to produce a key of the specified size (ks).
    /// </summary>
    /// <param name="h">The primary hash input for generating the key.</param>
    /// <param name="blk">Additional block data to be incorporated into the key derivation process.</param>
    /// <param name="ks">The desired size of the key in bytes.</param>
    /// <returns>A byte array representing the generated encryption key.</returns>
    private byte[] GenerateKey(byte[] h, byte[] blk, int ks)
    {
        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(h, 0, h.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash;
        var k = new byte[ks];
        if (d != null) Array.Copy(d, k, Math.Min(d.Length, ks));
        return k;
    }
    
    /// <summary>
    /// Generates an initialization vector (IV) based on the provided salt, optional block, and block size.
    /// The method uses a cryptographic hash algorithm if a block is provided or directly derives the IV from the salt if the block is null.
    /// </summary>
    /// <param name="salt">The salt used as a base for generating the initialization vector.</param>
    /// <param name="blk">An optional byte array used to further derive the initialization vector. Can be null.</param>
    /// <param name="bs">The size of the initialization vector to output.</param>
    /// <returns>A byte array representing the generated initialization vector.</returns>
    private byte[] GenerateIv(byte[] salt, byte[]? blk, int bs)
    {
        if (blk == null)
        {
            var iv1 = new byte[bs];
            Array.Copy(salt, iv1, Math.Min(salt.Length, bs));
            return iv1;
        }
        
        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blk, 0, blk.Length);
        var d = hashAlg.Hash;
        var iv = new byte[bs];
        if (d != null) Array.Copy(d, iv, Math.Min(d.Length, bs));
        return iv;
    }
    
    /// <summary>
    /// Generates an initialization vector (IV) for AES encryption based on the provided salt,
    /// block key, and block size using the configured hash algorithm.
    /// </summary>
    /// <param name="salt">A byte array used as the salt value to influence the hashing process and ensure uniqueness.</param>
    /// <param name="blockKey">A 32-bit unsigned integer representing the block key, used to differentiate IVs for each block.</param>
    /// <param name="bs">The size of the block in bytes for which the IV is generated.</param>
    /// <returns>A byte array representing the generated initialization vector (IV) of the specified block size.</returns>
    private byte[] GenerateBlockIv(byte[] salt, uint blockKey, int bs)
    {
        var blockBytes = BitConverter.GetBytes(blockKey);
        using var hashAlg = CreateHashAlgorithm();
        hashAlg.TransformBlock(salt, 0, salt.Length, null, 0);
        hashAlg.TransformFinalBlock(blockBytes, 0, 4);
        var hash = hashAlg.Hash;
        var iv = new byte[bs];
        if (hash != null) Array.Copy(hash, iv, Math.Min(hash.Length, bs));
        return iv;
    }
    
    
    /// <summary>
    /// Encrypts data using AES encryption with the specified key, initialization vector, and padding mode.
    /// </summary>
    /// <param name="d">The byte array of data to be encrypted.</param>
    /// <param name="k">The encryption key to use during the encryption process.</param>
    /// <param name="iv">The initialization vector (IV) to use during the encryption process.</param>
    /// <param name="isLast">Specifies whether padding should be applied. If true, PKCS7 padding is used; otherwise, no padding is applied.</param>
    /// <returns>Returns the encrypted byte array.</returns>
    private static byte[] EncryptWithAes(byte[] d, byte[] k, byte[] iv, bool isLast)
    {
        using var aes = Aes.Create();
        aes.Key = k;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = isLast ? PaddingMode.PKCS7 : PaddingMode.None;
        return aes.CreateEncryptor().TransformFinalBlock(d, 0, d.Length);
    }
}