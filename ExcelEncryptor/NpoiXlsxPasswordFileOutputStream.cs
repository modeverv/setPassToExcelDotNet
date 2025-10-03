using System.IO;

namespace ExcelEncryptor;

/// <summary>
///     stream for NPOI IWorkbook with password
/// </summary>
public class NpoiXlsxPasswordFileOutputStream : Stream
{
    private readonly MemoryStream _buffer = new();
    private readonly string _outputPath;
    private readonly string _password;

    public NpoiXlsxPasswordFileOutputStream(string outputPath, string password)
    {
        _outputPath = outputPath;
        _password = password;
    }

    public override bool CanRead => false;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => _buffer.Length;

    public override long Position
    {
        get => _buffer.Position;
        set => _buffer.Position = value;
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        _buffer.Write(buffer, offset, count);
    }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        return 0;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        return _buffer.Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        _buffer.SetLength(value);
    }

    public override void Close()
    {
        base.Close();
        var raw = _buffer.ToArray();
        Encrypt.FromBytesToFile(raw, _outputPath, _password);
    }
}
