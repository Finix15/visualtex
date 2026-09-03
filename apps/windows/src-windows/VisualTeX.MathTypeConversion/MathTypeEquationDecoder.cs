using System.Security.Cryptography;

namespace VisualTeX.MathTypeConversion;

public enum MathTypeParseStatus
{
    Corrupt,
    Convertible,
    Unsupported,
}

public sealed class MathTypeDecodeResult
{
    public MathTypeParseStatus Status { get; set; }
    public string ReasonCode { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
    public int ErrorOffset { get; set; } = -1;
    public string OleFingerprint { get; set; } = string.Empty;
    public string MathMl { get; set; } = string.Empty;
    public bool Display { get; set; }
    public int MtefVersion { get; set; }
    public string Product { get; set; } = string.Empty;
    public bool CanConvert => Status == MathTypeParseStatus.Convertible && MathMl.Length != 0;
}

internal sealed class MtefParseException : Exception
{
    internal MtefParseException(string code, string message, int offset, bool unsupported = false)
        : base(message) { Code = code; Offset = offset; Unsupported = unsupported; }
    internal string Code { get; }
    internal int Offset { get; }
    internal bool Unsupported { get; }
}

public static class MathTypeEquationDecoder
{
    private const int EquationNativeHeaderSize = 28;

    public static MathTypeDecodeResult DecodeOle(byte[] oleBytes)
    {
        if (oleBytes is null) throw new ArgumentNullException(nameof(oleBytes));
        var fingerprint = Fingerprint(oleBytes);
        try
        {
            if (oleBytes.Length == 0 || oleBytes.Length > MathTypeLimits.MaximumPayloadBytes)
                throw new MtefParseException("OLE_SIZE_INVALID", "The MathType OLE payload size is invalid.", 0);
            var compound = new CompoundFileReader(oleBytes);
            var native = compound.ReadStream("Equation Native");
            if (native.Length <= EquationNativeHeaderSize || BitConverter.ToUInt32(native, 0) != EquationNativeHeaderSize)
                throw new MtefParseException("EQUATION_NATIVE_HEADER_INVALID", "The Equation Native header is invalid.", 0);
            var mtef = new byte[native.Length - EquationNativeHeaderSize];
            Buffer.BlockCopy(native, EquationNativeHeaderSize, mtef, 0, mtef.Length);
            var result = Mtef5MathMlConverter.Convert(mtef);
            result.OleFingerprint = fingerprint;
            return result;
        }
        catch (MtefParseException error)
        {
            return Failed(fingerprint, error.Unsupported ? MathTypeParseStatus.Unsupported : MathTypeParseStatus.Corrupt,
                error.Code, error.Message, error.Offset);
        }
        catch (NotSupportedException error)
        {
            return Failed(fingerprint, MathTypeParseStatus.Unsupported, "OLE_UNSUPPORTED", error.Message, -1);
        }
        catch (Exception error) when (error is InvalidDataException or OverflowException or ArgumentException)
        {
            return Failed(fingerprint, MathTypeParseStatus.Corrupt, "OLE_CORRUPT", error.Message, -1);
        }
    }

    private static MathTypeDecodeResult Failed(string fingerprint, MathTypeParseStatus status,
        string code, string reason, int offset) => new()
    {
        Status = status, ReasonCode = code, Reason = reason, ErrorOffset = offset, OleFingerprint = fingerprint,
    };

    private static string Fingerprint(byte[] value)
    {
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(value).Select(item => item.ToString("x2")));
    }
}
