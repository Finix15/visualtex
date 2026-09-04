using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;

namespace VisualTeX.MathTypeConversion;

public sealed class MathTypeEquationRecord
{
    public int Index { get; set; }
    public string RelationshipId { get; set; } = string.Empty;
    public string ProgId { get; set; } = string.Empty;
    public string PartName { get; set; } = string.Empty;
    public string? MathMl { get; set; }
    public bool Display { get; set; }
    public string? Error { get; set; }
    public MathTypeParseStatus Status { get; set; }
    public string ReasonCode { get; set; } = string.Empty;
    public int ErrorOffset { get; set; } = -1;
    public string OleFingerprint { get; set; } = string.Empty;
    public bool CanConvert => Status == MathTypeParseStatus.Convertible
        && !string.IsNullOrWhiteSpace(MathMl) && string.IsNullOrWhiteSpace(Error);
}

public sealed class MathTypePackageOleRecord
{
    public int Index { get; set; }
    public string RelationshipId { get; set; } = string.Empty;
    public string ProgId { get; set; } = string.Empty;
    public string PartName { get; set; } = string.Empty;
    public string OleFingerprint { get; set; } = string.Empty;
    public byte[] OleBytes { get; set; } = Array.Empty<byte>();
}

public sealed class MathTypePackageSnapshot
{
    public string SourcePath { get; set; } = string.Empty;
    public string DocumentFingerprint { get; set; } = string.Empty;
    public IReadOnlyList<MathTypePackageOleRecord> Items { get; set; }
        = Array.Empty<MathTypePackageOleRecord>();
}

public static class MathTypeDocumentScanner
{
    private const string MathTypeProgId = "Equation.DSMT4";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace RelationshipAttribute = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static MathTypePackageSnapshot ReadPackageSnapshot(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("A DOCX path is required.", nameof(path));
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".docm", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Only DOCX and DOCM packages are supported.");

        string documentFingerprint;
        using (var hashStream = File.Open(
                   path,
                   FileMode.Open,
                   FileAccess.Read,
                   FileShare.ReadWrite | FileShare.Delete))
        using (var sha = SHA256.Create())
            documentFingerprint = ToHex(sha.ComputeHash(hashStream));

        using var stream = File.Open(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var document = LoadXml(archive, "word/document.xml");
        var relationships = LoadXml(archive, "word/_rels/document.xml.rels")
            .Root?.Elements(Relationships + "Relationship")
            .Where(item => ((string?)item.Attribute("Type"))
                ?.EndsWith("/oleObject", StringComparison.Ordinal) == true)
            .ToDictionary(
                item => (string)item.Attribute("Id")!,
                item => (string)item.Attribute("Target")!,
                StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);
        var items = new List<MathTypePackageOleRecord>();
        foreach (var ole in document.Descendants(Office + "OLEObject"))
        {
            var progId = (string?)ole.Attribute("ProgID") ?? string.Empty;
            if (!string.Equals(progId, MathTypeProgId, StringComparison.OrdinalIgnoreCase))
                continue;
            var objectType = (string?)ole.Attribute("Type") ?? "Embed";
            if (!string.Equals(objectType, "Embed", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Linked MathType OLE objects are not supported.");
            var relationshipId = (string?)ole.Attribute(RelationshipAttribute + "id")
                ?? string.Empty;
            if (!relationships.TryGetValue(relationshipId, out var target))
                throw new InvalidDataException("The MathType OLE relationship is missing.");
            var partName = NormalizeWordTarget(target);
            var entry = archive.GetEntry(partName)
                ?? throw new InvalidDataException("The MathType OLE part is missing.");
            if (entry.Length <= 0 || entry.Length > MathTypeLimits.MaximumPayloadBytes)
                throw new InvalidDataException("The MathType OLE part size is invalid.");
            byte[] bytes;
            using (var input = entry.Open())
            using (var buffer = new MemoryStream(checked((int)entry.Length)))
            {
                input.CopyTo(buffer);
                bytes = buffer.ToArray();
            }
            items.Add(new MathTypePackageOleRecord
            {
                Index = items.Count + 1,
                RelationshipId = relationshipId,
                ProgId = progId,
                PartName = partName,
                OleBytes = bytes,
                OleFingerprint = ComputeSha256(bytes),
            });
        }
        return new MathTypePackageSnapshot
        {
            SourcePath = Path.GetFullPath(path),
            DocumentFingerprint = documentFingerprint,
            Items = items,
        };
    }

    public static string ComputeFileFingerprint(string path)
    {
        using var stream = File.Open(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var sha = SHA256.Create();
        return ToHex(sha.ComputeHash(stream));
    }

    public static int ValidateAndCountMathTypeOleObjects(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A DOCX path is required.", nameof(path));
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var document = LoadXml(archive, "word/document.xml");
        var relationships = LoadXml(archive, "word/_rels/document.xml.rels")
            .Root?.Elements(Relationships + "Relationship")
            .Where(item => ((string?)item.Attribute("Type"))?.EndsWith("/oleObject", StringComparison.Ordinal) == true)
            .ToDictionary(item => (string)item.Attribute("Id")!, item => (string)item.Attribute("Target")!, StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);
        var count = 0;
        foreach (var ole in document.Descendants(Office + "OLEObject"))
        {
            var progId = (string?)ole.Attribute("ProgID") ?? string.Empty;
            if (!string.Equals(progId, MathTypeProgId, StringComparison.OrdinalIgnoreCase)) continue;
            count++;
            var relationId = (string?)ole.Attribute(RelationshipAttribute + "id") ?? string.Empty;
            if (!relationships.TryGetValue(relationId, out var target))
                throw new InvalidDataException("The MathType OLE relationship is missing.");
            var entry = archive.GetEntry(NormalizeWordTarget(target))
                ?? throw new InvalidDataException("The MathType OLE part is missing.");
            if (entry.Length <= 0 || entry.Length > MathTypeLimits.MaximumPayloadBytes)
                throw new InvalidDataException("The MathType OLE part size is invalid.");
        }
        return count;
    }

    public static IReadOnlyList<MathTypeEquationRecord> ScanDocx(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A DOCX path is required.", nameof(path));
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var document = LoadXml(archive, "word/document.xml");
        var relationships = LoadXml(archive, "word/_rels/document.xml.rels")
            .Root?.Elements(Relationships + "Relationship")
            .Where(item => ((string?)item.Attribute("Type"))?.EndsWith("/oleObject", StringComparison.Ordinal) == true)
            .ToDictionary(item => (string)item.Attribute("Id")!, item => (string)item.Attribute("Target")!, StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);
        var result = new List<MathTypeEquationRecord>();
        foreach (var ole in document.Descendants(Office + "OLEObject"))
        {
            var progId = (string?)ole.Attribute("ProgID") ?? string.Empty;
            if (!string.Equals(progId, MathTypeProgId, StringComparison.OrdinalIgnoreCase)) continue;
            var relationId = (string?)ole.Attribute(RelationshipAttribute + "id") ?? string.Empty;
            var record = new MathTypeEquationRecord
            {
                Index = result.Count + 1,
                RelationshipId = relationId,
                ProgId = progId,
            };
            result.Add(record);
            try
            {
                if (!relationships.TryGetValue(relationId, out var target))
                    throw new InvalidDataException("The MathType OLE relationship is missing.");
                var normalized = NormalizeWordTarget(target);
                record.PartName = normalized;
                var entry = archive.GetEntry(normalized)
                    ?? throw new InvalidDataException("The MathType OLE part is missing.");
                if (entry.Length <= 0 || entry.Length > MathTypeLimits.MaximumPayloadBytes)
                    throw new InvalidDataException("The MathType OLE part size is invalid.");
                using var input = entry.Open();
                using var buffer = new MemoryStream(checked((int)entry.Length));
                input.CopyTo(buffer);
                var decoded = MathTypeEquationDecoder.DecodeOle(buffer.ToArray());
                record.Status = decoded.Status;
                record.ReasonCode = decoded.ReasonCode;
                record.ErrorOffset = decoded.ErrorOffset;
                record.OleFingerprint = decoded.OleFingerprint;
                record.MathMl = decoded.MathMl;
                record.Display = decoded.Display;
                record.Error = decoded.CanConvert ? null : decoded.Reason;
            }
            catch (Exception error) when (error is InvalidDataException or NotSupportedException)
            {
                record.Status = error is NotSupportedException ? MathTypeParseStatus.Unsupported : MathTypeParseStatus.Corrupt;
                record.ReasonCode = error is NotSupportedException ? "DOCUMENT_UNSUPPORTED" : "DOCUMENT_CORRUPT";
                record.Error = error.Message;
            }
        }
        return result;
    }

    public static MathTypeDecodeResult DecodeFlatOpc(string flatOpc)
    {
        return MathTypeEquationDecoder.DecodeOle(ExtractOleFromFlatOpc(flatOpc));
    }

    public static byte[] ExtractOleFromFlatOpc(string flatOpc)
    {
        if (string.IsNullOrWhiteSpace(flatOpc)) throw new InvalidDataException("The Word Flat OPC package is empty.");
        var package = XDocument.Parse(flatOpc, LoadOptions.None);
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        var binary = package.Descendants(pkg + "part")
            .Where(part => ((string?)part.Attribute(pkg + "contentType"))?.Contains("oleObject") == true)
            .Select(part => (string?)part.Element(pkg + "binaryData"))
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (string.IsNullOrWhiteSpace(binary))
            throw new InvalidDataException("Word did not include the selected OLE binary in Flat OPC.");
        byte[] bytes;
        try { bytes = Convert.FromBase64String(binary!); }
        catch (FormatException error) { throw new InvalidDataException("The Flat OPC OLE payload is invalid.", error); }
        if (bytes.Length == 0 || bytes.Length > MathTypeLimits.MaximumPayloadBytes)
            throw new InvalidDataException("The Flat OPC OLE payload size is invalid.");
        return bytes;
    }

    private static XDocument LoadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name) ?? throw new InvalidDataException($"The DOCX part '{name}' is missing.");
        if (entry.Length > MathTypeLimits.MaximumPayloadBytes) throw new InvalidDataException($"The DOCX part '{name}' is too large.");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    private static string NormalizeWordTarget(string target)
    {
        var value = target.Replace('\\', '/');
        while (value.StartsWith("../", StringComparison.Ordinal)) value = value.Substring(3);
        value = value.TrimStart('/');
        if (!value.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) value = "word/" + value;
        if (value.Contains("../")) throw new InvalidDataException("The MathType relationship target is unsafe.");
        return value;
    }

    private static string ComputeSha256(byte[] bytes)
    {
        using var sha = SHA256.Create();
        return ToHex(sha.ComputeHash(bytes));
    }

    private static string ToHex(byte[] bytes) =>
        string.Concat(bytes.Select(value => value.ToString("x2")));
}
