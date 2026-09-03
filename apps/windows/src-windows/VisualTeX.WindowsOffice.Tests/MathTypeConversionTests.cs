using VisualTeX.MathTypeConversion;
using VisualTeX.WordVsto;
using System.IO.Compression;
using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeConversionTests
{
    [Fact]
    public void FullMathType7FixtureContainsExpectedOleInventory()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        var records = MathTypeDocumentScanner.ScanDocx(path);

        Assert.Equal(680, records.Count);
        Assert.All(records, record => Assert.Equal("Equation.DSMT4", record.ProgId));
        Assert.All(records, record => Assert.True(Enum.IsDefined(record.Status)));
        Assert.Contains(records, record => record.CanConvert);
        Assert.All(records.Where(record => !record.CanConvert), record =>
            Assert.False(string.IsNullOrWhiteSpace(record.Error)));
        Assert.All(records.Where(record => record.CanConvert), record =>
        {
            Assert.StartsWith("<math", record.MathMl, StringComparison.Ordinal);
            Assert.Contains("MathML", record.MathMl, StringComparison.Ordinal);
            Assert.DoesNotContain("<mi> </mi>", record.MathMl, StringComparison.Ordinal);
            var xml = XElement.Parse(record.MathMl!);
            XNamespace m = "http://www.w3.org/1998/Math/MathML";
            Assert.DoesNotContain(xml.Descendants(), element =>
                (element.Name == m + "mfrac" || element.Name == m + "msub" || element.Name == m + "msup"
                    || element.Name == m + "msubsup" || element.Name == m + "mroot")
                && element.Elements().Any(child => child.Name == m + "mrow" && !child.Nodes().Any()));
            var omml = WordOmmlConverter.TransformMathMlToOmml(record.MathMl!);
            Assert.False(string.IsNullOrWhiteSpace(WordOmmlConverter.ExtractSingleOMath(omml)));
        });

        using var archive = ZipFile.OpenRead(path);
        using var documentStream = archive.GetEntry("word/document.xml")!.Open();
        var document = XDocument.Load(documentStream);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        Assert.Empty(document.Descendants(math + "oMath"));

        var combined = string.Concat(records.Where(record => record.CanConvert).Select(record => record.MathMl));
        Assert.Contains("<mfrac>", combined, StringComparison.Ordinal);
        Assert.Contains("<msub>", combined, StringComparison.Ordinal);
        Assert.Contains("<msup>", combined, StringComparison.Ordinal);
        Assert.Contains("<mtable>", combined, StringComparison.Ordinal);
    }

    [Fact]
    public void BackupCopyIsByteIdenticalAndKeepsAllMathTypeOleObjects()
    {
        var source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        var backup = Path.Combine(Path.GetTempPath(), $"visualtex-backup-test-{Guid.NewGuid():N}.docx");
        try
        {
            MathTypeBackupVerifier.CopyAndVerify(source, backup, 680);
            Assert.True(File.Exists(backup));
            Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(backup));
            Assert.Equal(680, MathTypeDocumentScanner.ValidateAndCountMathTypeOleObjects(backup));
        }
        finally
        {
            try { File.Delete(backup); } catch { }
        }
    }

    [Fact]
    public void BackupWithUnexpectedOleCountFailsClosedAndRemovesCopy()
    {
        var source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        var backup = Path.Combine(Path.GetTempPath(), $"visualtex-backup-test-{Guid.NewGuid():N}.docx");

        Assert.Throws<InvalidDataException>(() =>
            MathTypeBackupVerifier.CopyAndVerify(source, backup, 679));
        Assert.False(File.Exists(backup));
    }

    [Fact]
    public void BackupSucceedsWhileWordLikeWriterKeepsSourceOpen()
    {
        var source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        var backup = Path.Combine(Path.GetTempPath(), $"visualtex-backup-test-{Guid.NewGuid():N}.docx");
        var expectedBytes = File.ReadAllBytes(source);
        try
        {
            using var wordLikeHandle = File.Open(
                source,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.ReadWrite | FileShare.Delete);

            MathTypeBackupVerifier.CopyAndVerify(source, backup, 680);

            Assert.True(File.Exists(backup));
            Assert.Equal(expectedBytes, File.ReadAllBytes(backup));
        }
        finally
        {
            try { File.Delete(backup); } catch { }
        }
    }

    [Fact]
    public void ExclusivelyLockedSourceFailsClosedAndLeavesNoBackup()
    {
        var fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        var source = Path.Combine(Path.GetTempPath(), $"visualtex-source-test-{Guid.NewGuid():N}.docx");
        var backup = Path.Combine(Path.GetTempPath(), $"visualtex-backup-test-{Guid.NewGuid():N}.docx");
        File.Copy(fixture, source);
        try
        {
            using var exclusiveHandle = File.Open(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            Assert.Throws<IOException>(() =>
                MathTypeBackupVerifier.CopyAndVerify(source, backup, 680));
            Assert.False(File.Exists(backup));
        }
        finally
        {
            try { File.Delete(backup); } catch { }
            try { File.Delete(source); } catch { }
        }
    }

    [Fact]
    public void InvalidOleFailsClosed()
    {
        var result = MathTypeEquationDecoder.DecodeOle(new byte[512]);
        Assert.Equal(MathTypeParseStatus.Corrupt, result.Status);
        Assert.False(result.CanConvert);
        Assert.False(string.IsNullOrWhiteSpace(result.ReasonCode));
        Assert.Equal(64, result.OleFingerprint.Length);
    }

    [Fact]
    public void OversizedOleFailsClosed()
    {
        var result = MathTypeEquationDecoder.DecodeOle(new byte[16 * 1024 * 1024 + 1]);
        Assert.Equal(MathTypeParseStatus.Corrupt, result.Status);
        Assert.Equal("OLE_SIZE_INVALID", result.ReasonCode);
    }

    [Fact]
    public void TruncatedMathTypeOleFailsClosed()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "MathType7", "mathtype7-harmonic-motion.docx");
        using var archive = ZipFile.OpenRead(path);
        using var source = archive.GetEntry("word/embeddings/oleObject1.bin")!.Open();
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        var bytes = buffer.ToArray();

        var result = MathTypeEquationDecoder.DecodeOle(bytes.Take(bytes.Length / 2).ToArray());
        Assert.Equal(MathTypeParseStatus.Corrupt, result.Status);
        Assert.False(result.CanConvert);
    }
}
