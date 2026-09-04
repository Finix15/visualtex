using VisualTeX.MathTypeConversion;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.VstoShared;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeConversionTests
{
    [Fact]
    public void TabPositionedRegressionFormulasAreClassifiedInline()
    {
        var path = Path.Combine(
            AppContext.BaseDirectory,
            "Fixtures",
            "MathType7",
            "mathtype7-harmonic-motion.docx");
        using var archive = ZipFile.OpenRead(path);
        using var documentStream = archive.GetEntry("word/document.xml")!.Open();
        var document = XDocument.Load(documentStream);
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace office = "urn:schemas-microsoft-com:office:office";
        var targets = new HashSet<int> { 548, 568, 630, 645, 659 };
        var found = new HashSet<int>();
        var ordinal = 0;

        foreach (var ole in document.Descendants(office + "OLEObject"))
        {
            ordinal++;
            if (!targets.Contains(ordinal)) continue;
            var paragraph = ole.Ancestors(word + "p").First();
            var owningObject = ole.Ancestors(word + "object").First();
            var before = true;
            var prefix = new StringBuilder();
            var suffix = new StringBuilder();
            AppendParagraphLayout(paragraph, owningObject, ref before, prefix, suffix, word);
            suffix.Append('\r');

            var display = WordMathTypeLayoutValidation.IsStandaloneDisplayCandidate(
                prefix.ToString(),
                suffix.ToString(),
                inlineShapeCount: 1,
                existingMathCount: 0,
                out var reason);

            Assert.False(display);
            Assert.Equal("inline-tab-layout", reason);
            Assert.Contains('\t', prefix.ToString());
            found.Add(ordinal);
        }

        Assert.Equal(targets, found);
    }

    [Fact]
    public void PackageSnapshotReadsAllMathTypeOleObjectsInDocumentOrder()
    {
        var path = Path.Combine(
            AppContext.BaseDirectory,
            "Fixtures",
            "MathType7",
            "mathtype7-harmonic-motion.docx");

        var snapshot = MathTypeDocumentScanner.ReadPackageSnapshot(path);

        Assert.Equal(Path.GetFullPath(path), snapshot.SourcePath);
        Assert.Equal(64, snapshot.DocumentFingerprint.Length);
        Assert.Equal(
            MathTypeDocumentScanner.ComputeFileFingerprint(path),
            snapshot.DocumentFingerprint);
        Assert.Equal(680, snapshot.Items.Count);
        Assert.Equal(Enumerable.Range(1, 680), snapshot.Items.Select(item => item.Index));
        Assert.All(snapshot.Items, item =>
        {
            Assert.Equal("Equation.DSMT4", item.ProgId);
            Assert.StartsWith("word/embeddings/oleObject", item.PartName, StringComparison.Ordinal);
            Assert.EndsWith(".bin", item.PartName, StringComparison.OrdinalIgnoreCase);
            Assert.NotEmpty(item.OleBytes);
            Assert.Equal(64, item.OleFingerprint.Length);
            Assert.Equal(
                MathTypeSidecarClient.ComputeFingerprint(item.OleBytes),
                item.OleFingerprint,
                ignoreCase: true);
        });
    }

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
    public void BackupWithStaleScanFingerprintFailsClosedAndRemovesCopy()
    {
        var source = Path.Combine(
            AppContext.BaseDirectory,
            "Fixtures",
            "MathType7",
            "mathtype7-harmonic-motion.docx");
        var backup = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-backup-test-{Guid.NewGuid():N}.docx");

        Assert.Throws<InvalidDataException>(() =>
            MathTypeBackupVerifier.CopyAndVerify(
                source,
                backup,
                680,
                new string('0', 64)));
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

    private static void AppendParagraphLayout(
        XElement element,
        XElement owningObject,
        ref bool before,
        StringBuilder prefix,
        StringBuilder suffix,
        XNamespace word)
    {
        foreach (var child in element.Elements())
        {
            if (ReferenceEquals(child, owningObject))
            {
                before = false;
                continue;
            }
            if (child.Name == word + "object")
            {
                (before ? prefix : suffix).Append('\u0001');
                continue;
            }
            if (child.Name == word + "t")
            {
                (before ? prefix : suffix).Append(child.Value);
                continue;
            }
            if (child.Name == word + "tab")
            {
                (before ? prefix : suffix).Append('\t');
                continue;
            }
            if (child.Name == word + "br" || child.Name == word + "cr")
            {
                (before ? prefix : suffix).Append('\v');
                continue;
            }
            AppendParagraphLayout(child, owningObject, ref before, prefix, suffix, word);
        }
    }
}
