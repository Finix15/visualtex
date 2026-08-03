using System.Diagnostics;
using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const int BulkImportPerformanceFormulaCount = 50;
    private const long BulkImportPerformanceLimitMilliseconds = 10_000;

    private static void RunWordBulkImportPerformance(
        VisualTeXSessionClient client,
        string artifactRoot,
        string requestedObjectMode)
    {
        var objectMode = requestedObjectMode.Trim().ToLowerInvariant();
        if (objectMode is not ("omml" or "ole"))
            throw new ArgumentOutOfRangeException(
                nameof(requestedObjectMode),
                requestedObjectMode,
                "Bulk import performance object mode must be 'omml' or 'ole'.");

        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-performance-{objectMode}.md");
        var logPath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-performance-{objectMode}.log");
        var outputPath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-performance-{objectMode}.docx");
        var source = CreateWordBulkImportPerformanceSource();
        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(outputPath);

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            objectMode == "ole"
                ? WordBulkFormulaObjectMode.Ole
                : WordBulkFormulaObjectMode.Omml);
        AssertEqual(
            BulkImportPerformanceFormulaCount,
            parsed.FormulaCount,
            "Bulk performance fixture formula count changed unexpectedly.");
        AssertTrue(
            source.Length >= 1_000,
            $"Bulk performance fixture contains only {source.Length} source characters.");
        var expectedLatex = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "markdown");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", objectMode);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            Console.WriteLine(
                $"Starting real Word bulk import performance acceptance: mode={objectMode}, "
                + $"sourceCharacters={source.Length}, formulas={parsed.FormulaCount}.");
            var wallClock = Stopwatch.StartNew();
            addIn.OnBulkImport(new object());
            var operationMilliseconds = WaitForBulkImportCompletion(
                logPath,
                TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(15));
            wallClock.Stop();

            var service = new WordFormulaService(application);
            var importedMetadata = ReadOrderedBulkMetadataAndAssertRouting(
                service,
                document,
                objectMode);
            AssertEqual(
                expectedLatex.Length,
                importedMetadata.Count,
                "Bulk import produced the wrong number of editable formulas.");
            AssertEqual(
                expectedLatex.Length,
                importedMetadata
                    .Select(item => item.FormulaId)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count(),
                "Bulk import reused a formula ID across multiple formulas.");
            for (var index = 0; index < expectedLatex.Length; index++)
            {
                AssertEqual(
                    expectedLatex[index],
                    importedMetadata[index].Latex,
                    $"Bulk formula {index + 1} is out of order or contains the wrong source.");
            }
            AssertBulkPerformanceSectionOrder(document);

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: true,
                AddToRecentFiles: false);
            var reopenedMetadata = ReadOrderedBulkMetadataAndAssertRouting(
                service,
                reopened,
                objectMode);
            AssertEqual(
                importedMetadata.Count,
                reopenedMetadata.Count,
                "Bulk formula count changed after save and reopen.");
            AssertTrue(
                reopenedMetadata
                    .Select(item => item.FormulaId)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase)
                    .SetEquals(importedMetadata.Select(item => item.FormulaId)),
                "Bulk formula IDs changed after save and reopen.");
            AssertTrue(
                operationMilliseconds <= BulkImportPerformanceLimitMilliseconds,
                $"Bulk {objectMode.ToUpperInvariant()} import took "
                + $"{operationMilliseconds} ms; limit is "
                + $"{BulkImportPerformanceLimitMilliseconds} ms.");

            Console.WriteLine(
                $"Bulk {objectMode.ToUpperInvariant()} performance passed: "
                + $"operation={operationMilliseconds} ms, wall={wallClock.ElapsedMilliseconds} ms, "
                + $"sourceCharacters={source.Length}, formulas={importedMetadata.Count}.");
            Console.WriteLine($"Artifact: {outputPath}");
            Console.WriteLine($"Timing log: {logPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", null);
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { application?.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static string CreateWordBulkImportPerformanceSource()
    {
        var source = new StringBuilder();
        for (var section = 1; section <= 10; section++)
        {
            source.AppendLine($"## Performance section {section:00}");
            source.Append(
                $"Performance section {section:00} keeps a long native-text paragraph before, "
                + "between, and after formulas so the acceptance detects shifted placeholders, "
                + "lost prose, merged runs, and incorrect insertion positions in a realistic document. ");
            for (var formula = 1; formula <= 4; formula++)
            {
                source.Append(
                    $"Inline marker {section:00}-{formula}: "
                    + $"$x_{{{section}{formula}}}+y_{{{section}{formula}}}=z_{{{section}{formula}}}$ ");
                source.Append(
                    "remains surrounded by ordinary editable Word text, punctuation, and spacing. ");
            }
            source.AppendLine(
                "The final sentence confirms that the paragraph continues normally after all inline formulas.");
            source.AppendLine();
            source.AppendLine("$$");
            source.AppendLine(
                $"\\sum_{{k=1}}^{{{section + 10}}} a_{{{section},k}}="
                + $"\\frac{{b_{{{section}}}}}{{c_{{{section}}}+1}}");
            source.AppendLine("$$");
            source.AppendLine();
        }
        return source.ToString();
    }

    private static long WaitForBulkImportCompletion(
        string logPath,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        string lastLog = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(25);
            if (!File.Exists(logPath)) continue;
            try { lastLog = File.ReadAllText(logPath, Encoding.UTF8); }
            catch { continue; }

            var failedLine = lastLog
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault(line => line.IndexOf(
                    "bulk-import-failed",
                    StringComparison.Ordinal) >= 0);
            if (!string.IsNullOrWhiteSpace(failedLine))
                throw new InvalidOperationException(
                    $"Bulk import acceptance failed inside the add-in: {failedLine}");

            var completionLine = lastLog
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault(line => line.IndexOf(
                    "bulk-import-complete",
                    StringComparison.Ordinal) >= 0);
            if (string.IsNullOrWhiteSpace(completionLine)) continue;
            var elapsedToken = completionLine
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault(token => token.StartsWith(
                    "elapsedMs=",
                    StringComparison.Ordinal));
            if (elapsedToken is null
                || !long.TryParse(
                    elapsedToken.Substring("elapsedMs=".Length),
                    out var elapsedMilliseconds))
            {
                throw new InvalidDataException(
                    $"Bulk import completion log does not contain elapsedMs: {completionLine}");
            }
            return elapsedMilliseconds;
        }
        throw new TimeoutException(
            $"Bulk import did not finish within {timeout}. Last log:\n{lastLog}");
    }

    private static List<FormulaMetadata> ReadOrderedBulkMetadataAndAssertRouting(
        WordFormulaService service,
        Word.Document document,
        string objectMode)
    {
        var metadata = new List<FormulaMetadata>();
        if (objectMode == "ole")
        {
            Word.InlineShapes? shapes = null;
            try
            {
                shapes = document.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    Word.Range? range = null;
                    try
                    {
                        shape = shapes[index];
                        range = shape.Range;
                        var item = WordFormulaMetadataReader.TryRead(shape);
                        if (item is null) continue;
                        range.Select();
                        var selection = service.ReadSelection();
                        AssertEqual(
                            item.FormulaId,
                            selection.FormulaId,
                            $"OLE formula {index} cannot be routed back into VisualTeX for editing.");
                        AssertEqual(
                            FormulaOleContract.NativeOleMode,
                            selection.ObjectMode,
                            $"OLE formula {index} reports the wrong object mode.");
                        metadata.Add(item);
                    }
                    finally
                    {
                        Release(range);
                        Release(shape);
                    }
                }
            }
            finally { Release(shapes); }
            return metadata;
        }

        Word.OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                Word.Bookmark? bookmark = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    bookmark = WordOmmlFormulaStore.FindAtRange(document, range);
                    if (bookmark is null) continue;
                    var item = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (item is null) continue;
                    range.Select();
                    var selection = service.ReadSelection();
                    AssertEqual(
                        item.FormulaId,
                        selection.FormulaId,
                        $"OMML formula {index} cannot be routed back into VisualTeX for editing.");
                    AssertEqual(
                        FormulaOleContract.WordOmmlMode,
                        selection.ObjectMode,
                        $"OMML formula {index} reports the wrong object mode.");
                    metadata.Add(item);
                }
                finally
                {
                    Release(bookmark);
                    Release(range);
                    Release(math);
                }
            }
        }
        finally { Release(maths); }
        return metadata;
    }

    private static void DeleteBulkPerformanceArtifact(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch { }
    }

    private static void AssertBulkPerformanceSectionOrder(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            var previous = -1;
            for (var section = 1; section <= 10; section++)
            {
                var marker = $"Performance section {section:00}";
                var current = text.IndexOf(
                    marker,
                    previous + 1,
                    StringComparison.Ordinal);
                AssertTrue(
                    current > previous,
                    $"Native text marker '{marker}' is missing or out of order after bulk import.");
                previous = current;
            }
        }
        finally { Release(content); }
    }
}
