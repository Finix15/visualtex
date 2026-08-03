using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Extensibility;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportLatexSpacing(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(artifactRoot, "word-bulk-import-latex-spacing.tex");
        var logPath = Path.Combine(artifactRoot, "word-bulk-import-latex-spacing.log");
        var outputPath = Path.Combine(artifactRoot, "word-bulk-import-latex-spacing.docx");
        const string source =
            "\\section{LaTeX 行内间距语义}\r\n"
            + "中文 $x=1$ 中文；English $y=2$ words；"
            + "中文\\ $z=3$\\ 中文；中文~$q=4$~中文。";
        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(outputPath);

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);
        AssertEqual(4, parsed.InlineFormulaCount,
            "LaTeX spacing fixture formula count changed unexpectedly.");
        var parsedVisible = string.Concat(
            parsed.Blocks
                .Where(block => block.Kind == WordBulkBlockKind.Paragraph)
                .SelectMany(block => block.Runs)
                .Select(run => run.IsFormula ? $"<{run.Latex}>" : run.Text));
        AssertEqual(
            "中文<x=1>中文；English <y=2> words；"
            + "中文\u00A0<z=3>\u00A0中文；中文\u00A0<q=4>\u00A0中文。",
            parsedVisible,
            "The native Word parser did not preserve LaTeX spacing semantics.");

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "omml");
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

            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(2));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(15));

            var service = new WordFormulaService(application);
            var imported = ReadOrderedBulkMetadataAndAssertRouting(service, document, "omml");
            AssertEqual(4, imported.Count,
                "LaTeX spacing import did not create four editable OMML formulas.");
            AssertNoInlineBoundaryBookmarks(
                document,
                imported.Select(item => item.FormulaId));
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
            AssertLatexSpacingDocumentXml(outputPath);

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: true,
                AddToRecentFiles: false);
            var reopenedMetadata = ReadOrderedBulkMetadataAndAssertRouting(
                service,
                reopened,
                "omml");
            AssertEqual(4, reopenedMetadata.Count,
                "LaTeX spacing formulas changed after save and reopen.");
            AssertNoInlineBoundaryBookmarks(
                reopened,
                reopenedMetadata.Select(item => item.FormulaId));
            AssertLatexSpacingDocumentXml(outputPath);

            Console.WriteLine(
                "LaTeX inline spacing acceptance passed: ordinary CJK boundary spaces "
                + "were suppressed, English word spaces and explicit LaTeX spacing were preserved, "
                + "no VTBL bookmark or boundary character remained.");
            Console.WriteLine($"Artifact: {outputPath}");
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

    private static void AssertLatexSpacingDocumentXml(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("The Word document has no word/document.xml part.");
        using var stream = entry.Open();
        var xml = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";

        var rawXml = xml.ToString(SaveOptions.DisableFormatting);
        AssertTrue(rawXml.IndexOf('\u200B') < 0,
            "The Word document still contains the legacy U+200B inline guard.");
        AssertTrue(rawXml.IndexOf('\u2060') < 0,
            "The Word document still contains the legacy U+2060 inline sentinel.");
        AssertTrue(rawXml.IndexOf('\uE000') < 0,
            "The Word document still contains a VisualTeX private-use placeholder.");

        var paragraph = xml
            .Descendants(word + "p")
            .SingleOrDefault(candidate => candidate.Elements(math + "oMath").Count() == 4)
            ?? throw new InvalidDataException(
                "The LaTeX spacing paragraph does not contain exactly four direct inline equations.");
        var visible = new StringBuilder();
        var hiddenBoundaryCharacterCount = 0;
        var formulaIndex = 0;
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == math + "oMath")
            {
                formulaIndex++;
                visible.Append($"<M{formulaIndex}>");
                continue;
            }
            if (child.Name != word + "r") continue;
            var runText = string.Concat(child.Descendants(word + "t").Select(text => text.Value));
            var hidden = child.Element(word + "rPr")?.Element(word + "vanish") is not null;
            if (hidden && runText.Length > 0)
                hiddenBoundaryCharacterCount++;
            visible.Append(runText);
        }

        AssertEqual(0, hiddenBoundaryCharacterCount,
            "The Word document still contains a hidden character after an inline formula.");
        AssertEqual(
            "中文<M1>中文；English <M2> words；"
            + "中文\u00A0<M3>\u00A0中文；中文\u00A0<M4>\u00A0中文。",
            visible.ToString(),
            "Word-visible inline spacing no longer matches LaTeX semantics.");
    }

    private static void AssertNoInlineBoundaryBookmarks(
        Word.Document document,
        IEnumerable<string> formulaIds)
    {
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formulaId in formulaIds)
            {
                var name = "VTBL_" + Guid.Parse(formulaId).ToString("N");
                AssertTrue(!bookmarks.Exists(name),
                    $"The native OMML formula still exposes the temporary boundary bookmark {name}.");
            }
        }
        finally { Release(bookmarks); }
    }
}
