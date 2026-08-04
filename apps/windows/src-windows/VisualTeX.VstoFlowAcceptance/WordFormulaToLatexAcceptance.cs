using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class FormulaToLatexCase
    {
        internal string FormulaId { get; set; } = string.Empty;
        internal string ObjectMode { get; set; } = string.Empty;
        internal string DisplayMode { get; set; } = string.Empty;
        internal string Latex { get; set; } = string.Empty;
        internal bool Numbered { get; set; }
    }

    private static void RunWordFormulaToLatex(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-Formula-To-Latex.docx");
        var logPath = Path.Combine(
            artifactRoot,
            "word-formula-to-latex.log");
        TryDeleteAcceptanceFile(documentPath);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
            logPath);
        const string referenceBookmark = "VTTestEquationNumberReference";
        try
        {
            using (var host = new WordPerformanceHost(documentPath: null))
            {
                var formulas = PopulateFormulaToLatexDocument(client, host);
                var oleInline = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.NativeOleMode
                    && item.DisplayMode == "inline");
                var ommlInline = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.WordOmmlMode
                    && item.DisplayMode == "inline");
                var oleDisplay = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.NativeOleMode
                    && item.DisplayMode == "block");

                // Edit the last OMML equation through Word's native equation
                // model before restoring it. The reverse command must export the
                // current Word content rather than stale VisualTeX metadata.
                AppendToLastWordOmmlAndSelect(host.Document, "+5");

                var insertedReference = InsertNumberingReference(
                    host.Application,
                    host.Document,
                    oleDisplay.FormulaId);
                AssertEqual(referenceBookmark, insertedReference,
                    "The equation reference acceptance bookmark changed unexpectedly.");
                AppendFormulaToLatexDocumentEnd(host.Application);
                AssertReferenceText(host.Document, referenceBookmark, "(1)");
                AssertFormulaObjectCounts(host.Document, expectedOle: 2, expectedOmml: 2);

                SelectFormula(host, oleInline);
                host.AddIn.OnRedrawSelectionOleToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 1);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 2);
                AssertDocumentContains(host.Document, "$a=1$");
                AssertDocumentDoesNotContain(host.Document, "$b=2$");

                SelectFormula(host, ommlInline);
                host.AddIn.OnRedrawSelectionOmmlToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 2);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 1);
                AssertDocumentContains(host.Document, "$b=2$");

                CollapseSelectionAtDocumentStart(host);
                host.AddIn.OnRedrawDocumentOleToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 3);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 0, expectedOmml: 1);
                AssertDocumentContains(host.Document, "$$c=3$$");
                AssertReferenceText(host.Document, referenceBookmark, "(1)");
                AssertNoBrokenReferenceText(host.Document);

                CollapseSelectionAtDocumentStart(host);
                host.AddIn.OnRedrawDocumentOmmlToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 4);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFinalFormulaToLatexDocument(host.Document, referenceBookmark);
                host.Save(documentPath);
            }

            using (var reopened = new WordPerformanceHost(documentPath))
            {
                AssertFinalFormulaToLatexDocument(
                    reopened.Document,
                    referenceBookmark);
            }

            Console.WriteLine(
                "Word formula-to-LaTeX acceptance passed: selection and whole-document Ribbon commands restored OLE/OMML formulas to source code independently, numbered tables were flattened, references were preserved as plain text, prose order survived, and the result persisted after save/reopen.");
            Console.WriteLine($"Artifact: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
                null);
        }
    }

    private static List<FormulaToLatexCase> PopulateFormulaToLatexDocument(
        VisualTeXSessionClient client,
        WordPerformanceHost host)
    {
        Word.Selection? selection = null;
        try
        {
            selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Font.Name = "宋体";
            selection.Font.Size = 10.5f;
            selection.TypeText("BEGIN ");
            var oleInline = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.NativeOleMode,
                displayMode: "inline",
                numbered: false,
                latex: "a=1",
                mathVariable: "a",
                mathNumber: "1");
            selection.TypeText(" BETWEEN ");
            var ommlInline = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "inline",
                numbered: false,
                latex: "b=2",
                mathVariable: "b",
                mathNumber: "2");
            selection.TypeText(" INLINE_END");
            selection.TypeParagraph();

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("BEFORE_OLE_DISPLAY");
            selection.TypeParagraph();
            var oleDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.NativeOleMode,
                displayMode: "block",
                numbered: true,
                latex: "c=3",
                mathVariable: "c",
                mathNumber: "3");
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("AFTER_OLE_DISPLAY");
            selection.TypeParagraph();

            var ommlDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "block",
                numbered: true,
                latex: "d=4",
                mathVariable: "d",
                mathNumber: "4");
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("AFTER_OMML_DISPLAY");
            selection.TypeParagraph();

            return new List<FormulaToLatexCase>
            {
                oleInline,
                ommlInline,
                oleDisplay,
                ommlDisplay,
            };
        }
        finally { Release(selection); }
    }

    private static FormulaToLatexCase InsertFormulaToLatexCase(
        VisualTeXSessionClient client,
        WordPerformanceHost host,
        string objectMode,
        string displayMode,
        bool numbered,
        string latex,
        string mathVariable,
        string mathNumber)
    {
        var existing = SnapshotSessionIds();
        var display = string.Equals(displayMode, "block", StringComparison.Ordinal);
        if (string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
        {
            if (display) host.AddIn.OnInsertDisplayOmml(new object());
            else host.AddIn.OnInsertInlineOmml(new object());
        }
        else
        {
            if (display) host.AddIn.OnInsertDisplay(new object());
            else host.AddIn.OnInsertInline(new object());
        }

        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        var mathMl = string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal)
            ? $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"{(display ? "block" : "inline")}\"><mi>{mathVariable}</mi><mo>=</mo><mn>{mathNumber}</mn></math>"
            : null;
        Commit(
            client,
            session,
            displayMode,
            objectMode,
            latex,
            numbered: numbered,
            mathMl: mathMl);
        var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
        AssertEqual("completed", final.Status,
            final.Error ?? "The formula-to-LaTeX source formula did not complete.");
        client.CloseEditorAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(15));
        return new FormulaToLatexCase
        {
            FormulaId = final.FormulaId
                ?? throw new InvalidDataException(
                    "The formula-to-LaTeX source formula has no formulaId."),
            ObjectMode = objectMode,
            DisplayMode = displayMode,
            Latex = latex,
            Numbered = numbered,
        };
    }

    private static void AppendFormulaToLatexDocumentEnd(Word.Application application)
    {
        Word.Selection? selection = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("DOCUMENT_END");
        }
        finally { Release(selection); }
    }

    private static void SelectFormula(
        WordPerformanceHost host,
        FormulaToLatexCase formula)
    {
        Word.Range? range = null;
        try
        {
            range = ResolveFormulaToLatexRange(host.Document, formula);
            host.Application.Selection.SetRange(range.Start, range.End);
        }
        finally { Release(range); }
    }

    private static Word.Range ResolveFormulaToLatexRange(
        Word.Document document,
        FormulaToLatexCase formula)
    {
        if (string.Equals(
                formula.ObjectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"OMML formula {formula.FormulaId} is missing.");
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                return range.Duplicate;
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }

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
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formula.FormulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    range = shape.Range;
                    return range.Duplicate;
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally { Release(shapes); }
        throw new InvalidDataException(
            $"OLE formula {formula.FormulaId} is missing.");
    }

    private static void CollapseSelectionAtDocumentStart(WordPerformanceHost host)
    {
        Word.Range? content = null;
        try
        {
            content = host.Document.Content;
            host.Application.Selection.SetRange(content.Start, content.Start);
        }
        finally { Release(content); }
    }

    private static string WaitForFormulaToLatex(
        string logPath,
        int expectedCompletions)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromMinutes(2);
        var last = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(25);
            try
            {
                if (!File.Exists(logPath)) continue;
                last = File.ReadAllText(logPath, Encoding.UTF8);
                if (last.IndexOf(
                        "formula-to-latex-failed",
                        StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        "Word formula-to-LaTeX command failed.\n" + last);
                var completed = last.Split(new[] { "formula-to-latex-complete" },
                    StringSplitOptions.None).Length - 1;
                if (completed >= expectedCompletions) return last;
            }
            catch (IOException)
            {
                // The add-in may be appending the current log line.
            }
        }
        throw new TimeoutException(
            $"Word formula-to-LaTeX command did not complete. Expected {expectedCompletions} completion records. Last log:\n{last}");
    }

    private static void AssertFormulaObjectCounts(
        Word.Document document,
        int expectedOle,
        int expectedOmml)
    {
        Word.InlineShapes? shapes = null;
        var oleCount = 0;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (WordFormulaMetadataReader.IsNativeOle(shape)
                        && WordFormulaMetadataReader.TryRead(shape) is not null)
                        oleCount++;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        var ommlCount = WordOmmlFormulaStore.FormulaIds(document).Count;
        AssertEqual(expectedOle, oleCount,
            "The formula-to-LaTeX command converted the wrong number of OLE formulas.");
        AssertEqual(expectedOmml, ommlCount,
            "The formula-to-LaTeX command converted the wrong number of OMML formulas.");
    }

    private static void AssertFinalFormulaToLatexDocument(
        Word.Document document,
        string referenceBookmark)
    {
        AssertFormulaObjectCounts(document, expectedOle: 0, expectedOmml: 0);
        foreach (var required in new[]
                 {
                     "BEGIN",
                     "$a=1$",
                     "BETWEEN",
                     "$b=2$",
                     "INLINE_END",
                     "BEFORE_OLE_DISPLAY",
                     "$$c=3$$",
                     "AFTER_OLE_DISPLAY",
                     "$$d=4+5$$",
                     "AFTER_OMML_DISPLAY",
                     "Reference:",
                     "DOCUMENT_END",
                 })
            AssertDocumentContains(document, required);

        var text = ReadFormulaToLatexDocumentText(document);
        var ordered = new[]
        {
            "BEGIN",
            "$a=1$",
            "BETWEEN",
            "$b=2$",
            "INLINE_END",
            "BEFORE_OLE_DISPLAY",
            "$$c=3$$",
            "AFTER_OLE_DISPLAY",
            "$$d=4+5$$",
            "AFTER_OMML_DISPLAY",
            "Reference:",
            "DOCUMENT_END",
        };
        var previous = -1;
        foreach (var marker in ordered)
        {
            var position = text.IndexOf(marker, StringComparison.Ordinal);
            if (position <= previous)
                throw new InvalidDataException(
                    $"Formula-to-LaTeX document order is wrong at '{marker}'.\n{text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}");
            previous = position;
        }

        Word.Tables? tables = null;
        try
        {
            tables = document.Tables;
            AssertEqual(0, tables.Count,
                "Numbered formula tables remained after all formulas were restored to LaTeX.");
        }
        finally { Release(tables); }
        AssertReferenceText(document, referenceBookmark, "(1)");
        AssertNoBrokenReferenceText(document);
    }

    private static void AssertDocumentContains(
        Word.Document document,
        string expected)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        if (text.IndexOf(expected, StringComparison.Ordinal) < 0)
            throw new InvalidDataException(
                $"Formula-to-LaTeX document is missing '{expected}'.\n{text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}");
    }

    private static void AssertDocumentDoesNotContain(
        Word.Document document,
        string forbidden)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        if (text.IndexOf(forbidden, StringComparison.Ordinal) >= 0)
            throw new InvalidDataException(
                $"Formula-to-LaTeX document unexpectedly contains '{forbidden}'.");
    }

    private static void AssertNoBrokenReferenceText(Word.Document document)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        foreach (var forbidden in new[]
                 {
                     "Error! Reference source not found.",
                     "错误! 未找到引用源。",
                     "错误！未找到引用源。",
                 })
        {
            if (text.IndexOf(forbidden, StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidDataException(
                    $"Formula-to-LaTeX conversion left a broken equation reference: {forbidden}");
        }
    }

    private static string ReadFormulaToLatexDocumentText(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            return content.Text ?? string.Empty;
        }
        finally { Release(content); }
    }
}
