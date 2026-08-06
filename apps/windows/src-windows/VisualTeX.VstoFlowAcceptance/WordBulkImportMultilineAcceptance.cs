using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportMultiline(
        VisualTeXSessionClient client,
        string artifactRoot,
        string objectMode)
    {
        Directory.CreateDirectory(artifactRoot);
        var normalizedMode = string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase)
            ? "ole"
            : "omml";
        var sourcePath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-multiline-{normalizedMode}.tex");
        var logPath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-multiline-{normalizedMode}.log");
        var outputPath = Path.Combine(
            artifactRoot,
            $"word-bulk-import-multiline-{normalizedMode}.docx");
        const string source = """
            \begin{proof}[Completeness]
            The following formulas are imported as one document. Ordinary prose stays upright; \textit{intentional italic stays italic}.
            \end{proof}

            \begin{equation*}
            \lim_{N\to +\infty}\bigl\|\,|\psi\rangle-|s_N\rangle\,\bigr\|
            =
            \lim_{N\to +\infty}
            \left\|\,|\psi\rangle-\sum_{i=1}^{N} c_i |u_i\rangle\,\right\|
            =0.
            \end{equation*}

            \begin{align*}
            a &= b+c \\
            d &= e-f \\
            g &= h
            \end{align*}

            \[
            \qty(\frac{a}{b})+\dv[2]{f}{x}+\pdv{g}{y}
            +\SI[round-mode=places]{3}{\meter\per\second}+\mathbbm{1}_{A}
            \]
            """;
        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(outputPath);

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            normalizedMode == "ole"
                ? WordBulkFormulaObjectMode.Ole
                : WordBulkFormulaObjectMode.Omml);
        AssertEqual(WordBulkSourceFormat.Latex, parsed.SourceFormat,
            "The multiline bulk fixture was not auto-detected as LaTeX.");
        AssertEqual(3, parsed.DisplayFormulaCount,
            "The multiline bulk fixture should contain three display formulas.");
        var parsedFormulas = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .ToArray();
        AssertEqual(3, parsedFormulas.Length,
            "The multiline bulk parser lost or duplicated a formula.");
        AssertTrue(parsedFormulas.All(run => run.Latex.IndexOf('\n') < 0),
            "Physical source newlines survived formula normalization.");
        AssertTrue(parsedFormulas[1].Latex.Contains("a &= b+c \\\\ d &= e-f \\\\ g &= h"),
            "The normalized align formula lost an explicit mathematical row.");

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", normalizedMode);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

        Word.Application? application = null;
        Word.Document? document = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        string? editSessionId = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            // Reproduce the user-visible failure deliberately. Word stores direct
            // character formatting on an empty paragraph mark; a bulk import that
            // omits explicit false values then turns every ordinary run italic.
            Word.Font? contaminatedTypingFont = null;
            try
            {
                contaminatedTypingFont = application.Selection.Font;
                contaminatedTypingFont.Bold = 1;
                contaminatedTypingFont.Italic = 1;
                contaminatedTypingFont.StrikeThrough = 1;
                contaminatedTypingFont.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
            finally { Release(contaminatedTypingFont); }

            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));

            var service = new WordFormulaService(application);
            AssertBulkImportTextFormattingIsSourceControlled(document, application, normalizedMode);
            var imported = ReadOrderedBulkMetadataAndAssertRouting(
                service,
                document,
                normalizedMode);
            AssertEqual(3, imported.Count,
                $"The {normalizedMode} multiline import did not create three editable formulas.");

            var equationMetadata = imported.Single(item =>
                item.Latex.IndexOf("\\lim_{N\\to +\\infty}", StringComparison.Ordinal) >= 0);
            AssertTrue(equationMetadata.Latex.IndexOf('\n') < 0,
                "The equation* metadata still contains physical source rows.");
            AssertTrue(equationMetadata.Latex.IndexOf("\\sum_{i=1}^{N}", StringComparison.Ordinal) >= 0,
                "The equation* metadata lost the finite expansion term.");
            AssertTrue(equationMetadata.Latex.EndsWith("=0.", StringComparison.Ordinal),
                "The equation* metadata lost its final equality.");

            var alignMetadata = imported.Single(item =>
                item.Latex.IndexOf("\\begin{aligned}", StringComparison.Ordinal) >= 0);
            AssertEqual(1, alignMetadata.Lines.Count,
                "A complete aligned environment should remain one logical FormulaLine.");
            AssertTrue(alignMetadata.Latex.IndexOf('\n') < 0,
                "The aligned metadata still contains physical source newlines.");
            AssertTrue(alignMetadata.Latex.IndexOf("a &= b+c", StringComparison.Ordinal) >= 0,
                "The aligned metadata lost its first row.");
            AssertTrue(alignMetadata.Latex.IndexOf("d &= e-f", StringComparison.Ordinal) >= 0,
                "The aligned metadata lost its second row.");
            AssertTrue(alignMetadata.Latex.IndexOf("g &= h", StringComparison.Ordinal) >= 0,
                "The aligned metadata lost its third row.");

            var packageMetadata = imported.Single(item =>
                item.Latex.IndexOf("\\dv[2]{f}{x}", StringComparison.Ordinal) >= 0);
            AssertTrue(packageMetadata.Latex.IndexOf("\\qty", StringComparison.Ordinal) >= 0
                && packageMetadata.Latex.IndexOf("\\SI[round-mode=places]", StringComparison.Ordinal) >= 0
                && packageMetadata.Latex.IndexOf("\\mathbbm", StringComparison.Ordinal) >= 0,
                "Package commands were not preserved in editable metadata.");

            if (normalizedMode == "ole")
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                try
                {
                    shape = document.InlineShapes[2];
                    range = shape.Range;
                    range.Select();
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
            else
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = document.OMaths[2];
                    range = math.Range;
                    range.Select();
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            var existingSessions = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            editSessionId = WaitForNewSession(
                existingSessions,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            var editableLatex = string.Join("\n", editSession.Lines.Select(line => line.Latex));
            AssertEqual("edit", editSession.Mode,
                "The imported aligned formula did not open an edit Session.");
            AssertEqual(
                normalizedMode == "ole"
                    ? FormulaOleContract.NativeOleMode
                    : FormulaOleContract.WordOmmlMode,
                editSession.ObjectMode,
                "The imported aligned formula reopened with the wrong object mode.");
            AssertTrue(editableLatex.IndexOf("a &= b+c", StringComparison.Ordinal) >= 0,
                "The edit Session lost the first aligned row.");
            AssertTrue(editableLatex.IndexOf("d &= e-f", StringComparison.Ordinal) >= 0,
                "The edit Session lost the second aligned row.");
            AssertTrue(editableLatex.IndexOf("g &= h", StringComparison.Ordinal) >= 0,
                "The edit Session lost the third aligned row.");
            AssertTrue(editableLatex.IndexOf("\\end{aligned}", StringComparison.Ordinal) >= 0,
                "The edit Session lost the aligned environment terminator.");

            var ready = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual(false, ready.Dirty,
                "The imported aligned formula became dirty before user input.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            var closed = WaitForTerminal(
                client,
                editSessionId,
                TimeSpan.FromSeconds(30));
            AssertEqual("completed", closed.Status,
                closed.Error ?? "The unchanged aligned editor did not close cleanly.");
            editSessionId = null;
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(15));

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"Word multiline bulk acceptance passed for {normalizedMode}: "
                + "equation* source lines collapsed, all aligned rows survived metadata/edit reopening, "
                + "the editor Session retained the complete aligned environment, and package commands stayed editable.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", null);
            if (!string.IsNullOrWhiteSpace(editSessionId))
            {
                try
                {
                    client.CloseEditorAsync(editSessionId!, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
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
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertBulkImportTextFormattingIsSourceControlled(
        Word.Document document,
        Word.Application application,
        string objectMode)
    {
        AssertWordTextFormatting(
            document,
            "证明（Completeness）：",
            expectBold: true,
            expectItalic: false,
            $"{objectMode} proof heading inherited italic formatting from the insertion point.");
        AssertWordTextFormatting(
            document,
            "The following formulas are imported as one document.",
            expectBold: false,
            expectItalic: false,
            $"{objectMode} ordinary proof prose inherited emphasis formatting from the insertion point.");
        AssertWordTextFormatting(
            document,
            "Ordinary prose stays upright;",
            expectBold: false,
            expectItalic: false,
            $"{objectMode} ordinary prose did not remain upright.");
        AssertWordTextFormatting(
            document,
            "intentional italic stays italic",
            expectBold: false,
            expectItalic: true,
            $"{objectMode} explicit LaTeX italic formatting was lost.");

        const string typingProbe = "VT_UPRIGHT_TYPING_PROBE";
        Word.Selection? selection = null;
        Word.Range? probeRange = null;
        Word.Font? typingFont = null;
        try
        {
            selection = application.Selection;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var probeStart = selection.Start;
            selection.TypeText(typingProbe);
            probeRange = document.Range(probeStart, probeStart + typingProbe.Length);
            typingFont = probeRange.Font;
            AssertEqual(0, typingFont.Bold,
                $"{objectMode} bulk import made subsequent Word typing bold.");
            AssertEqual(0, typingFont.Italic,
                $"{objectMode} bulk import made subsequent Word typing italic.");
            AssertEqual(0, typingFont.StrikeThrough,
                $"{objectMode} bulk import made subsequent Word typing struck through.");
            AssertEqual(Word.WdUnderline.wdUnderlineNone, typingFont.Underline,
                $"{objectMode} bulk import made subsequent Word typing underlined.");
            probeRange.Delete();
            selection.SetRange(probeStart, probeStart);
        }
        finally
        {
            Release(typingFont);
            Release(probeRange);
            Release(selection);
        }
    }

    private static void AssertWordTextFormatting(
        Word.Document document,
        string text,
        bool expectBold,
        bool expectItalic,
        string failureMessage)
    {
        Word.Range? range = null;
        Word.Find? find = null;
        Word.Font? font = null;
        try
        {
            range = document.Content.Duplicate;
            find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            AssertTrue(find.Execute(), $"Could not locate imported Word text: {text}");
            font = range.Font;
            AssertEqual(expectBold, font.Bold != 0, failureMessage);
            AssertEqual(expectItalic, font.Italic != 0, failureMessage);
            AssertEqual(0, font.StrikeThrough, failureMessage);
            AssertEqual(Word.WdUnderline.wdUnderlineNone, font.Underline, failureMessage);
        }
        finally
        {
            Release(font);
            Release(find);
            Release(range);
        }
    }
}
