using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportOleViewBox(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(artifactRoot, "word-bulk-import-ole-viewbox.tex");
        var logPath = Path.Combine(artifactRoot, "word-bulk-import-ole-viewbox.log");
        var outputPath = Path.Combine(artifactRoot, "word-bulk-import-ole-viewbox.docx");
        const string source = """
            \textbf{二阶线性常微分算子的形式伴随算子}

            在取权函数 \(\rho = 1\) 的情况下，二阶线性常微分算子 \(L\) 的形式伴随算子 \(L^\dagger\) 具有形式
            \begin{equation}
            L^\dagger
            =
            p_2(x)\frac{d^2}{dx^2}
            +
            \bigl[2p_2'(x)-p_1(x)\bigr]\frac{d}{dx}
            +
            \bigl[p_2''(x)-p_1'(x)+p_0(x)\bigr].
            \tag{9.27}
            \end{equation}

            不难验证 \(L\) 的形式伴随的形式伴随正是 \(L\) 自身，即
            \[
            \left(L^\dagger\right)^\dagger = L.
            \]

            \(L\) 和 \(L^\dagger\) 满足 Lagrange 恒等式：\(\forall \langle f\rangle,\langle g\rangle\)，有
            \begin{equation}
            \langle f|L|g\rangle - \langle g|L^\dagger|f\rangle
            =
            Q[f^*,g]\Big|_a^b.
            \tag{9.28}
            \end{equation}

            此处 \(Q[f^*,g]\) 被称为函数 \(|f\rangle\) 和 \(|g\rangle\) 的结合式：
            \begin{equation}
            Q[f^*,g]
            =
            p_2(x)\left[
            f^*(x)\frac{d}{dx}g(x)
            -
            \frac{d}{dx}f^*(x)\,g(x)
            \right]
            +
            \left[
            p_1(x)-\frac{d}{dx}p_2(x)
            \right]
            f^*(x)g(x).
            \tag{9.29}
            \end{equation}
            \section{伴随边界条件}
            二阶微分算子
            \begin{gather*}
            L=-\frac{d^{2}}{dx^{2}}\\
            \end{gather*}在齐次第三类边界条件和扭曲的周期性边界条件都是自伴的

            求自伴的原因是我们需要用到本征值相互正交和实数性，但是一个无穷维的一般算子不一定有本征值和本征向量
            """;

        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(outputPath);

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Ole);
        AssertEqual(WordBulkSourceFormat.Latex, parsed.SourceFormat,
            "The OLE viewBox fixture was not auto-detected as LaTeX.");
        AssertEqual(16, parsed.FormulaCount,
            "The OLE viewBox fixture should contain sixteen formulas.");
        AssertEqual(5, parsed.DisplayFormulaCount,
            "The OLE viewBox fixture should contain five display formulas.");
        var parsedTags = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula && !string.IsNullOrWhiteSpace(run.EquationTag))
            .Select(run => run.EquationTag)
            .ToArray();
        AssertEqual(3, parsedTags.Length,
            "The OLE viewBox fixture should contain three explicit equation tags.");
        AssertTrue(parsedTags.SequenceEqual(new[] { "9.27", "9.28", "9.29" }),
            "The OLE viewBox parser changed the explicit equation tags.");

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "ole");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_CLEANUP_DELAY_MS", "15000");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? repeatDocument = null;
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

            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));

            // The first import deliberately keeps companion Session cleanup
            // sleeping for 15 seconds. Word's operation gate must already be
            // available, so a second import into another document can start and
            // finish immediately instead of reporting a permanently busy add-in.
            repeatDocument = application.Documents.Add();
            repeatDocument.Activate();
            addIn.OnBulkImport(new object());
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            AssertEqual(16, repeatDocument.InlineShapes.Count,
                "A second OLE bulk import could not start while the first import's companion cleanup was still pending.");
            repeatDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(repeatDocument);
            repeatDocument = null;
            document.Activate();
            Console.WriteLine(
                "    repeated bulk import passed while the first Session cleanup was deliberately delayed by 15 seconds");

            var service = new WordFormulaService(application);
            var imported = ReadOrderedBulkMetadataAndAssertRouting(service, document, "ole");
            AssertEqual(16, imported.Count,
                "The exact OLE document import did not create sixteen editable formulas.");
            AssertEqual(16, document.InlineShapes.Count,
                "The exact OLE document import did not create sixteen inline OLE shapes.");

            var importedTags = imported
                .Where(metadata => !string.IsNullOrWhiteSpace(metadata.EquationTag))
                .Select(metadata => metadata.EquationTag)
                .ToArray();
            AssertTrue(importedTags.SequenceEqual(new[] { "9.27", "9.28", "9.29" }),
                "OLE metadata lost or reordered an explicit equation tag.");

            foreach (var tagged in imported
                         .Select((metadata, index) => new { Metadata = metadata, Index = index })
                         .Where(item => !string.IsNullOrWhiteSpace(item.Metadata.EquationTag)))
            {
                Word.InlineShape? taggedShape = null;
                try
                {
                    taggedShape = document.InlineShapes[tagged.Index + 1];
                    AssertTrue(taggedShape.Width > 180f,
                        $"Tagged OLE formula {tagged.Metadata.EquationTag} was collapsed to an invalid narrow preview ({taggedShape.Width:0.###} pt).");
                    AssertTrue(taggedShape.Height > 20f,
                        $"Tagged OLE formula {tagged.Metadata.EquationTag} was collapsed to an invalid short preview ({taggedShape.Height:0.###} pt).");
                    Console.WriteLine(
                        $"    tagged OLE {tagged.Metadata.EquationTag}: {taggedShape.Width:0.###}x{taggedShape.Height:0.###} pt");
                }
                finally
                {
                    Release(taggedShape);
                }
            }

            var longestTagged = imported
                .Select((metadata, index) => new { Metadata = metadata, Index = index })
                .Single(item => string.Equals(
                    item.Metadata.EquationTag,
                    "9.29",
                    StringComparison.Ordinal));
            AssertTrue(longestTagged.Metadata.Latex.IndexOf(
                    "p_2(x)\\left[",
                    StringComparison.Ordinal) >= 0,
                "The longest tagged OLE formula lost its editable source.");

            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[longestTagged.Index + 1];
                range = shape.Range;
                range.Select();
            }
            finally
            {
                Release(range);
                Release(shape);
            }

            var existingSessions = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            editSessionId = WaitForNewSession(
                existingSessions,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode,
                "The longest tagged OLE formula did not open an edit Session.");
            AssertEqual(FormulaOleContract.NativeOleMode, editSession.ObjectMode,
                "The longest tagged formula reopened with the wrong object mode.");
            AssertEqual("9.29", editSession.OriginalMetadata?.EquationTag,
                "The tagged OLE edit Session lost equation number 9.29.");
            AssertTrue(string.Join("\n", editSession.Lines.Select(line => line.Latex)).IndexOf(
                    "Q[f^*,g]",
                    StringComparison.Ordinal) >= 0,
                "The tagged OLE edit Session lost its formula source.");

            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            var closed = WaitForTerminal(
                client,
                editSessionId,
                TimeSpan.FromSeconds(30));
            AssertEqual("completed", closed.Status,
                closed.Error ?? "The unchanged tagged OLE editor did not close cleanly.");
            editSessionId = null;
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(15));

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "Word exact OLE viewBox acceptance passed: all sixteen formulas imported, "
                + "all three equation tags survived, and the longest tagged formula reopened for editing.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_CLEANUP_DELAY_MS", null);
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
            try { repeatDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(repeatDocument);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }
}
