using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const int LatexRedrawPerformanceLimitMilliseconds = 250;

    private static void RunWordLatexRedraw(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        RunWordLatexRedrawSourceContextStress();
        Console.WriteLine("[Word LaTeX redraw] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();

        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.WordOmmlMode,
            wholeDocument: false,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-OMML.docx");
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.NativeOleMode,
            wholeDocument: true,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-OLE.docx");
    }

    private static void RunWordLatexRedrawSourceContextStress()
    {
        using var host = new WordPerformanceHost(documentPath: null);
        var builder = new StringBuilder(120_000);
        var sourceFormats = new List<(int Start, int Length, float Size)>(1000);
        for (var index = 1; index <= 500; index++)
        {
            builder.Append($"第{index}段中文正文用于检查行内公式：");
            var inlineStart = builder.Length;
            var inlineSource = $"$a_{{{index}}}+b_{{{index}}}=c_{{{index}}}$";
            builder.Append(inlineSource);
            sourceFormats.Add((inlineStart, inlineSource.Length, 9.5f));
            builder.Append("。\r");

            var displayStart = builder.Length;
            var displaySource = $@"\[E_{{{index}}}=m_{{{index}}}c^2\]";
            builder.Append(displaySource);
            sourceFormats.Add((displayStart, displaySource.Length, 8.5f));
            builder.Append('\r');
        }

        Word.Range? content = null;
        Microsoft.Office.Interop.Word.Font? contentFont = null;
        try
        {
            content = host.Document.Content;
            var documentStart = content.Start;
            content.Text = builder.ToString();
            contentFont = content.Font;
            contentFont.Name = "宋体";
            contentFont.Size = 10.5f;
            foreach (var sourceFormat in sourceFormats)
            {
                Word.Range? sourceRange = null;
                Microsoft.Office.Interop.Word.Font? sourceFont = null;
                try
                {
                    sourceRange = host.Document.Range(
                        documentStart + sourceFormat.Start,
                        documentStart + sourceFormat.Start + sourceFormat.Length);
                    sourceFont = sourceRange.Font;
                    sourceFont.Size = sourceFormat.Size;
                }
                finally
                {
                    Release(sourceFont);
                    Release(sourceRange);
                }
            }
        }
        finally
        {
            Release(contentFont);
            Release(content);
        }

        var service = new WordFormulaService(host.Application);
        var stopwatch = Stopwatch.StartNew();
        var plan = service.CaptureLatexRedrawPlan(wholeDocument: true);
        stopwatch.Stop();
        if (plan.Targets.Count != 1000)
            throw new InvalidDataException(
                $"1000-formula redraw scan found {plan.Targets.Count} formulas.");
        var inline = plan.Targets
            .Where(target => string.Equals(target.DisplayMode, "inline", StringComparison.Ordinal))
            .ToArray();
        var display = plan.Targets
            .Where(target => string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            .ToArray();
        if (inline.Length != 500 || display.Length != 500)
            throw new InvalidDataException(
                $"1000-formula redraw scan produced inline={inline.Length}, display={display.Length}.");
        var wrongInline = inline.Count(target => Math.Abs(target.FontSizePt - 10.5) > 0.1);
        var wrongDisplay = display.Count(target => Math.Abs(target.FontSizePt - 10.5) > 0.1);
        var unpreservedDisplay = display.Count(target => !target.PreserveDisplayParagraphBoundary);
        if (wrongInline != 0 || wrongDisplay != 0 || unpreservedDisplay != 0)
            throw new InvalidDataException(
                "1000-formula redraw context inheritance failed: "
                + $"wrongInline={wrongInline}, wrongDisplay={wrongDisplay}, "
                + $"unpreservedDisplay={unpreservedDisplay}.");
        if (stopwatch.Elapsed > TimeSpan.FromSeconds(30))
            throw new InvalidDataException(
                $"1000-formula redraw source scan took {stopwatch.Elapsed.TotalSeconds:F2}s.");
        Console.WriteLine(
            "[Word LaTeX redraw] 1000-formula source context scan passed: "
            + $"{stopwatch.ElapsedMilliseconds} ms, inline=500, display=500.");
    }

    private static void RunWordLatexRedrawScenario(
        string artifactRoot,
        string objectMode,
        bool wholeDocument,
        string expectedFileName)
    {
        var modeName = objectMode == FormulaOleContract.NativeOleMode ? "OLE" : "OMML";
        var logPath = Path.Combine(
            artifactRoot,
            $"word-latex-redraw-{modeName.ToLowerInvariant()}.log");
        var documentPath = Path.Combine(artifactRoot, expectedFileName);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", logPath);
        try
        {
            using var host = new WordPerformanceHost(documentPath: null);
            PopulateLatexRedrawDocument(host);
            Word.Range? content = null;
            try
            {
                content = host.Document.Content;
                if (wholeDocument)
                {
                    host.Application.Selection.SetRange(content.Start, content.Start);
                    if (objectMode == FormulaOleContract.NativeOleMode)
                        host.AddIn.OnRedrawDocumentToOle(new object());
                    else
                        host.AddIn.OnRedrawDocumentToOmml(new object());
                }
                else
                {
                    host.Application.Selection.SetRange(content.Start, content.End - 1);
                    if (objectMode == FormulaOleContract.NativeOleMode)
                        host.AddIn.OnRedrawSelectionToOle(new object());
                    else
                        host.AddIn.OnRedrawSelectionToOmml(new object());
                }
            }
            finally { Release(content); }

            var redrawLog = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            host.Save(documentPath);
            AssertLatexRedrawDocument(host.Document, objectMode);
            AssertLatexRedrawPerformance(redrawLog, modeName);
            Console.WriteLine(
                $"[Word LaTeX redraw] {modeName} {(wholeDocument ? "document" : "selection")} redraw passed: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", null);
        }
    }

    private static void PopulateLatexRedrawDocument(WordPerformanceHost host)
    {
        var selection = host.Application.Selection;
        selection.HomeKey(Word.WdUnits.wdStory);

        selection.Font.Name = "宋体";
        selection.Font.Size = 10.5f;
        // A supplementary Unicode character before later formulas reproduces
        // Word's story-coordinate/UTF-16 offset mismatch from real documents.
        // The raw LaTeX run is intentionally smaller than the prose, matching
        // word_10000_chinese_500_inline_500_display.docx.
        selection.TypeText("普通正文🙂前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText("$UVI>2$");
        selection.Font.Size = 10.5f;
        selection.TypeText(" 普通正文后。");
        selection.TypeParagraph();

        selection.Font.Size = 8.5f;
        selection.TypeText(@"\[E=mc^2\]");
        selection.TypeParagraph();

        selection.Font.Size = 16;
        selection.TypeText("AFTER_DISPLAY_BODY 大字号正文前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText(@"\(f_x:V\to\mathbb{R},\ f_x(y):=\ip{x}{y}\)");
        selection.Font.Size = 16;
        selection.TypeText(" 大字号正文后。");
        selection.TypeParagraph();

        selection.Font.Size = 12;
        selection.TypeText("无线信号前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText(@"$\left( \text{约}1.4\times 10^{-5}eV \right)$");
        selection.Font.Size = 12;
        selection.TypeText(" 无线信号后。");
        selection.TypeParagraph();
    }

    private static string WaitForLatexRedraw(string logPath, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var last = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(25);
            try
            {
                if (!File.Exists(logPath)) continue;
                last = File.ReadAllText(logPath, Encoding.UTF8);
                if (last.IndexOf("redraw-failed", StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        "Word LaTeX redraw failed.\n" + last);
                if (last.IndexOf("redraw-complete", StringComparison.Ordinal) >= 0)
                    return last;
            }
            catch (IOException)
            {
                // The add-in may be appending the current timing line.
            }
        }
        throw new TimeoutException(
            $"Word LaTeX redraw did not complete within {timeout.TotalSeconds:F0}s. Last log:\n{last}");
    }

    private static void AssertLatexRedrawDocument(
        Word.Document document,
        string objectMode)
    {
        Word.Range? content = null;
        Word.OMaths? maths = null;
        Word.InlineShapes? shapes = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            foreach (var required in new[]
                     {
                         "普通正文🙂前",
                         "普通正文后",
                         "大字号正文前",
                         "大字号正文后",
                         "无线信号前",
                         "无线信号后",
                     })
            {
                if (text.IndexOf(required, StringComparison.Ordinal) < 0)
                    throw new InvalidDataException(
                        $"Word LaTeX redraw lost surrounding prose: {required}");
            }
            foreach (var forbidden in new[] { "$UVI>2$", @"\(", @"\)", @"\[", @"\]", @"\ip" })
            {
                if (text.IndexOf(forbidden, StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        $"Word LaTeX redraw left source LaTeX text in the document: {forbidden}");
            }

            if (objectMode == FormulaOleContract.WordOmmlMode)
            {
                maths = document.OMaths;
                if (maths.Count != 4)
                    throw new InvalidDataException(
                        $"OMML redraw created {maths.Count} equations instead of 4.");
                AssertOmmlFontSize(maths[1], 10.5, 1);
                AssertOmmlFontSize(maths[2], 10.5, 2);
                AssertOmmlFontSize(maths[3], 16, 3);
                AssertOmmlFontSize(maths[4], 12, 4);
                Word.OMath? displayMath = null;
                Word.Range? displayRange = null;
                try
                {
                    displayMath = maths[2];
                    displayRange = displayMath.Range;
                    AssertDisplayFormulaFollowedImmediatelyByText(
                        document,
                        displayRange,
                        "AFTER_DISPLAY_BODY");
                }
                finally
                {
                    Release(displayRange);
                    Release(displayMath);
                }
            }
            else
            {
                shapes = document.InlineShapes;
                var native = new List<(Word.InlineShape Shape, FormulaMetadata Metadata)>();
                for (var index = 1; index <= shapes.Count; index++)
                {
                    var shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null && WordFormulaMetadataReader.IsNativeOle(shape))
                        native.Add((shape, metadata));
                    else
                        Release(shape);
                }
                try
                {
                    if (native.Count != 4)
                        throw new InvalidDataException(
                            $"OLE redraw created {native.Count} VisualTeX objects instead of 4.");
                    AssertOleFontSize(native, "UVI>2", 10.5, "OLE UVI formula font size");
                    AssertOleFontSize(
                        native,
                        @"f_x:V\to\mathbb{R},\ f_x(y):=\ip{x}{y}",
                        16,
                        "OLE inner-product formula font size");
                    AssertOleFontSize(native, "E=mc^2", 10.5, "OLE display formula font size");
                    AssertOleFontSize(
                        native,
                        @"\left( \text{约}1.4\times 10^{-5}eV \right)",
                        12,
                        "OLE wireless formula font size");
                    var display = native.Single(item =>
                        string.Equals(item.Metadata.Latex, "E=mc^2", StringComparison.Ordinal));
                    Word.Range? displayRange = null;
                    try
                    {
                        displayRange = display.Shape.Range;
                        AssertDisplayFormulaFollowedImmediatelyByText(
                            document,
                            displayRange,
                            "AFTER_DISPLAY_BODY");
                    }
                    finally { Release(displayRange); }
                }
                finally
                {
                    foreach (var item in native) Release(item.Shape);
                }
            }
        }
        finally
        {
            Release(shapes);
            Release(maths);
            Release(content);
        }
    }

    private static void AssertDisplayFormulaFollowedImmediatelyByText(
        Word.Document document,
        Word.Range formulaRange,
        string expectedText)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? nextAnchor = null;
        Word.Paragraphs? nextParagraphs = null;
        Word.Paragraph? nextParagraph = null;
        Word.Range? nextRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Display formula paragraph was not found.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            nextAnchor = document.Range(paragraphRange.End, paragraphRange.End);
            nextParagraphs = nextAnchor.Paragraphs;
            if (nextParagraphs.Count == 0)
                throw new InvalidDataException("Paragraph after the display formula was not found.");
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range;
            var nextText = nextRange.Text ?? string.Empty;
            if (nextText.IndexOf(expectedText, StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "Display formula introduced an empty paragraph before the following prose. "
                    + $"Expected next paragraph to contain '{expectedText}', actual='{nextText.Replace("\r", "<CR>").Replace("\v", "<BR>")}'.");
        }
        finally
        {
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextAnchor);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertOleFontSize(
        IReadOnlyList<(Word.InlineShape Shape, FormulaMetadata Metadata)> formulas,
        string latex,
        double expected,
        string label)
    {
        var matches = formulas
            .Where(item => string.Equals(item.Metadata.Latex, latex, StringComparison.Ordinal))
            .ToArray();
        if (matches.Length != 1)
            throw new InvalidDataException(
                $"{label} matched {matches.Length} OLE formulas for LaTeX: {latex}");
        AssertNear(matches[0].Metadata.FontSizePt ?? 0, expected, label);
    }

    private static void AssertOmmlFontSize(Word.OMath math, double expected, int index)
    {
        Word.Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = math.Range;
            font = range.Font;
            AssertNear(font.Size, expected, $"OMML formula {index} font size");
        }
        finally
        {
            Release(font);
            Release(range);
            Release(math);
        }
    }

    private static void AssertNear(double actual, double expected, string label)
    {
        if (double.IsNaN(actual)
            || double.IsInfinity(actual)
            || Math.Abs(actual - expected) > 0.6)
            throw new InvalidDataException(
                $"{label} was {actual.ToString("0.##", CultureInfo.InvariantCulture)} pt; expected {expected:0.##} pt.");
    }

    private static void AssertLatexRedrawPerformance(string log, string modeName)
    {
        var timings = Regex.Matches(log, @"\brender index=\d+ elapsedMs=(?<ms>\d+)")
            .Cast<Match>()
            .Select(match => long.Parse(match.Groups["ms"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        if (timings.Length != 4)
            throw new InvalidDataException(
                $"{modeName} redraw logged {timings.Length} render timings instead of 4.\n{log}");
        var maximum = timings.Max();
        Console.WriteLine(
            $"[Word LaTeX redraw] {modeName} render timings: {string.Join(", ", timings)} ms; max={maximum} ms");
        if (maximum > LatexRedrawPerformanceLimitMilliseconds)
            throw new InvalidDataException(
                $"{modeName} formula redraw exceeded the {LatexRedrawPerformanceLimitMilliseconds} ms target: max={maximum} ms.\n{log}");
    }

    private static void TryDeleteAcceptanceFile(string path)
    {
        try { File.Delete(path); } catch { }
    }
}
