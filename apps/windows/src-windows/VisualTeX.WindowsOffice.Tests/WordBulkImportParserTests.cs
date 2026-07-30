using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordBulkImportParserTests
{
    [Fact]
    public void MarkdownProducesNativeTextAndIndependentInlineAndDisplayFormulas()
    {
        const string source = """
            # 示例标题

            这是 **粗体**、*斜体* 与行内公式 $E=mc^2$。

            $$
            \int_0^1 x^2\,\mathrm{d}x=\frac13
            $$

            - 第一项含公式 $a+b$
            - 第二项
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(5, document.Blocks.Count);
        Assert.Equal(3, document.FormulaCount);
        Assert.Equal(2, document.InlineFormulaCount);
        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Contains(document.Blocks[1].Runs, run => !run.IsFormula && run.Bold && run.Text == "粗体");
        Assert.Contains(document.Blocks[1].Runs, run => !run.IsFormula && run.Italic && run.Text == "斜体");
        Assert.Contains(document.Blocks[1].Runs, run => run.IsFormula && run.Latex == "E=mc^2");
        Assert.Equal(WordBulkBlockKind.DisplayFormula, document.Blocks[2].Kind);
        Assert.Equal("\\int_0^1 x^2\\,\\mathrm{d}x=\\frac13", document.Blocks[2].Runs.Single().Latex);
        Assert.All(document.Blocks.Skip(3), block => Assert.Equal(WordBulkBlockKind.Bullet, block.Kind));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void LatexDocumentStripsPreambleAndConvertsSectionsListsAndMath()
    {
        const string source = """
            \documentclass{article}
            \usepackage{amsmath}
            \begin{document}
            \section{理论背景}
            普通文字与 \textbf{重点}，以及 \(\alpha+\beta\)。
            \begin{equation}
            F=ma
            \end{equation}
            \begin{enumerate}
            \item 第一项 $x_1$
            \item 第二项
            \end{enumerate}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Ole);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(WordBulkFormulaObjectMode.Ole, document.FormulaObjectMode);
        Assert.Equal(3, document.FormulaCount);
        Assert.Equal(2, document.InlineFormulaCount);
        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Equal(1, document.Blocks[0].Level);
        Assert.Contains(document.Blocks[1].Runs, run => run.Bold && run.Text == "重点");
        Assert.Contains(document.Blocks[1].Runs, run => run.IsFormula && run.Latex == "\\alpha+\\beta");
        Assert.Equal("F=ma", document.Blocks[2].Runs.Single().Latex);
        Assert.All(document.Blocks.Skip(3), block => Assert.Equal(WordBulkBlockKind.Numbered, block.Kind));
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run => run.Text.Contains("documentclass"));
    }

    [Fact]
    public void LatexAmsDisplayEnvironmentsRemainRenderableAndEditable()
    {
        const string source = """
            \begin{document}
            \begin{align}
            a+b &= c \\
            d &= e+f
            \end{align}
            \begin{gather*}
            x=1 \\
            y=2
            \end{gather*}
            \begin{multline}
            p+q+r \\
            = s+t
            \end{multline}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(3, document.DisplayFormulaCount);
        var formulas = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex.Replace("\r", string.Empty))
            .ToArray();
        Assert.Equal(
            """
            \begin{aligned}a+b &= c \\
            d &= e+f\end{aligned}
            """.Replace("\r", string.Empty),
            formulas[0]);
        Assert.Equal(
            """
            \begin{gathered}x=1 \\
            y=2\end{gathered}
            """.Replace("\r", string.Empty),
            formulas[1]);
        Assert.Equal(
            """
            \begin{gathered}p+q+r \\
            = s+t\end{gathered}
            """.Replace("\r", string.Empty),
            formulas[2]);
    }

    [Fact]
    public void MarkdownAndLatexPreserveExtendedIntegralCommandsVerbatim()
    {
        const string markdown = """
            闭合曲面积分 $\oiint_{\Sigma} a\,\mathrm{d}S$。

            $$
            \oiiint_V f\,\mathrm{d}V+\intclockwise_C g\,\mathrm{d}s
            $$
            """;
        const string latex = """
            \documentclass{article}
            \begin{document}
            行内 \(\varointclockwise_C f\,\mathrm{d}s\)。
            \[
            \ointctrclockwise_C g\,\mathrm{d}s+\intctrclockwise_C h\,\mathrm{d}s
            \]
            \end{document}
            """;

        var markdownDocument = WordBulkImportParser.Parse(
            markdown,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Ole);
        var latexDocument = WordBulkImportParser.Parse(
            latex,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        var markdownFormulas = markdownDocument.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();
        var latexFormulas = latexDocument.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();

        Assert.Equal(2, markdownFormulas.Length);
        Assert.Contains("\\oiint_{\\Sigma}", markdownFormulas[0], StringComparison.Ordinal);
        Assert.Contains("\\oiiint_V", markdownFormulas[1], StringComparison.Ordinal);
        Assert.Contains("\\intclockwise_C", markdownFormulas[1], StringComparison.Ordinal);
        Assert.Equal(2, latexFormulas.Length);
        Assert.Contains("\\varointclockwise_C", latexFormulas[0], StringComparison.Ordinal);
        Assert.Contains("\\ointctrclockwise_C", latexFormulas[1], StringComparison.Ordinal);
        Assert.Contains("\\intctrclockwise_C", latexFormulas[1], StringComparison.Ordinal);
        Assert.Empty(markdownDocument.Warnings);
        Assert.Empty(latexDocument.Warnings);
    }

    [Fact]
    public void EscapedDollarRemainsNativeText()
    {
        var document = WordBulkImportParser.Parse(
            "价格写作 \\$5，而公式是 $x+1$。",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(1, document.FormulaCount);
        var runs = document.Blocks.Single().Runs;
        Assert.Contains(runs, run => !run.IsFormula && run.Text.Contains("$5"));
        Assert.Contains(runs, run => run.IsFormula && run.Latex == "x+1");
    }

    [Fact]
    public void UnclosedDisplayFormulaIsImportedWithWarning()
    {
        var document = WordBulkImportParser.Parse(
            "正文\n\n$$\na+b",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Single(document.Warnings);
        Assert.Contains("缺少结束标记", document.Warnings[0]);
    }

    [Fact]
    public void LatexNestedListsQuotesAndLiteralCodeRemainStructured()
    {
        const string source = """
            \documentclass{book}
            \begin{document}
            \chapter{结构测试}
            \begin{quote}
            第一行包含 \textbf{重点}。
            第二行包含行内公式 $x+y$。
            \end{quote}
            \begin{itemize}
            \item 外层项目
            \begin{enumerate}
            \item 内层编号项目 $n=1$
            \end{enumerate}
            \end{itemize}
            \begin{verbatim}
            value_1 = 20 % 这里不是注释
            \end{verbatim}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Equal(1, document.Blocks[0].Level);
        var quote = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Quote);
        Assert.Contains(quote.Runs, run => !run.IsFormula && run.Bold && run.Text == "重点");
        Assert.Contains(quote.Runs, run => run.IsFormula && run.Latex == "x+y");
        var lists = document.Blocks
            .Where(block => block.Kind is WordBulkBlockKind.Bullet or WordBulkBlockKind.Numbered)
            .ToArray();
        Assert.Equal(2, lists.Length);
        Assert.Equal(WordBulkBlockKind.Bullet, lists[0].Kind);
        Assert.Equal(0, lists[0].Level);
        Assert.Equal(WordBulkBlockKind.Numbered, lists[1].Kind);
        Assert.Equal(1, lists[1].Level);
        var code = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Code);
        Assert.True(code.Runs.Single().Code);
        Assert.Contains("value_1 = 20 % 这里不是注释", code.Runs.Single().Text);
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void MarkdownMultilineQuoteAndIndentedListsPreserveLevels()
    {
        const string source = """
            > 第一行引用
            > 第二行含公式 $q=1$

            - 一级项目
              - 二级项目
            1. 一级编号
                1. 三级编号
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Ole);

        var quote = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Quote);
        Assert.Contains(quote.Runs, run => !run.IsFormula && run.Text.Contains("第一行引用 第二行含公式"));
        Assert.Contains(quote.Runs, run => run.IsFormula && run.Latex == "q=1");
        var lists = document.Blocks.Skip(1).ToArray();
        Assert.Equal(
            new[] { 0, 1, 0, 2 },
            lists.Select(block => block.Level).ToArray());
        Assert.Equal(
            new[]
            {
                WordBulkBlockKind.Bullet,
                WordBulkBlockKind.Bullet,
                WordBulkBlockKind.Numbered,
                WordBulkBlockKind.Numbered,
            },
            lists.Select(block => block.Kind).ToArray());
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void EmbeddedAndAdjacentDisplayDelimitersAreSplitFromNativeParagraphs()
    {
        const string source = """
            对于高频相应，也可以对时间常数做一个近似计算： $\tau_H=\sum_{i=1}^{n}\tau_i$

            对于共射极放大电路： \[ R_1=\left(R_{si}\parallel R_b+r_{bb'}\right)\parallel r_{be}\]\[\tau_1=R_1C_{be}\]\[\tau_2=R_2C_{bc}=\left[\left(1+g_mR_L'\right)R_1+R_L'\right]C_{bc}\]
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(4, document.FormulaCount);
        Assert.Equal(1, document.InlineFormulaCount);
        Assert.Equal(3, document.DisplayFormulaCount);
        Assert.Equal(5, document.Blocks.Count);
        Assert.Equal(WordBulkBlockKind.Paragraph, document.Blocks[0].Kind);
        Assert.Equal(WordBulkBlockKind.Paragraph, document.Blocks[1].Kind);
        Assert.Equal("对于共射极放大电路：", string.Concat(
            document.Blocks[1].Runs.Where(run => !run.IsFormula).Select(run => run.Text)).Trim());
        var display = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex)
            .ToArray();
        Assert.Equal("R_1=\\left(R_{si}\\parallel R_b+r_{bb'}\\right)\\parallel r_{be}", display[0]);
        Assert.Equal("\\tau_1=R_1C_{be}", display[1]);
        Assert.Equal("\\tau_2=R_2C_{bc}=\\left[\\left(1+g_mR_L'\\right)R_1+R_L'\\right]C_{bc}", display[2]);
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run =>
            !run.IsFormula && (run.Text.Contains("\\[") || run.Text.Contains("\\]")));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void EmbeddedDisplayFormulaCanSpanLinesAndKeepTrailingText()
    {
        const string source = """
            前文 \[
            \begin{aligned}
            a+b&=c\\
            d&=e
            \end{aligned}
            \] 后文仍然是正文。
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Ole);

        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(3, document.Blocks.Count);
        Assert.Equal("前文", document.Blocks[0].Runs.Single().Text.Trim());
        Assert.Contains("\\begin{aligned}", document.Blocks[1].Runs.Single().Latex);
        Assert.Equal("后文仍然是正文。", document.Blocks[2].Runs.Single().Text.Trim());
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void DisplaySourceFormattingDoesNotCreateInvalidAlignedRows()
    {
        const string source = """
            对于高频相应，也可以对时间常数做一个近似计算：$$\tau _H=\sum_{i=1}^n{\tau _i}$$
            $$f_{_{\mathrm{H}}}=\frac{1}{2\pi \tau _{_{\mathrm{H}}}}$$
            对于共射级放大电路：\[
            R_1
            =
            \left(R_{si}\parallel R_b+r_{bb'}\right)
            \parallel r_{b'e}
            \]\[
            \tau_1
            =
            R_1 C_{b'e}
            \]\[
            \tau_2
            =
            R_2 C_{b'c}
            =
            \left[
            \left(1+g_mR_L'\right)R_1
            +
            R_L'
            \right]C_{b'c}
            \]
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(5, document.FormulaCount);
        Assert.Equal(0, document.InlineFormulaCount);
        Assert.Equal(5, document.DisplayFormulaCount);
        var formulas = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex)
            .ToArray();
        Assert.All(formulas, formula =>
        {
            Assert.DoesNotContain('\r', formula);
            Assert.DoesNotContain('\n', formula);
        });
        Assert.Equal(
            "\\tau_2 = R_2 C_{b'c} = \\left[ \\left(1+g_mR_L'\\right)R_1 + R_L' \\right]C_{b'c}",
            formulas[4]);
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run =>
            !run.IsFormula && (run.Text.Contains("$$") || run.Text.Contains("\\[") || run.Text.Contains("\\]")));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void RejectsEmptyOrExcessivelyLargeInput()
    {
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.Parse(
            "   ",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml));
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.Parse(
            new string('a', 5_000_001),
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml));
    }
}
