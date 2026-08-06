using System.Text;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunOfficeOleViewBoxProbe(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
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

        File.WriteAllText(
            Path.Combine(artifactRoot, "office-ole-viewbox-source.tex"),
            source,
            new UTF8Encoding(false));

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Ole);
        var formulas = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .GroupBy(
                run => string.Join("\u001F", run.DisplayMode, run.Latex, run.EquationTag ?? string.Empty),
                StringComparer.Ordinal)
            .Select(group => group.First())
            .ToArray();
        Console.WriteLine(
            $"OLE viewBox probe parsed {parsed.FormulaCount} formulas, {formulas.Length} unique renderings.");

        var failures = new List<string>();
        for (var index = 0; index < formulas.Length; index++)
        {
            var run = formulas[index];
            string? sessionId = null;
            string? svgPath = null;
            string? emfPath = null;
            try
            {
                var formulaId = Guid.NewGuid().ToString("D");
                var line = new FormulaLine
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = run.Latex,
                };
                var session = client.CreateSessionAsync(
                        new CreateVstoSessionRequest
                        {
                            Mode = "create",
                            Host = "word",
                            Title = "OLE viewBox diagnostic",
                            Lines = new List<FormulaLine> { line },
                            ActiveLineId = line.Id,
                            CodeFormat = "latex",
                            DisplayMode = run.DisplayMode,
                            ObjectMode = FormulaOleContract.NativeOleMode,
                            Numbered = false,
                            FontSizePt = 11d,
                            FormulaId = formulaId,
                            OriginalMetadata = new FormulaMetadata
                            {
                                FormulaId = formulaId,
                                Title = "OLE viewBox diagnostic",
                                Latex = run.Latex,
                                Lines = new List<FormulaLine> { line },
                                CodeFormat = "latex",
                                DisplayMode = run.DisplayMode,
                                Numbered = false,
                                EquationTag = run.EquationTag,
                                FontSizePt = 11d,
                                RenderFontSizePt = 11d,
                                CreatedWithVersion = "1.2.4",
                                UpdatedWithVersion = "1.2.4",
                                CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
                                UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
                            },
                            AutoCommitOnClose = false,
                        },
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                sessionId = session.Id;
                client.OpenConverterAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                session = client.WaitForCommitAsync(
                        sessionId,
                        TimeSpan.FromSeconds(45),
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (string.Equals(session.Status, "failed", StringComparison.Ordinal)
                    || session.ExportResult is null)
                    throw new InvalidDataException(
                        session.Error ?? "The production converter returned no export result.");

                svgPath = client.MaterializeSvg(session);
                var artifactSvg = Path.Combine(
                    artifactRoot,
                    $"formula-{index + 1:00}.svg");
                File.Copy(svgPath, artifactSvg, overwrite: true);
                var viewBoxes = ReadSvgViewBoxes(artifactSvg);
                emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                    svgPath,
                    session.ExportResult.Width,
                    session.ExportResult.Height);
                File.Copy(
                    emfPath,
                    Path.Combine(artifactRoot, $"formula-{index + 1:00}.emf"),
                    overwrite: true);
                Console.WriteLine(
                    $"[PASS] #{index + 1:00} {run.DisplayMode} "
                    + $"{session.ExportResult.Width:0.###}x{session.ExportResult.Height:0.###} "
                    + $"viewBoxes=[{string.Join(" | ", viewBoxes)}] latex={run.Latex}");
            }
            catch (Exception error)
            {
                var viewBoxes = svgPath is not null && File.Exists(svgPath)
                    ? ReadSvgViewBoxes(svgPath)
                    : Array.Empty<string>();
                var detail =
                    $"#{index + 1:00} display={run.DisplayMode} tag={run.EquationTag ?? "<none>"} "
                    + $"latex={run.Latex}\n"
                    + $"viewBoxes=[{string.Join(" | ", viewBoxes)}]\n"
                    + error;
                failures.Add(detail);
                Console.WriteLine("[FAIL] " + detail);
            }
            finally
            {
                if (emfPath is not null)
                {
                    try { File.Delete(emfPath); } catch { }
                }
                if (svgPath is not null)
                {
                    try { File.Delete(svgPath); } catch { }
                }
                if (!string.IsNullOrWhiteSpace(sessionId))
                {
                    try
                    {
                        client.CompleteAsync(sessionId!, CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    catch { }
                    try
                    {
                        client.CloseEditorAsync(sessionId!, CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    catch { }
                }
            }
        }

        if (failures.Count > 0)
            throw new InvalidDataException(
                "OLE viewBox probe failed:\n" + string.Join("\n\n", failures));
        Console.WriteLine(
            $"OLE viewBox probe passed: {formulas.Length} unique formula renderings.");
    }

    private static string[] ReadSvgViewBoxes(string path)
    {
        try
        {
            var document = XDocument.Load(path, LoadOptions.None);
            return (document.Root?.DescendantsAndSelf() ?? Enumerable.Empty<XElement>())
                .Where(element => string.Equals(
                    element.Name.LocalName,
                    "svg",
                    StringComparison.Ordinal))
                .Select(element => element.Attribute("viewBox")?.Value ?? "<missing>")
                .ToArray();
        }
        catch (Exception error)
        {
            return new[] { "<unable to parse: " + error.Message + ">" };
        }
    }
}
