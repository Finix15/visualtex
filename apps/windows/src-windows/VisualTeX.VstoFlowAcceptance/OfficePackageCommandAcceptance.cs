using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunOfficePackageCommandProbe(VisualTeXSessionClient client)
    {
        var cases = new (string Name, string Latex)[]
        {
            ("physics-qty-parentheses", "\\qty(\\frac{a}{b})"),
            ("physics-dv", "\\dv{f}{x}"),
            ("physics-pdv", "\\pdv{g}{y}"),
            ("physics-abs", "\\abs{x}"),
            ("physics-norm", "\\norm{\\bm v}"),
            ("siunitx-SI", "\\SI{3}{\\meter\\per\\second}"),
            ("siunitx-si", "\\si{\\kilogram}"),
            ("siunitx-unit", "\\unit{\\joule}"),
            ("siunitx-v3-qty", "\\qty{5}{\\tesla}"),
            ("bbm", "\\mathbbm{1}_{A}"),
            ("physics-derivative-orders", "\\dv{x}+\\dv[2]{f}{x}+\\pdv[3]{g}{y}+\\fdv{S}{\\phi}"),
            ("physics-vb", "\\vb{v}"),
            ("physics-va", "\\va{a}"),
            ("physics-vu", "\\vu{n}"),
            ("physics-pb", "\\pb{f}{g}"),
            ("physics-order", "\\order{x^2}"),
            ("physics-Tr", "\\Tr A"),
            ("physics-rank", "\\rank A"),
            ("physics-vectors-operators", "\\vb{v}+\\va{a}+\\vu{n}+\\pb{f}{g}+\\order{x^2}+\\Tr A+\\rank A"),
            ("physics-mqty", "\\mqty{a&b\\\\c&d}"),
            ("physics-pmqty", "\\pmqty{1&0\\\\0&1}"),
            ("physics-vmqty", "\\vmqty{x&y\\\\z&w}"),
            ("physics-matrix-quantities", "\\mqty{a&b\\\\c&d}+\\pmqty{1&0\\\\0&1}+\\vmqty{x&y\\\\z&w}"),
            ("siunitx-SI-options", "\\SI[round-mode=places]{3.14}{\\kilo\\meter\\per\\second}"),
            ("siunitx-qty-options", "\\qty[round-mode=figures]{5}{\\tesla}"),
            ("siunitx-qtyrange", "\\qtyrange{1}{10}{\\milli\\second}"),
            ("siunitx-ang", "\\ang{30}"),
            ("siunitx-options-ranges", "\\SI[round-mode=places]{3.14}{\\kilo\\meter\\per\\second}+\\qty[round-mode=figures]{5}{\\tesla}+\\qtyrange{1}{10}{\\milli\\second}+\\ang{30}"),
            ("combined", "\\qty(\\frac{a}{b})+\\dv{f}{x}+\\pdv{g}{y}+\\SI{3}{\\meter\\per\\second}+\\mathbbm{1}_{A}"),
        };

        var failures = new List<string>();
        foreach (var testCase in cases)
        {
            string? sessionId = null;
            try
            {
                var line = new FormulaLine
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = testCase.Latex,
                };
                var session = client.CreateSessionAsync(
                        new CreateVstoSessionRequest
                        {
                            Mode = "create",
                            Host = "word",
                            Title = "Office package compatibility probe",
                            Lines = new List<FormulaLine> { line },
                            ActiveLineId = line.Id,
                            CodeFormat = "latex",
                            DisplayMode = "block",
                            ObjectMode = FormulaOleContract.WordOmmlMode,
                            Numbered = false,
                            FontSizePt = 11d,
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
                    || session.ExportResult is null
                    || string.IsNullOrWhiteSpace(session.ExportResult.MathMl))
                {
                    failures.Add(
                        $"{testCase.Name}: {session.Error ?? "converter returned no MathML"}");
                    Console.WriteLine($"[FAIL] {testCase.Name}: {session.Error}");
                }
                else
                {
                    Console.WriteLine(
                        $"[PASS] {testCase.Name}: "
                        + $"{session.ExportResult.Width:0.###}x{session.ExportResult.Height:0.###}");
                }
            }
            catch (Exception error)
            {
                failures.Add($"{testCase.Name}: {error.Message}");
                Console.WriteLine($"[FAIL] {testCase.Name}: {error.Message}");
            }
            finally
            {
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
                "Office package command probe failed:\n" + string.Join("\n", failures));
        Console.WriteLine(
            $"Office package command probe passed: {cases.Length} production converter cases.");
    }
}
