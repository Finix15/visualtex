using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class PreparedWordBulkFormula
{
    internal WordBulkRun Run { get; set; } = new();
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? EmfPath { get; set; }
}

internal sealed class RenderedWordBulkFormulaTemplate
{
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? SvgPath { get; set; }
    internal string? EmfPath { get; set; }
}

internal sealed class WordBulkInsertResult
{
    internal int BlockCount { get; set; }
    internal int FormulaCount { get; set; }
    internal List<string> FormulaIds { get; set; } = new();
}
