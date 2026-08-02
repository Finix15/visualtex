using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class WordDoubleClickRouting
{
    internal static bool ShouldOpenVisualTeX(OfficeSelection? selection)
    {
        if (selection?.Metadata is null
            || string.IsNullOrWhiteSpace(selection.FormulaId))
            return false;

        // Every VisualTeX-managed formula reopens the same editor, including
        // native Word OMML. Ordinary Word equations have no VisualTeX metadata
        // and therefore continue to use Word's built-in double-click editor.
        return true;
    }
}
