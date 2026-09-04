namespace VisualTeX.WindowsOffice.VstoShared;

internal static class WordMathTypeLayoutValidation
{
    internal static bool IsStandaloneDisplayCandidate(
        string? prefix,
        string? suffix,
        int inlineShapeCount,
        int existingMathCount,
        out string reason)
    {
        if (inlineShapeCount != 1)
        {
            reason = "inline-object-count";
            return false;
        }
        if (existingMathCount != 0)
        {
            reason = "inline-existing-omath";
            return false;
        }

        var surrounding = (prefix ?? string.Empty) + StripParagraphTerminators(suffix);
        if (surrounding.Length == 0)
        {
            reason = "display-ole-only";
            return true;
        }
        if (surrounding.Any(character => !char.IsWhiteSpace(character)))
            reason = "inline-surrounding-content";
        else if (surrounding.IndexOf('\t') >= 0)
            reason = "inline-tab-layout";
        else if (surrounding.IndexOf('\v') >= 0 || surrounding.IndexOf('\n') >= 0)
            reason = "inline-manual-break";
        else if (surrounding.Any(char.IsWhiteSpace))
            reason = "inline-whitespace-layout";
        else
            reason = "inline-whitespace-layout";
        return false;
    }

    internal static bool SurroundingsChanged(
        string expectedPrefix,
        string expectedSuffix,
        string currentPrefix,
        string currentSuffix)
    {
        return !string.Equals(currentPrefix, expectedPrefix, StringComparison.Ordinal)
            || !string.Equals(currentSuffix, expectedSuffix, StringComparison.Ordinal);
    }

    private static string StripParagraphTerminators(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var end = value!.Length;
        while (end > 0 && value[end - 1] is '\r' or '\a') end--;
        return end == value.Length ? value : value.Substring(0, end);
    }
}
