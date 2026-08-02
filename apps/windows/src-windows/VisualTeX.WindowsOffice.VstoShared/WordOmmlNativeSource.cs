using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlNativeSource
{
    internal static FormulaMetadata RefreshForVisualTeX(
        Document document,
        Bookmark bookmark,
        FormulaMetadata stored)
    {
        Range? equationRange = null;
        try
        {
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var wordOpenXml = ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                stored.FormulaId);
            var fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(wordOpenXml);
            if (string.Equals(
                    stored.NativeOmmlFingerprint,
                    fingerprint,
                    StringComparison.OrdinalIgnoreCase))
                return stored;

            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                wordOpenXml,
                display: string.Equals(
                    stored.DisplayMode,
                    "block",
                    StringComparison.Ordinal));
            var latex = MathMlToLatexConverter.Convert(mathMl);
            if (string.IsNullOrWhiteSpace(latex))
                throw new InvalidDataException(
                    "The Word-native OMML equation could not be converted back to editable LaTeX.");

            var refreshed = Clone(stored);
            var lineId = refreshed.Lines.FirstOrDefault()?.Id;
            if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString();
            refreshed.Latex = latex;
            refreshed.Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            };
            refreshed.CodeFormat = "raw";
            refreshed.NativeOmmlFingerprint = fingerprint;
            refreshed.Validate();
            return refreshed;
        }
        finally
        {
            Release(equationRange);
        }
    }

    internal static void StampFingerprint(FormulaMetadata metadata, Range equationRange)
    {
        Document? document = null;
        try
        {
            document = equationRange.Document;
            metadata.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                ReadCompleteEquationWordOpenXml(
                    document,
                    equationRange,
                    metadata.FormulaId));
        }
        finally { Release(document); }
    }

    private static string ReadCompleteEquationWordOpenXml(
        Document document,
        Range equationRange,
        string formulaId)
    {
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? boundaryBookmark = null;
        Range? boundaryRange = null;
        try
        {
            content = document.Content;
            var probeEnd = equationRange.End;
            if (Guid.TryParse(formulaId, out var parsed))
            {
                bookmarks = document.Bookmarks;
                var boundaryName = "VTBL_" + parsed.ToString("N");
                if (bookmarks.Exists(boundaryName))
                {
                    boundaryBookmark = bookmarks[boundaryName];
                    boundaryRange = boundaryBookmark.Range;
                    probeEnd = Math.Max(probeEnd, boundaryRange.End);
                }
            }

            // A field boundary can span dozens of Word structure characters.
            // Include the complete field, not merely its first marker; otherwise
            // Word may serialize only the leading fragment of a compound OMath.
            object start = equationRange.Start;
            object end = Math.Min(content.End, Math.Max(probeEnd, equationRange.End));
            probe = document.Range(ref start, ref end);
            var xml = probe.WordOpenXML;
            WordOmmlConverter.ExtractSingleOMath(xml);
            return xml;
        }
        finally
        {
            Release(boundaryRange);
            Release(boundaryBookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    private static FormulaMetadata Clone(FormulaMetadata metadata)
    {
        var clone = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        return clone
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
    }

    private static void Release(object? value)
    {
        if (value is null || !System.Runtime.InteropServices.Marshal.IsComObject(value)) return;
        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
    }
}
