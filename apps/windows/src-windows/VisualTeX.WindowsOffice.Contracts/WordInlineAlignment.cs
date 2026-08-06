using System;

namespace VisualTeX.WindowsOffice.Contracts;

public static class WordInlineAlignment
{
    public const float LegacyDescentRatio = 0.25f;
    public const int OpticalBaselineLiftPoints = 1;

    public static int CalculateFontPositionWithLegacyFallback(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        if (HasValidExportedBaseline(exportedHeight, exportedBaseline))
            return CalculateFontPosition(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline,
                targetSemanticFontSizePoints);

        var sourceSize = FormulaFontSize.Normalize(sourceSemanticFontSizePoints);
        var targetSize = FormulaFontSize.Normalize(targetSemanticFontSizePoints);
        if (existingFontPosition.HasValue
            && IsFinite(existingFontPosition.Value)
            && Math.Abs(existingFontPosition.Value) <= 256f
            && Math.Abs(existingFontPosition.Value) >= 0.01f)
        {
            // Word stores Font.Position as whole points. Remove half a point of
            // quantisation before scaling so an old -4 pt position at 14 pt maps
            // to -11 pt at 42 pt, rather than magnifying the original rounding
            // error to -12 pt. The same rule maps -11 pt back to -4 pt.
            var sign = Math.Sign(existingFontPosition.Value);
            var dequantizedMagnitude = Math.Max(
                0,
                Math.Abs(existingFontPosition.Value) - 0.5f);
            return sign * Math.Max(
                0,
                (int)Math.Round(
                    dequantizedMagnitude * targetSize / sourceSize,
                    MidpointRounding.AwayFromZero));
        }

        if (!(actualHeightPoints > 0) || !IsFinite(actualHeightPoints))
            return 0;
        return -Math.Max(
            0,
            (int)Math.Round(
                actualHeightPoints * LegacyDescentRatio,
                MidpointRounding.AwayFromZero));
    }

    public static int CalculateFontPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline) =>
        CalculateFontPosition(
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            semanticFontSizePoints: null);

    private static int CalculateFontPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        double? semanticFontSizePoints)
    {
        if (!(actualHeightPoints > 0)
            || !IsFinite(actualHeightPoints)
            || !HasValidExportedBaseline(exportedHeight, exportedBaseline))
            return 0;

        var baseline = exportedBaseline.GetValueOrDefault();
        var descentRatio = (exportedHeight - baseline) / exportedHeight;
        var downwardShiftPoints = actualHeightPoints * descentRatio;
        if (!(downwardShiftPoints > 0) || float.IsInfinity(downwardShiftPoints))
            return 0;

        // Word positions inline objects in whole points. Rounding the complete
        // MathJax descent places OLE previews about one point below native OMML
        // on ordinary text lines. Keep the geometric descent, but apply a small
        // optical lift so the visible glyph baseline matches Word's math zone.
        var roundedDescent = Math.Max(
            0,
            (int)Math.Round(downwardShiftPoints, MidpointRounding.AwayFromZero));
        var largeFontLift = CalculateLargeFontOpticalLift(semanticFontSizePoints);
        return -Math.Max(
            0,
            roundedDescent - OpticalBaselineLiftPoints - largeFontLift);
    }

    private static int CalculateLargeFontOpticalLift(double? semanticFontSizePoints)
    {
        if (!semanticFontSizePoints.HasValue
            || double.IsNaN(semanticFontSizePoints.Value)
            || double.IsInfinity(semanticFontSizePoints.Value))
            return 0;
        var fontSize = FormulaFontSize.Normalize(semanticFontSizePoints.Value);
        // Office 2021's inline OLE box starts to sit visibly below adjacent
        // Times New Roman text only at large display sizes. Keep 12/18/24 pt
        // untouched; add one point around 33–41 pt and two points at 42 pt+.
        return Math.Max(
            0,
            Math.Min(2, (int)Math.Floor((fontSize - 24.0) / 9.0)));
    }

    private static bool HasValidExportedBaseline(
        float exportedHeight,
        float? exportedBaseline) =>
        exportedHeight > 0
        && IsFinite(exportedHeight)
        && exportedBaseline.HasValue
        && IsFinite(exportedBaseline.Value)
        && exportedBaseline.Value >= 0
        && exportedBaseline.Value < exportedHeight;

    private static bool IsFinite(float value) =>
        !float.IsNaN(value) && !float.IsInfinity(value);
}
