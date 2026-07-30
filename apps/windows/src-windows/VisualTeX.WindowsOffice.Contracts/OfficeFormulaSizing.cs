using System;

namespace VisualTeX.WindowsOffice.Contracts;

public static class OfficeFormulaSizing
{
    private const float PointsPerPixel = 0.75f;
    private const float MinimumPoints = 12f;

    public static (float Width, float Height) NaturalSize(
        float renderWidth,
        float renderHeight)
    {
        return (
            Math.Max(MinimumPoints, Math.Max(1f, renderWidth) * PointsPerPixel),
            Math.Max(MinimumPoints, Math.Max(1f, renderHeight) * PointsPerPixel));
    }

    public static (float Width, float Height) EditedSize(
        float currentWidth,
        float currentHeight,
        double? originalRenderWidth,
        double? originalRenderHeight,
        float newRenderWidth,
        float newRenderHeight,
        float maximumWidth = float.PositiveInfinity,
        float maximumHeight = float.PositiveInfinity,
        double? originalFontSizePt = null,
        double? originalRenderFontSizePt = null)
    {
        var next = NaturalSize(newRenderWidth, newRenderHeight);
        var scale = 1f;
        if (originalRenderWidth is > 0 && originalRenderHeight is > 0
            && currentWidth > 0 && currentHeight > 0)
        {
            // Use the preview's real physical dimensions when recovering the
            // existing object scale. NaturalSize applies a 12 pt minimum box,
            // which is useful for creating/selecting tiny objects but is not a
            // valid font-scale reference. A 25x11 px inline formula is really
            // 18.75x8.25 pt; comparing its 8.5 pt Word height with a clamped
            // 12 pt height incorrectly shrinks a wider edited formula to 71%.
            var previousWidth = Math.Max(0.01f,
                (float)originalRenderWidth.Value * PointsPerPixel);
            var previousHeight = Math.Max(0.01f,
                (float)originalRenderHeight.Value * PointsPerPixel);
            var horizontalScale = currentWidth / previousWidth;
            var verticalScale = currentHeight / previousHeight;

            // Formula height is the visual font-size reference. Prefer it over
            // the geometric mean so picture→OLE conversion cannot become
            // shorter and wider when the old picture box was non-uniform or
            // came from a legacy raster export.
            if (IsPositiveFinite(verticalScale))
                scale = verticalScale;
            else if (IsPositiveFinite(horizontalScale))
                scale = horizontalScale;
        }
        else if (currentHeight > 0 && IsPositiveFinite(next.Height))
        {
            // Legacy picture metadata can lack natural render dimensions.
            // Preserve its physical height and recover the new width from the
            // replacement formula's natural aspect ratio.
            scale = currentHeight / next.Height;
        }
        else if (currentWidth > 0 && IsPositiveFinite(next.Width))
        {
            scale = currentWidth / next.Width;
        }
        if (originalFontSizePt is > 0 && originalRenderFontSizePt is > 0)
        {
            var semanticScale = FormulaFontSize.Normalize(originalFontSizePt)
                / FormulaFontSize.Normalize(originalRenderFontSizePt);
            if (IsPositiveFinite(semanticScale)) scale /= semanticScale;
        }
        if (!IsPositiveFinite(scale)) scale = 1f;
        scale = Math.Max(0.1f, Math.Min(10f, scale));

        var width = next.Width * scale;
        var height = next.Height * scale;
        var fitScale = Math.Min(
            1f,
            Math.Min(
                IsPositiveFinite(maximumWidth) ? maximumWidth / width : 1f,
                IsPositiveFinite(maximumHeight) ? maximumHeight / height : 1f));
        if (IsPositiveFinite(fitScale) && fitScale < 1f)
        {
            width *= fitScale;
            height *= fitScale;
        }
        return (Math.Max(1f, width), Math.Max(1f, height));
    }

    private static bool IsPositiveFinite(float value) =>
        value > 0 && !float.IsNaN(value) && !float.IsInfinity(value);
}
