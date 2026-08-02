using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordInlineAlignmentTests
{
    [Fact]
    public void AlignsExportedFormulaBaselineWithWordTextBaseline()
    {
        Assert.Equal(-3, WordInlineAlignment.CalculateFontPosition(15, 20, 15));
    }

    [Fact]
    public void RawGeometryAlignmentScalesWithTheObjectHeight()
    {
        Assert.Equal(-7, WordInlineAlignment.CalculateFontPosition(30, 40, 30));
    }

    [Fact]
    public void CompleteMetadataKeepsTheStructureAwareScaledBaseline()
    {
        Assert.Equal(
            -10,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                18.9867f,
                14.16f,
                existingFontPosition: -4,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
    }

    [Fact]
    public void LegacyFormulaScalesItsExistingCorrectBaseline()
    {
        Assert.Equal(
            -11,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: -4,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
    }

    [Fact]
    public void LegacyBaselineScalingRoundTripsAcrossCommonSizes()
    {
        var enlarged = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
            43,
            exportedHeight: 0,
            exportedBaseline: null,
            existingFontPosition: -4,
            sourceSemanticFontSizePoints: 14,
            targetSemanticFontSizePoints: 42);
        Assert.Equal(-11, enlarged);
        Assert.Equal(
            -4,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                14.333f,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: enlarged,
                sourceSemanticFontSizePoints: 42,
                targetSemanticFontSizePoints: 14));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0d)]
    [InlineData(9999999d)]
    public void LegacyFormulaWithoutUsablePositionUsesDescentFallback(
        double? existingPosition)
    {
        Assert.Equal(
            -11,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: existingPosition is null
                    ? null
                    : (float)existingPosition.Value,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
    }

    [Theory]
    [InlineData(20d, 20d, null)]
    [InlineData(20d, 0d, 0d)]
    [InlineData(20d, 20d, -1d)]
    [InlineData(20d, 20d, 21d)]
    [InlineData(20d, 20d, 20d)]
    public void InvalidOrBottomEdgeBaselinesDoNotMoveTheFormula(
        double actualHeight,
        double exportedHeight,
        double? baseline)
    {
        Assert.Equal(0, WordInlineAlignment.CalculateFontPosition(
            (float)actualHeight,
            (float)exportedHeight,
            baseline is null ? null : (float)baseline.Value));
    }
}
