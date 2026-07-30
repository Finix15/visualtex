using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class OfficeFormulaSizingTests
{
    [Fact]
    public void NaturalSizeUsesOfficePointScale()
    {
        var size = OfficeFormulaSizing.NaturalSize(160f, 32f);

        Assert.Equal(120f, size.Width, 3);
        Assert.Equal(24f, size.Height, 3);
    }

    [Fact]
    public void EditedSizeAdoptsNewNaturalAspectRatio()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 120f,
            currentHeight: 24f,
            originalRenderWidth: 160d,
            originalRenderHeight: 32d,
            newRenderWidth: 320f,
            newRenderHeight: 32f);

        Assert.Equal(240f, size.Width, 3);
        Assert.Equal(24f, size.Height, 3);
        Assert.Equal(10f, size.Width / size.Height, 3);
    }

    [Fact]
    public void EditedSizePreservesUniformUserScale()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 180f,
            currentHeight: 36f,
            originalRenderWidth: 160d,
            originalRenderHeight: 32d,
            newRenderWidth: 320f,
            newRenderHeight: 32f);

        Assert.Equal(360f, size.Width, 3);
        Assert.Equal(36f, size.Height, 3);
    }

    [Fact]
    public void EditedSizeDoesNotApplySemanticFontScaleTwice()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 450f,
            currentHeight: 112.5f,
            originalRenderWidth: 400d,
            originalRenderHeight: 100d,
            newRenderWidth: 600f,
            newRenderHeight: 100f,
            originalFontSizePt: 21d,
            originalRenderFontSizePt: 14d);

        Assert.Equal(450f, size.Width, 3);
        Assert.Equal(75f, size.Height, 3);
    }

    [Fact]
    public void EditedSizeFitsPowerPointBoundsWithoutDistortion()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 120f,
            currentHeight: 24f,
            originalRenderWidth: 160d,
            originalRenderHeight: 32d,
            newRenderWidth: 1600f,
            newRenderHeight: 160f,
            maximumWidth: 600f,
            maximumHeight: 400f);

        Assert.Equal(600f, size.Width, 3);
        Assert.Equal(60f, size.Height, 3);
        Assert.Equal(10f, size.Width / size.Height, 3);
    }

    [Fact]
    public void EditedSizeUsesFormulaHeightWhenOldBoxIsNonUniform()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 300f,
            currentHeight: 48f,
            originalRenderWidth: 160d,
            originalRenderHeight: 32d,
            newRenderWidth: 320f,
            newRenderHeight: 32f);

        Assert.Equal(480f, size.Width, 3);
        Assert.Equal(48f, size.Height, 3);
    }

    [Fact]
    public void EditedSmallInlineFormulaUsesRawPreviewHeightInsteadOfTwelvePointFloor()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 19f,
            currentHeight: 8.5f,
            originalRenderWidth: 25d,
            originalRenderHeight: 11d,
            newRenderWidth: 49f,
            newRenderHeight: 11f,
            originalFontSizePt: 11d,
            originalRenderFontSizePt: 11d);

        Assert.Equal(37.864f, size.Width, 3);
        Assert.True(size.Width > 36f);
    }

    [Fact]
    public void LegacyPictureConversionKeepsItsPhysicalHeight()
    {
        var size = OfficeFormulaSizing.EditedSize(
            currentWidth: 240f,
            currentHeight: 54f,
            originalRenderWidth: null,
            originalRenderHeight: null,
            newRenderWidth: 400f,
            newRenderHeight: 50f);

        Assert.Equal(432f, size.Width, 3);
        Assert.Equal(54f, size.Height, 3);
        Assert.Equal(8f, size.Width / size.Height, 3);
    }
}
