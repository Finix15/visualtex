using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordMathTypeLayoutValidationTests
{
    [Theory]
    [InlineData("", "\r")]
    [InlineData("", "\r\a")]
    public void OleOnlyParagraphIsDisplay(
        string prefix,
        string suffix)
    {
        var display = WordMathTypeLayoutValidation.IsStandaloneDisplayCandidate(
            prefix,
            suffix,
            inlineShapeCount: 1,
            existingMathCount: 0,
            out var reason);

        Assert.True(display);
        Assert.Equal("display-ole-only", reason);
    }

    [Theory]
    [InlineData("\t\t\t", "\r", "inline-tab-layout")]
    [InlineData(" ", "\r", "inline-whitespace-layout")]
    [InlineData("", " \r", "inline-whitespace-layout")]
    [InlineData("\v", "\r", "inline-manual-break")]
    [InlineData("biên độ tăng với ", " khi vật\r", "inline-surrounding-content")]
    public void LayoutContentAroundOleForcesInline(
        string prefix,
        string suffix,
        string expectedReason)
    {
        var display = WordMathTypeLayoutValidation.IsStandaloneDisplayCandidate(
            prefix,
            suffix,
            inlineShapeCount: 1,
            existingMathCount: 0,
            out var reason);

        Assert.False(display);
        Assert.Equal(expectedReason, reason);
    }

    [Theory]
    [InlineData(2, 0, "inline-object-count")]
    [InlineData(1, 1, "inline-existing-omath")]
    public void OtherObjectsOrExistingMathForceInline(
        int inlineShapeCount,
        int existingMathCount,
        string expectedReason)
    {
        var display = WordMathTypeLayoutValidation.IsStandaloneDisplayCandidate(
            string.Empty,
            "\r",
            inlineShapeCount,
            existingMathCount,
            out var reason);

        Assert.False(display);
        Assert.Equal(expectedReason, reason);
    }

    [Fact]
    public void ExistingInlineOmmlCarriageReturnsDoNotCountAsNewParagraphBreaks()
    {
        const string prefix = "A ";
        const string suffixWithInlineOmml = " B \rπ\r2\r C\r";

        Assert.False(WordMathTypeLayoutValidation.SurroundingsChanged(
            prefix,
            suffixWithInlineOmml,
            prefix,
            suffixWithInlineOmml));
    }

    [Theory]
    [InlineData("A unexpected\r", " B \rπ\r2\r C\r")]
    [InlineData("A ", " B \rπ\r2\r C\runexpected")]
    public void ChangedSurroundingContentIsRejected(
        string currentPrefix,
        string currentSuffix)
    {
        Assert.True(WordMathTypeLayoutValidation.SurroundingsChanged(
            "A ",
            " B \rπ\r2\r C\r",
            currentPrefix,
            currentSuffix));
    }
}
