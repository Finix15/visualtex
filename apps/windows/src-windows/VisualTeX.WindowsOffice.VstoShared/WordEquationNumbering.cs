using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed class EquationReferenceTarget
{
    public EquationReferenceTarget(
        string formulaId,
        int nativeReferenceItem,
        string numberText,
        string latexPreview,
        int position)
    {
        FormulaId = formulaId;
        NativeReferenceItem = nativeReferenceItem;
        NumberText = numberText;
        LatexPreview = latexPreview;
        Position = position;
    }

    public string FormulaId { get; }
    public int NativeReferenceItem { get; }
    public string NumberText { get; }
    public string LatexPreview { get; }
    public int Position { get; }

    public override string ToString() => $"({NumberText})    {LatexPreview}";
}

internal enum EquationReferenceStyle
{
    Parenthesized,
    EquationPrefix,
    NumberOnly,
}

internal sealed class EquationNumberFormat
{
    public const string ContinuousId = "continuous";
    public const string Heading1DotId = "heading1-dot";
    public const string Heading1DashId = "heading1-dash";
    public const string Heading2DotId = "heading2-dot";
    public const string Heading2DashId = "heading2-dash";

    private EquationNumberFormat(
        string id,
        string displayName,
        int headingLevel,
        string separator)
    {
        Id = id;
        DisplayName = displayName;
        HeadingLevel = headingLevel;
        Separator = separator;
    }

    public string Id { get; }
    public string DisplayName { get; }
    public int HeadingLevel { get; }
    public string Separator { get; }
    public bool UsesHeading => HeadingLevel > 0;

    public static EquationNumberFormat Resolve(string? id) => id switch
    {
        Heading1DotId => new EquationNumberFormat(Heading1DotId, "按章编号（1.1）", 1, "."),
        Heading1DashId => new EquationNumberFormat(Heading1DashId, "按章编号（1-1）", 1, "-"),
        Heading2DotId => new EquationNumberFormat(Heading2DotId, "按节编号（1.1.1）", 2, "."),
        Heading2DashId => new EquationNumberFormat(Heading2DashId, "按节编号（1.1-1）", 2, "-"),
        _ => new EquationNumberFormat(ContinuousId, "全文连续编号（1）", 0, string.Empty),
    };
}

internal static class WordEquationNumbering
{
    private const int WdTabAlignmentCenter = 1;
    private const int WdTabAlignmentRight = 2;
    private const int WdTabLeaderSpaces = 0;
    private const int WdFieldEmpty = -1;
    private const string EquationNumberFontName = "Cambria Math";
    private const string LegacyEquationSequenceName = "VisualTeXEquation";
    private const string EquationBookmarkPrefix = "VTEq_";
    private const string NativeCaptionBookmarkPrefix = "VTEqCap_";
    private const string NativeNumberBookmarkPrefix = "VTEqNum_";
    private const string EquationNumberFormatVariableName = "VisualTeXEquationNumberFormat";

    internal static string GetEquationNumberFormatId(Document document) =>
        ReadEquationNumberFormat(document).Id;

    internal static string GetEquationNumberFormatDisplayName(Document document) =>
        ReadEquationNumberFormat(document).DisplayName;

    internal static int SetEquationNumberFormat(Document document, string formatId)
    {
        SetEquationNumberFormatPreference(document, formatId);
        return Reconcile(document);
    }

    internal static void SetEquationNumberFormatPreference(
        Document document,
        string formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        WriteEquationNumberFormat(document, format.Id);
    }

    private static EquationNumberFormat ReadEquationNumberFormat(Document document)
    {
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                return EquationNumberFormat.Resolve(variable.Value);
            }
            catch (COMException)
            {
                return EquationNumberFormat.Resolve(null);
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    private static void WriteEquationNumberFormat(Document document, string formatId)
    {
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                variable.Value = formatId;
            }
            catch (COMException)
            {
                object value = formatId;
                variable = variables.Add(EquationNumberFormatVariableName, ref value);
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    public static void TryReconcile(Document document)
    {
        try { Reconcile(document); }
        catch
        {
            // Formula insertion/update is already durable. The user can retry
            // only the numbering command without duplicating or losing it.
        }
    }

    public static void TryReconcileFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata)
    {
        try
        {
            ReconcileFormula(document, formulaRange, formulaHeightPoints, metadata);
        }
        catch
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                throw;
            // The inserted or edited formula is already durable. The explicit
            // update-number command still performs a complete reconciliation.
        }
    }

    internal static void ReconcileFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata)
    {
        var hadNumberingArtifacts = HasFormulaNumberingArtifacts(
            document,
            metadata.FormulaId);
        if (metadata.DisplayMode != "block")
        {
            if (!hadNumberingArtifacts) return;
            RemoveFormulaNumberingArtifacts(document, metadata.FormulaId);
            UpdateMainStoryFields(document);
            UpdateNativeCrossReferences(document);
            return;
        }

        var formulaFontSizePoints = (float)FormulaFontSize.ResolveSemanticFontSize(metadata);
        if (metadata.Numbered)
        {
            ConfigureNumberedDisplayFormula(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId);
        }
        else if (!hadNumberingArtifacts)
        {
            // Ordinary unnumbered display formulas have no numbering artifacts
            // to remove. Configure only the local paragraph; scanning bookmarks,
            // fields and cross-references here made an unrelated 100-formula
            // document pay a document-wide cost for every single edit.
            ConfigureEquationParagraph(formulaRange, numbered: false);
            return;
        }
        else
        {
            ConfigureUnnumberedDisplayFormula(
                document,
                formulaRange,
                metadata.FormulaId);
        }

        // Insertion and editing already know the exact changed formula, so avoid
        // walking every VisualTeX object through COM. Word stores the visible
        // REF before the hidden SEQ caption in document order; a plain
        // Fields.Update therefore leaves every new visible number one step
        // behind. Refresh the lightweight native SEQ field inventory first,
        // then let Word update all REF results in one native pass.
        if (metadata.Numbered)
            UpdateNativeEquationSequenceFields(document);
        UpdateMainStoryFields(document);
        if (metadata.Numbered)
        {
            UpdateEquationNumberFields(
                document,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId);
        }
        // Formula-format conversion updates the document field collection. Word
        // can copy the hidden 1 pt SEQ target appearance back into body REF
        // results during that update, so always restore visible references at
        // the end of the same reconciliation transaction.
        UpdateNativeCrossReferences(document);
    }

    private static bool HasFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                || bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                || bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    internal static bool HasCompleteFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                && bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                && bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    private static void UpdateMainStoryFields(Document document)
    {
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            if (fields.Count > 0) fields.Update();
        }
        finally { Release(fields); }
    }

    internal static Range? EnsureNormalTypingParagraphAfterNumberedDisplay(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Paragraphs? captionParagraphs = null;
        Paragraph? captionParagraph = null;
        Range? captionParagraphRange = null;
        Paragraphs? documentParagraphs = null;
        Paragraph? typingParagraph = null;
        Range? typingRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return null;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            if (captionRange.End < Math.Max(content.Start, content.End - 1))
                return null;

            captionParagraphs = captionRange.Paragraphs;
            captionParagraph = captionParagraphs[1];
            captionParagraphRange = captionParagraph.Range;
            captionParagraphRange.InsertParagraphAfter();

            documentParagraphs = document.Paragraphs;
            typingParagraph = documentParagraphs[documentParagraphs.Count];
            typingRange = typingParagraph.Range.Duplicate;
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                typingRange.set_Style(ref normalStyle);
            }
            catch
            {
                // Documents with locked/custom style collections can reject the
                // built-in Normal style. Direct-format reset below still removes
                // the hidden caption's one-point formatting.
            }

            font = typingRange.Font;
            font.Reset();
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = typingRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            typingRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var result = typingRange;
            typingRange = null;
            return result;
        }
        finally
        {
            Release(format);
            Release(font);
            Release(typingRange);
            Release(typingParagraph);
            Release(documentParagraphs);
            Release(captionParagraphRange);
            Release(captionParagraph);
            Release(captionParagraphs);
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static void RemoveFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(document, formulaId);
    }

    public static int Reconcile(Document document)
    {
        var numberedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Freeze and verify every display-OMML identity before the first table
        // insertion, then migrate from the end of the document toward the start.
        // Later insertions therefore cannot drag the collapsed bookmark of an
        // equation that has not yet been processed.
        var ommlFormulaIds = GetOmmlDisplayFormulaIdsForStructuralEdit(document);
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var inlineCount = inlineShapes.Count;
            var numberedFormulaLocations = CaptureNumberedFormulaLocations(
                document,
                inlineShapes,
                ommlFormulaIds);
            RepairSharedNativeCaptionArtifacts(document, numberedFormulaLocations);

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                Range? formulaRange = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange = shape.Range;
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId);
                        numberedFormulaIds.Add(metadata.FormulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(
                            document,
                            formulaRange,
                            metadata.FormulaId);
                    }
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }

            foreach (var formulaId in ommlFormulaIds)
            {
                Range? formulaRange = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange =
                        WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                    var formulaHeightPoints =
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange);
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            formulaHeightPoints,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId);
                        numberedFormulaIds.Add(formulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(document, formulaRange, formulaId);
                    }
                }
                finally { Release(formulaRange); }
            }

            RemoveOrphanEquationArtifacts(document, numberedFormulaIds);
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);

            // Word caches SEQ results independently from REF results. After a
            // numbered formula is deleted, refresh every native Equation SEQ
            // field in document order before updating any visible or body REF
            // field. Otherwise a REF can continue displaying the removed
            // formula's old ordinal until Word performs a later global update.
            UpdateNativeEquationSequenceFields(document);
            // Writing chapter/section prefixes mutates the hidden caption
            // paragraphs. Word can invalidate a number bookmark whose range was
            // created before that mutation even though the SEQ field and caption
            // bookmark survive. Re-wrap every final SEQ result after all caption
            // text is stable so each formula keeps an independent REF target.
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);
            // Newly batch-numbered formulas create their visible REF fields
            // before the native number bookmarks are rewritten with the final
            // chapter/section prefix. Word can temporarily cache "reference
            // source not found" for those fresh fields. Refresh the complete
            // main story once after all target bookmarks are stable, then apply
            // formula-specific alignment below.
            UpdateMainStoryFields(document);

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId);
                }
                finally { Release(shape); }
            }
            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            WordOmmlFormulaStore.EstimateHeightPoints(bookmark),
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId);
                }
                finally { Release(bookmark); }
            }
            UpdateNativeCrossReferences(document);
        }
        finally { Release(inlineShapes); }

        return numberedFormulaIds.Count;
    }

    private sealed class OmmlStructuralFormulaEntry
    {
        public OmmlStructuralFormulaEntry(string formulaId, int position)
        {
            FormulaId = formulaId;
            Position = position;
        }

        public string FormulaId { get; }
        public int Position { get; }
    }

    private static IReadOnlyList<string> GetOmmlDisplayFormulaIdsForStructuralEdit(
        Document document)
    {
        var entries = new List<OmmlStructuralFormulaEntry>();
        var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document);
        foreach (var formulaId in formulaIds)
        {
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block") continue;
                range = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                entries.Add(new OmmlStructuralFormulaEntry(formulaId, range.Start));
            }
            finally { Release(range); }
        }
        var ordered = entries
            .OrderByDescending(entry => entry.Position)
            .ToArray();
        return ordered
            .Select(entry => entry.FormulaId)
            .ToArray();
    }

    private static FormulaMetadata? ReadMetadata(InlineShape shape) =>
        WordFormulaMetadataReader.TryRead(shape);

    private static void ConfigureNumberedDisplayFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId)
    {
        EnsureNumberedOmmlIsDisplay(formulaRange);
        EnsureStandardNumberedEquationTable(document, formulaRange, formulaId);
        ConfigureEquationParagraph(formulaRange, numbered: false);
        ConfigureNumberedEquationTable(formulaRange);
        var sequenceName = GetNativeEquationSequenceName(document);
        EnsureNativeCaption(document, formulaRange, formulaId, sequenceName);
        EnsureVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            formulaId);
    }

    private static void EnsureStandardNumberedEquationTable(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (IsNumberedEquationTable(formulaRange)) return;
        RemoveVisibleEquationNumber(document, formulaId);

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefixRange = null;
        Range? suffixRange = null;
        Range? sourceContent = null;
        Range? formattedContent = null;
        Range? tableAnchor = null;
        Range? documentContent = null;
        Range? sourceDeleteRange = null;
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Range? centerInsertion = null;
        Table? table = null;
        Columns? columns = null;
        Column? addedColumn = null;
        var sourceDeleted = false;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "VisualTeX cannot safely number a display formula spanning multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "VisualTeX cannot safely migrate this nonstandard table formula to the numbered layout.");

            object prefixStart = paragraphRange.Start;
            object prefixEnd = Math.Max(paragraphRange.Start, formulaRange.Start);
            prefixRange = document.Range(ref prefixStart, ref prefixEnd);
            var suffixStartPosition = Math.Min(formulaRange.End, paragraphRange.End);
            object suffixStart = suffixStartPosition;
            object suffixEnd = Math.Max(suffixStartPosition, paragraphRange.End - 1);
            suffixRange = document.Range(ref suffixStart, ref suffixEnd);
            if (!IsNumberingParagraphAdornment(prefixRange.Text)
                || !IsNumberingParagraphAdornment(suffixRange.Text))
                throw new InvalidOperationException(
                    "VisualTeX only batch-numbers display formulas that occupy their own paragraph.");

            var sourceStart = paragraphRange.Start;
            var sourceIsOmml = formulaRange.OMaths.Count > 0;
            if (!sourceIsOmml)
            {
                ConvertStandaloneOleParagraphToNumberedTable(
                    document,
                    paragraphRange,
                    formulaId,
                    formulaRange);
                return;
            }
            sourceContent = paragraphRange.Duplicate;
            sourceContent.End = Math.Max(sourceContent.Start, sourceContent.End - 1);
            formattedContent = sourceContent.FormattedText;

            // Insert one ordinary paragraph after the source formula. Word expands
            // paragraphRange to include that new paragraph, so paragraphRange.End
            // becomes the start of the following paragraph (which may itself begin
            // with another OMath). Keep the pre-insertion end as the table anchor;
            // using the expanded End would ask Tables.Add to operate inside the next
            // mathematical formula.
            var insertedParagraphStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            documentContent = document.Content;
            var tableAnchorPosition = Math.Max(
                documentContent.Start,
                Math.Min(insertedParagraphStart, documentContent.End - 1));
            object anchorStart = tableAnchorPosition;
            object anchorEnd = tableAnchorPosition;
            tableAnchor = document.Range(ref anchorStart, ref anchorEnd);
            try
            {
                table = document.Tables.Add(tableAnchor, 1, 3);
            }
            catch (Exception error)
            {
                throw new InvalidOperationException(
                    $"Word could not create the numbered OMML table for {formulaId} "
                    + $"at {sourceStart} (paragraph {paragraphRange.Start}-{paragraphRange.End}, "
                    + $"story {(int)paragraphRange.StoryType}).",
                    error);
            }
            columns = table.Columns;
            while (columns.Count < 3)
            {
                object appendAtRight = Type.Missing;
                addedColumn = columns.Add(ref appendAtRight);
                Release(addedColumn);
                addedColumn = null;
            }
            centerCell = table.Cell(1, 2);
            centerCellRange = centerCell.Range;
            centerInsertion = centerCellRange.Duplicate;
            centerInsertion.End = Math.Max(
                centerInsertion.Start,
                centerInsertion.End - 1);
            centerInsertion.Collapse(WdCollapseDirection.wdCollapseStart);
            centerInsertion.FormattedText = formattedContent;

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);

            DeleteOriginalStandaloneFormulaContent(
                document,
                table,
                formulaId,
                sourceIsOmml);
            sourceDeleted = true;

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);
            RemoveNumberingTableCenterDecorations(document, table, formulaRange);
            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);

            var ommlMetadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (ommlMetadata is not null && formulaRange.OMaths.Count > 0)
            {
                Bookmark? migratedBookmark = null;
                try
                {
                    migratedBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        formulaRange,
                        ommlMetadata,
                        replaceExisting: true);
                }
                finally { Release(migratedBookmark); }
            }
        }
        catch
        {
            if (!sourceDeleted && table is not null)
            {
                Range? rollbackRange = null;
                try
                {
                    rollbackRange = table.Range;
                    rollbackRange.Delete();
                }
                catch { }
                finally { Release(rollbackRange); }
            }
            throw;
        }
        finally
        {
            Release(addedColumn);
            Release(columns);
            Release(table);
            Release(centerInsertion);
            Release(centerCellRange);
            Release(centerCell);
            Release(sourceDeleteRange);
            Release(documentContent);
            Release(tableAnchor);
            Release(formattedContent);
            Release(sourceContent);
            Release(suffixRange);
            Release(prefixRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void ConvertStandaloneOleParagraphToNumberedTable(
        Document document,
        Range paragraphRange,
        string formulaId,
        Range formulaRange)
    {
        Range? conversionRange = null;
        Table? table = null;
        Columns? columns = null;
        Column? originalColumn = null;
        Column? addedColumn = null;
        try
        {
            conversionRange = paragraphRange.Duplicate;
            object separator = WdTableFieldSeparator.wdSeparateByParagraphs;
            object numRows = 1;
            object numColumns = 1;
            object initialColumnWidth = Type.Missing;
            object format = Type.Missing;
            object applyBorders = Type.Missing;
            object applyShading = Type.Missing;
            object applyFont = Type.Missing;
            object applyColor = Type.Missing;
            object applyHeadingRows = Type.Missing;
            object applyLastRow = Type.Missing;
            object applyFirstColumn = Type.Missing;
            object applyLastColumn = Type.Missing;
            object autoFit = false;
            object autoFitBehavior = WdAutoFitBehavior.wdAutoFitFixed;
            object defaultTableBehavior = WdDefaultTableBehavior.wdWord9TableBehavior;
            table = conversionRange.ConvertToTable(
                ref separator,
                ref numRows,
                ref numColumns,
                ref initialColumnWidth,
                ref format,
                ref applyBorders,
                ref applyShading,
                ref applyFont,
                ref applyColor,
                ref applyHeadingRows,
                ref applyLastRow,
                ref applyFirstColumn,
                ref applyLastColumn,
                ref autoFit,
                ref autoFitBehavior,
                ref defaultTableBehavior);

            columns = table.Columns;
            if (columns.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create the expected one-column OLE migration table.");
            originalColumn = columns[1];
            object beforeOriginal = originalColumn;
            addedColumn = columns.Add(ref beforeOriginal);
            Release(addedColumn);
            addedColumn = null;
            object appendAtRight = Type.Missing;
            addedColumn = columns.Add(ref appendAtRight);

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange);
            RemoveNumberingTableCenterDecorations(document, table, formulaRange);
            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange);
        }
        finally
        {
            Release(addedColumn);
            Release(originalColumn);
            Release(columns);
            Release(table);
            Release(conversionRange);
        }
    }

    private static void DeleteOriginalStandaloneFormulaContent(
        Document document,
        Table migratedTable,
        string formulaId,
        bool sourceIsOmml)
    {
        Bookmark? bookmark = null;
        Range? originalFormulaRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? editableRange = null;
        Range? migratedTableRange = null;
        try
        {
            migratedTableRange = migratedTable.Range;
            var migratedTableStart = migratedTableRange.Start;
            if (sourceIsOmml)
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException(
                        "Word lost the original VisualTeX OMML anchor during numbered-layout migration.");
                originalFormulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            }
            else
            {
                shapes = document.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Release(shape);
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is null
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    Range? candidate = null;
                    try
                    {
                        candidate = shape.Range;
                        var isMigratedCopy =
                            (bool)candidate.get_Information(WdInformation.wdWithInTable)
                            && candidate.Tables.Count > 0
                            && candidate.Tables[1].Range.Start == migratedTableStart;
                        if (isMigratedCopy) continue;
                        originalFormulaRange = candidate;
                        candidate = null;
                        break;
                    }
                    finally { Release(candidate); }
                }
                if (originalFormulaRange is null)
                    throw new InvalidOperationException(
                        "Word lost the original VisualTeX OLE object during numbered-layout migration.");
            }

            if ((bool)originalFormulaRange.get_Information(WdInformation.wdWithInTable)
                && originalFormulaRange.Tables.Count > 0
                && originalFormulaRange.Tables[1].Range.Start == migratedTableStart)
                throw new InvalidOperationException(
                    "Word could not distinguish the original formula from its numbered-layout copy.");

            paragraphs = originalFormulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The original VisualTeX display formula no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            editableRange = paragraphRange.Duplicate;
            editableRange.End = Math.Max(
                editableRange.Start,
                editableRange.End - 1);
            editableRange.Delete();
        }
        finally
        {
            Release(migratedTableRange);
            Release(editableRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shape);
            Release(shapes);
            Release(originalFormulaRange);
            Release(bookmark);
        }
    }

    private static bool IsNumberingParagraphAdornment(string? text)
    {
        foreach (var character in text ?? string.Empty)
        {
            if (character is '\t' or '\r' or '\n' or '\v'
                or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    private static void RefreshFormulaRangeInNumberedTable(
        Document document,
        Table table,
        string formulaId,
        Range formulaRange,
        bool allowUnanchoredOmml = false)
    {
        Cell? centerCell = null;
        Range? centerRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OMaths? maths = null;
        OMath? math = null;
        Bookmark? bookmark = null;
        Range? refreshed = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            shapes = centerRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is null
                    || !string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                refreshed = shape.Range;
                formulaRange.SetRange(refreshed.Start, refreshed.End);
                return;
            }

            if (allowUnanchoredOmml)
            {
                maths = centerRange.OMaths;
                if (maths.Count == 1)
                {
                    math = maths[1];
                    refreshed = math.Range;
                    formulaRange.SetRange(refreshed.Start, refreshed.End);
                    return;
                }
            }

            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (bookmark is not null)
            {
                refreshed = WordOmmlFormulaStore.GetEquationRange(bookmark);
                if ((bool)refreshed.get_Information(WdInformation.wdWithInTable)
                    && refreshed.Tables.Count > 0
                    && refreshed.Tables[1].Range.Start == table.Range.Start)
                {
                    formulaRange.SetRange(refreshed.Start, refreshed.End);
                    return;
                }
            }

            throw new InvalidOperationException(
                "Word did not preserve the VisualTeX formula while creating its numbered layout.");
        }
        finally
        {
            Release(refreshed);
            Release(bookmark);
            Release(math);
            Release(maths);
            Release(shape);
            Release(shapes);
            Release(centerRange);
            Release(centerCell);
        }
    }

    private static void RemoveNumberingTableCenterDecorations(
        Document document,
        Table table,
        Range formulaRange)
    {
        Cell? centerCell = null;
        Range? centerRange = null;
        Range? characterRange = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            for (var position = centerRange.End - 2;
                 position >= centerRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                characterRange = document.Range(ref characterStart, ref characterEnd);
                if (string.Equals(characterRange.Text, "\t", StringComparison.Ordinal)
                    || string.Equals(characterRange.Text, "\v", StringComparison.Ordinal))
                    characterRange.Delete();
                Release(characterRange);
                characterRange = null;
            }
        }
        finally
        {
            Release(characterRange);
            Release(centerRange);
            Release(centerCell);
        }
    }

    private static void EnsureNumberedOmmlIsDisplay(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? refreshed = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count == 0) return;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
            {
                math.Type = WdOMathType.wdOMathDisplay;
                math.BuildUp();
            }
            refreshed = math.Range;
            formulaRange.SetRange(refreshed.Start, refreshed.End);
        }
        finally
        {
            Release(refreshed);
            Release(math);
            Release(maths);
        }
    }

    private static bool IsNumberedEquationTable(Range formulaRange)
    {
        try
        {
            return (bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                && formulaRange.Tables.Count > 0
                && formulaRange.Tables[1].Columns.Count >= 3;
        }
        catch { return false; }
    }

    private static void ConfigureNumberedEquationTable(Range formulaRange)
    {
        Table? table = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? numberFormat = null;
        Borders? borders = null;
        try
        {
            table = formulaRange.Tables[1];
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            table.TopPadding = 0f;
            table.BottomPadding = 0f;
            try { table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed); } catch { }
            columns = table.Columns;
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            centerColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            rightColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            leftColumn.PreferredWidth = 20f;
            centerColumn.PreferredWidth = 60f;
            rightColumn.PreferredWidth = 20f;
            borders = table.Borders;
            borders.Enable = 0;
            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            numberCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            numberFormat = numberRange.ParagraphFormat;
            centerFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            numberFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            centerFormat.LeftIndent = centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;
            numberFormat.LeftIndent = numberFormat.RightIndent = 0f;
            numberFormat.FirstLineIndent = 0f;
            centerFormat.SpaceBefore = centerFormat.SpaceAfter = 0f;
            numberFormat.SpaceBefore = numberFormat.SpaceAfter = 0f;
            centerFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            numberFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            try { centerFormat.DisableLineHeightGrid = -1; } catch { }
            try { numberFormat.DisableLineHeightGrid = -1; } catch { }
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(borders);
            Release(numberFormat);
            Release(centerFormat);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(table);
        }
    }

    private static void ConfigureUnnumberedDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(document, formulaId);
        RemoveLeadingEquationTab(document, formulaRange);
        ConfigureEquationParagraph(formulaRange, numbered: false);
    }

    private static void ConfigureEquationParagraph(Range formulaRange, bool numbered)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        ListFormat? listFormat = null;
        OMaths? maths = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            var nativeOmmlParagraph = false;
            try
            {
                maths = formulaRange.OMaths;
                nativeOmmlParagraph = maths.Count > 0
                    && !(bool)formulaRange.get_Information(WdInformation.wdWithInTable);
            }
            catch { }
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            if (!nativeOmmlParagraph)
            {
                // OLE/picture formulas carry their own transparent preview
                // margins, so their host paragraph must stay compact. Native
                // OMML formulas have no image padding and should retain the
                // document paragraph style exactly as a formula inserted by
                // Word itself would.
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch
            {
                // Protected/custom stories can reject list normalization. The
                // page-break flags above still remove Word's black paragraph
                // marker when formatting marks are shown.
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();

            if (!numbered)
            {
                format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                return;
            }

            var pageWidth = 612f;
            var leftMargin = 72f;
            var rightMargin = 72f;
            try
            {
                sections = paragraphRange.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                    pageWidth = pageSetup.PageWidth;
                    leftMargin = pageSetup.LeftMargin;
                    rightMargin = pageSetup.RightMargin;
                }
            }
            catch
            {
                // Standard US Letter and one-inch margins are a safe fallback
                // for protected/custom stories without an exposed PageSetup.
            }

            var positions = CalculateEquationTabStops(pageWidth, leftMargin, rightMargin, 0, 0);
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            tabStops.Add(
                positions.Center,
                (WdTabAlignment)WdTabAlignmentCenter,
                (WdTabLeader)WdTabLeaderSpaces);
            tabStops.Add(
                positions.Right,
                (WdTabAlignment)WdTabAlignmentRight,
                (WdTabLeader)WdTabLeaderSpaces);
        }
        finally
        {
            Release(maths);
            Release(listFormat);
            Release(tabStops);
            Release(format);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static (float Center, float Right) CalculateEquationTabStops(
        float pageWidth,
        float leftMargin,
        float rightMargin,
        float leftIndent,
        float rightIndent)
    {
        var availableWidth = Math.Max(
            72f,
            pageWidth
                - Math.Max(0f, leftMargin)
                - Math.Max(0f, rightMargin)
                - Math.Max(0f, leftIndent)
                - Math.Max(0f, rightIndent));
        return (availableWidth / 2f, availableWidth);
    }

    internal static int CalculateEquationNumberFontPosition(
        float formulaHeightPoints,
        float numberFontSizePoints)
    {
        if (float.IsNaN(formulaHeightPoints)
            || float.IsInfinity(formulaHeightPoints)
            || formulaHeightPoints <= 0)
            return 0;
        if (float.IsNaN(numberFontSizePoints)
            || float.IsInfinity(numberFontSizePoints)
            || numberFontSizePoints <= 0
            || numberFontSizePoints > 256)
            numberFontSizePoints = 11f;
        return Math.Max(
            0,
            (int)Math.Round(
                (formulaHeightPoints - numberFontSizePoints) / 2f,
                MidpointRounding.AwayFromZero));
    }

    internal static (string Text, int FieldOffset) EquationNumberScaffold() => ("\t()", 2);

    private static void EnsureLeadingEquationTab(Document document, Range formulaRange)
    {
        Range? preceding = null;
        Range? insertion = null;
        try
        {
            var insertionPosition = formulaRange.Start;
            if (formulaRange.Start > 0)
            {
                object precedingStart = formulaRange.Start - 1;
                object precedingEnd = formulaRange.Start;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;

                // A display OMath is preceded by Word's vertical-tab math
                // separator (0x0B). Its layout tab therefore sits one character
                // earlier and must be inspected/inserted outside the OMath edge.
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                {
                    insertionPosition = formulaRange.Start - 1;
                    if (formulaRange.Start > 1)
                    {
                        preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                        if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;
                    }
                }
            }

            object start = insertionPosition;
            object end = insertionPosition;
            insertion = document.Range(ref start, ref end);
            insertion.Text = "\t";
        }
        finally
        {
            Release(insertion);
            Release(preceding);
        }
    }

    private static void RemoveLeadingEquationTab(Document document, Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? preceding = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start <= paragraphRange.Start) return;
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
            {
                preceding.Delete();
                return;
            }
            if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                && formulaRange.Start - 2 >= paragraphRange.Start)
            {
                preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
                    preceding.Delete();
            }
        }
        finally
        {
            Release(preceding);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void EnsureNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName)
    {
        if (TryGetNativeCaptionRanges(
                document,
                formulaId,
                nativeSequenceName,
                out var captionRange,
                out var numberRange)
            && captionRange is not null
            && numberRange is not null)
        {
            try { StyleNativeCaption(captionRange, numberRange); }
            finally
            {
                Release(numberRange);
                Release(captionRange);
            }
            return;
        }
        Release(numberRange);
        Release(captionRange);

        RemoveNativeCaption(document, formulaId);
        CreateNativeCaption(document, formulaRange, formulaId, nativeSequenceName);
    }

    private static void CreateNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName)
    {
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Table? formulaTable = null;
        Range? formulaTableRange = null;
        Range? fieldInsertion = null;
        Fields? fields = null;
        Field? captionField = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Bookmarks? bookmarks = null;
        try
        {
            // Word's Range.InsertCaption mutates the equation paragraph. For a
            // trailing inline OMath it moves the ordinary run after the equation
            // into a new paragraph, so the visible REF field is subsequently
            // absorbed into m:oMath. Build the native SEQ caption in a dedicated
            // hidden paragraph instead and leave the equation paragraph intact.
            int captionStart;
            if (IsNumberedEquationTable(formulaRange))
            {
                formulaTable = formulaRange.Tables[1];
                formulaTableRange = formulaTable.Range;
                formulaTableRange.InsertParagraphAfter();
                // InsertParagraphAfter expands the table range so its new End
                // is outside the table. The original End (and new End - 1) are
                // structural cell boundaries that reject Fields.Add.
                captionStart = formulaTableRange.End;
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                captionStart = formulaParagraphRange.End;
                formulaParagraphRange.InsertParagraphAfter();
            }

            object insertionStart = captionStart;
            object insertionEnd = captionStart;
            fieldInsertion = document.Range(ref insertionStart, ref insertionEnd);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"SEQ {nativeSequenceName} \\* ARABIC";
            object preserveFormatting = true;
            captionField = fields.Add(
                fieldInsertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            captionField.Update();
            numberRange = captionField.Result;
            paragraphs = numberRange.Paragraphs;
            paragraph = paragraphs[1];
            captionRange = paragraph.Range;
            try
            {
                object captionStyle = WdBuiltinStyle.wdStyleCaption;
                captionRange.set_Style(ref captionStyle);
            }
            catch
            {
                // Some locked/custom documents reject assigning the built-in
                // caption style. The SEQ field and bookmarks remain valid.
            }

            bookmarks = document.Bookmarks;
            bookmarks.Add(NativeNumberBookmarkName(formulaId), numberRange);
            bookmarks.Add(NativeCaptionBookmarkName(formulaId), captionRange);
            StyleNativeCaption(captionRange, numberRange);
        }
        finally
        {
            Release(bookmarks);
            Release(paragraph);
            Release(paragraphs);
            Release(captionRange);
            Release(numberRange);
            Release(captionField);
            Release(fields);
            Release(fieldInsertion);
            Release(formulaTableRange);
            Release(formulaTable);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static Field? FindNewNativeEquationField(
        Document document,
        string nativeSequenceName,
        ISet<int> existingPositions,
        int formulaPosition)
    {
        Fields? fields = null;
        Field? result = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? fieldResult = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    fieldResult = field.Result;
                    if (existingPositions.Contains(fieldResult.Start)) continue;
                    var distance = Math.Abs(fieldResult.Start - formulaPosition);
                    if (distance >= bestDistance) continue;
                    Release(result);
                    result = field;
                    field = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
            return result;
        }
        finally
        {
            Release(fields);
        }
    }

    private static void StyleNativeCaption(Range captionRange, Range numberRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Frames? frames = null;
        Frame? frame = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Borders? borders = null;
        try
        {
            // Keep the native SEQ target in ordinary visible formatting. Word's
            // own InsertCrossReference creates a plain REF ... \\h field, which
            // inherits the target formatting on every F9 update. Hiding this
            // paragraph with white 1 pt text therefore made native references
            // white and one point as well.
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = 11f;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = 11f;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch { }

            // A legacy Word Frame remains in the main document story, so Word's
            // native Cross-reference dialog and plain REF fields still recognize
            // the SEQ target. Negative frame coordinates are clamped by Word to
            // the top-left of the page, which exposed the number and frame edge.
            // Use an exact 0.1 pt clipping frame at the bottom-right page boundary
            // instead. The SEQ retains normal black 11 pt formatting for native
            // REF inheritance, while its rendered content lies beyond the page.
            var document = captionRange.Document;
            RemoveLegacyEmptyCaptionFrames(document, captionRange);
            frames = captionRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(captionRange);
            sections = captionRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            const float clippedFrameSize = 0.1f;
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = clippedFrameSize;
            frame.Height = clippedFrameSize;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = Math.Max(0f, pageSetup.PageWidth - clippedFrameSize);
            frame.VerticalPosition = Math.Max(0f, pageSetup.PageHeight - clippedFrameSize);
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = captionRange.Borders;
            borders.Enable = 0;
        }
        finally
        {
            Release(borders);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(frame);
            Release(frames);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    private static void RemoveLegacyEmptyCaptionFrames(
        Document document,
        Range keepRange)
    {
        Frames? frames = null;
        try
        {
            frames = document.Frames;
            for (var index = frames.Count; index >= 1; index--)
            {
                Frame? candidate = null;
                Range? range = null;
                Fields? fields = null;
                try
                {
                    candidate = frames[index];
                    range = candidate.Range;
                    if (range.Start <= keepRange.End && range.End >= keepRange.Start)
                        continue;
                    fields = range.Fields;
                    var text = (range.Text ?? string.Empty)
                        .Trim('\r', '\n', '\t', '\v', ' ');
                    var oldVisualTeXFrame =
                        candidate.HorizontalPosition <= -999f
                        && candidate.VerticalPosition <= -999f
                        && candidate.Width >= 70f
                        && candidate.Height >= 17f;
                    if (oldVisualTeXFrame
                        && fields.Count == 0
                        && string.IsNullOrEmpty(text))
                        candidate.Delete();
                }
                finally
                {
                    Release(fields);
                    Release(range);
                    Release(candidate);
                }
            }
        }
        finally { Release(frames); }
    }

    private static bool TryGetNativeCaptionRanges(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out Range? captionRange,
        out Range? numberRange)
    {
        captionRange = null;
        numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Field? nativeField = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName)) return false;
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            numberRange = numberBookmark.Range;
            nativeField = FindNativeEquationFieldAtRange(
                document,
                numberRange,
                nativeSequenceName);
            if (nativeField is not null) return true;
            Release(numberRange);
            numberRange = null;
            Release(captionRange);
            captionRange = null;
            return false;
        }
        finally
        {
            Release(nativeField);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void EnsureVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        if (HasVisibleEquationNumber(
                document,
                formulaRange,
                formulaId,
                targetBookmarkName)) return;
        RemoveVisibleEquationNumber(document, formulaId);
        InsertVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            formulaId,
            targetBookmarkName);
    }

    private static bool HasVisibleEquationNumber(
        Document document,
        Range formulaRange,
        string formulaId,
        string targetBookmarkName)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        OMaths? maths = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? numberParagraph = null;
        Range? numberParagraphRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return false;
            bookmark = bookmarks[name];
            range = bookmark.Range;

            // Older builds inserted the visible REF field at OMath.Range.End.
            // Word still considers that position part of m:oMath, so the tab,
            // parentheses and number became equation content. A valid number is
            // an ordinary Word run after the formula, never a child of OMML.
            if (range.Start < formulaRange.End) return false;
            maths = range.OMaths;
            if (maths.Count > 0) return false;
            var tableLayout = IsNumberedEquationTable(formulaRange);
            var visibleText = range.Text ?? string.Empty;
            var expectedPrefix = tableLayout ? "(" : "\t(";
            if (!visibleText.StartsWith(expectedPrefix, StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            if (tableLayout)
            {
                if (!(bool)range.get_Information(WdInformation.wdWithInTable)
                    || range.Tables.Count == 0
                    || range.Tables[1].Range.Start != formulaRange.Tables[1].Range.Start)
                    return false;
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                numberParagraphs = range.Paragraphs;
                numberParagraph = numberParagraphs[1];
                numberParagraphRange = numberParagraph.Range;
                if (formulaParagraphRange.Start != numberParagraphRange.Start) return false;
            }

            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (IsReferenceToBookmark(code.Text, targetBookmarkName)) return true;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        finally
        {
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(maths);
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void InsertVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        string targetBookmarkName)
    {
        Range? scaffoldRange = null;
        Range? fieldRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? fieldResult = null;
        Range? bookmarkRange = null;
        Bookmarks? bookmarks = null;
        try
        {
            var tableLayout = IsNumberedEquationTable(formulaRange);
            var suffixStart = PrepareEquationNumberInsertionPosition(formulaRange);
            var scaffold = tableLayout ? (Text: "()", FieldOffset: 1) : EquationNumberScaffold();
            object suffixStartObject = suffixStart;
            object suffixEndObject = suffixStart;
            scaffoldRange = document.Range(ref suffixStartObject, ref suffixEndObject);
            scaffoldRange.Text = scaffold.Text;

            var fieldStart = suffixStart + scaffold.FieldOffset;
            object fieldStartObject = fieldStart;
            object fieldEndObject = fieldStart;
            fieldRange = document.Range(ref fieldStartObject, ref fieldEndObject);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {targetBookmarkName} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            NormalizeReferenceResult(field);
            fieldResult = field.Result;

            // Do not rely on scaffoldRange.End. Word 2013/2016 perpetual,
            // Microsoft 365 and compatibility mode expand a Range around an
            // inserted field differently. Resolve the actual closing parenthesis
            // from the document text so the bookmark always contains the tab,
            // both brackets and the complete REF field.
            bookmarkRange = ResolveEquationNumberLabelRange(
                document,
                suffixStart,
                scaffoldRange,
                tableLayout);
            bookmarks = document.Bookmarks;
            bookmarks.Add(EquationBookmarkName(formulaId), bookmarkRange);
            if (tableLayout)
            {
                ParagraphFormat? format = null;
                try
                {
                    format = bookmarkRange.ParagraphFormat;
                    format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }
                finally { Release(format); }
            }
            AlignEquationNumberVertically(
                bookmarkRange,
                tableLayout ? 0f : formulaHeightPoints,
                formulaFontSizePoints);
        }
        finally
        {
            Release(bookmarks);
            Release(bookmarkRange);
            Release(fieldResult);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(scaffoldRange);
        }
    }

    private static Range ResolveEquationNumberLabelRange(
        Document document,
        int labelStart,
        Range scaffoldRange,
        bool tableLayout)
    {
        Range? character = null;
        Range? candidate = null;
        try
        {
            var searchEnd = Math.Min(document.Content.End, labelStart + 512);
            for (var position = labelStart + 1; position < searchEnd; position++)
            {
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(ref characterStart, ref characterEnd);
                if (!string.Equals(character.Text, ")", StringComparison.Ordinal))
                {
                    Release(character);
                    character = null;
                    continue;
                }

                object candidateStart = labelStart;
                object candidateEnd = character.End;
                candidate = document.Range(ref candidateStart, ref candidateEnd);
                var text = candidate.Text ?? string.Empty;
                var expectedPrefix = tableLayout ? "(" : "\t(";
                if (text.StartsWith(expectedPrefix, StringComparison.Ordinal)
                    && text.EndsWith(")", StringComparison.Ordinal))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                Release(candidate);
                candidate = null;
                Release(character);
                character = null;
            }

            object fallbackStart = labelStart;
            object fallbackEnd = scaffoldRange.End;
            candidate = document.Range(ref fallbackStart, ref fallbackEnd);
            var fallbackText = candidate.Text ?? string.Empty;
            var fallbackPrefix = tableLayout ? "(" : "\t(";
            if (!fallbackText.StartsWith(fallbackPrefix, StringComparison.Ordinal)
                || !fallbackText.EndsWith(")", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word did not preserve the complete VisualTeX equation-number label.");
            var fallback = candidate;
            candidate = null;
            return fallback;
        }
        finally
        {
            Release(candidate);
            Release(character);
        }
    }

    private static int PrepareEquationNumberInsertionPosition(Range formulaRange)
    {
        if (IsNumberedEquationTable(formulaRange))
        {
            Table? table = null;
            Cell? cell = null;
            Range? cellRange = null;
            Range? editableRange = null;
            try
            {
                table = formulaRange.Tables[1];
                cell = table.Cell(1, 3);
                cellRange = cell.Range;
                // This cell is reserved exclusively for the generated number.
                // Word can leave an empty paragraph behind when a REF field is
                // removed/reconciled. Centering that empty paragraph together
                // with the new number pushes the visible number downward.
                // Clear everything except the structural cell mark, which
                // normalizes the cell to exactly one paragraph, then insert at
                // the beginning of that paragraph.
                editableRange = cellRange.Duplicate;
                editableRange.End = Math.Max(
                    editableRange.Start,
                    editableRange.End - 1);
                editableRange.Text = string.Empty;
                return cellRange.Start;
            }
            finally
            {
                Release(editableRange);
                Release(cellRange);
                Release(cell);
                Release(table);
            }
        }
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return Math.Max(paragraphRange.Start, paragraphRange.End - 1);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RemoveVisibleEquationNumber(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Range? trailing = null;
        OMaths? maths = null;
        OMath? containingMath = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            var start = range.Start;
            var text = range.Text ?? string.Empty;
            try
            {
                maths = range.OMaths;
                if (maths.Count > 0) containingMath = maths[1];
            }
            catch { }
            range.Delete();

            // Legacy OMML numbering bookmarks stopped immediately before the
            // closing parenthesis. Remove that orphan as part of migration.
            if (!text.EndsWith(")", StringComparison.Ordinal)
                && start < document.Content.End)
            {
                object trailingStart = start;
                object trailingEnd = Math.Min(document.Content.End, start + 1);
                trailing = document.Range(ref trailingStart, ref trailingEnd);
                if (string.Equals(trailing.Text, ")", StringComparison.Ordinal))
                    trailing.Delete();
            }
            try { containingMath?.BuildUp(); } catch { }
        }
        finally
        {
            Release(containingMath);
            Release(maths);
            Release(trailing);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveNativeCaption(Document document, string formulaId)
    {
        DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
        DeleteBookmarkedRange(document, NativeCaptionBookmarkName(formulaId));
    }

    private static void RemoveOrphanEquationArtifacts(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        RemoveOrphanBookmarks(
            document,
            EquationBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeCaptionBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeNumberBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: false);
    }

    private static void RemoveOrphanBookmarks(
        Document document,
        string prefix,
        ISet<string> activeFormulaIds,
        bool deleteRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(bookmark.Name, prefix, out var formulaId)
                        || activeFormulaIds.Contains(formulaId))
                        continue;
                    if (deleteRange)
                    {
                        range = bookmark.Range;
                        range.Delete();
                    }
                    else
                    {
                        bookmark.Delete();
                    }
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static void UpdateEquationNumberFields(
        Document document,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId)
    {
        UpdateFieldInBookmark(
            document,
            EquationBookmarkName(formulaId),
            code => IsReferenceToBookmark(code, NativeNumberBookmarkName(formulaId)));

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return;
            bookmark = bookmarks[visibleName];
            range = bookmark.Range;
            // A numbered table centers both cells vertically. Applying the
            // legacy height-derived baseline shift as well makes OLE and OMML
            // numbers disagree because their measured heights differ.
            AlignEquationNumberVertically(
                range,
                IsNumberedEquationTable(range) ? 0f : formulaHeightPoints,
                formulaFontSizePoints);
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static IReadOnlyDictionary<string, FormulaDocumentLocation>
        CaptureNumberedFormulaLocations(
            Document document,
            InlineShapes inlineShapes,
            IReadOnlyList<string> ommlFormulaIds)
    {
        var result = new Dictionary<string, FormulaDocumentLocation>(
            StringComparer.OrdinalIgnoreCase);
        for (var index = 1; index <= inlineShapes.Count; index++)
        {
            InlineShape? shape = null;
            Range? range = null;
            try
            {
                shape = inlineShapes[index];
                var metadata = ReadMetadata(shape);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                range = shape.Range;
                result[metadata.FormulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }

        foreach (var formulaId in ommlFormulaIds)
        {
            Bookmark? bookmark = null;
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                result[formulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }
        return result;
    }

    private static void RepairSharedNativeCaptionArtifacts(
        Document document,
        IReadOnlyDictionary<string, FormulaDocumentLocation> formulaLocations)
    {
        var ownersByArtifactRange = new Dictionary<string, HashSet<string>>(
            StringComparer.Ordinal);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    string artifactKind;
                    string formulaId;
                    if (TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out formulaId))
                    {
                        artifactKind = "number";
                    }
                    else if (TryFormulaIdFromBookmark(
                                 bookmark.Name,
                                 NativeCaptionBookmarkPrefix,
                                 out formulaId))
                    {
                        artifactKind = "caption";
                    }
                    else
                    {
                        continue;
                    }
                    if (!formulaLocations.ContainsKey(formulaId)) continue;

                    range = bookmark.Range;
                    var key = artifactKind + ":" + range.Start + ":" + range.End;
                    if (!ownersByArtifactRange.TryGetValue(key, out var owners))
                    {
                        owners = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        ownersByArtifactRange[key] = owners;
                    }
                    owners.Add(formulaId);
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        var conflictedFormulaIds = ownersByArtifactRange.Values
            .Where(owners => owners.Count > 1)
            .SelectMany(owners => owners)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (conflictedFormulaIds.Count == 0) return;

        // A shared caption paragraph cannot safely keep one "owner": deleting or
        // rewriting that paragraph invalidates every overlapping VTEqCap/VTEqNum
        // bookmark. Remove the complete numbering scaffold for every participant
        // first, then the normal reconciliation loop rebuilds one independent
        // visible REF and one independent native SEQ target per formula.
        foreach (var formulaId in conflictedFormulaIds
                     .OrderByDescending(id => formulaLocations[id].Start))
        {
            RemoveVisibleEquationNumber(document, formulaId);
            RemoveNativeCaption(document, formulaId);
        }
    }

    private static void RebuildNativeNumberBookmarksFromCaptions(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formulaId in numberedFormulaIds)
            {
                Bookmark? captionBookmark = null;
                Range? captionRange = null;
                Field? sequenceField = null;
                Range? fieldResult = null;
                Range? completeNumberRange = null;
                Bookmark? rebuiltNumberBookmark = null;
                try
                {
                    var captionName = NativeCaptionBookmarkName(formulaId);
                    if (!bookmarks.Exists(captionName))
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native caption after reconciliation.");
                    captionBookmark = bookmarks[captionName];
                    captionRange = captionBookmark.Range;
                    sequenceField = FindNativeEquationFieldInCaption(
                        captionRange,
                        nativeSequenceName);
                    if (sequenceField is null)
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native SEQ field inside its caption.");
                    fieldResult = sequenceField.Result;
                    // After heading-aware renumbering, the chapter/section
                    // prefix is ordinary text immediately before the SEQ field.
                    // Rebuilding from Field.Result alone silently changes a
                    // target such as "1-1" back to "1". Preserve the complete
                    // visible caption number from the paragraph start through
                    // the field result; continuous numbering naturally has an
                    // empty prefix and uses the same range.
                    completeNumberRange = document.Range(
                        captionRange.Start,
                        fieldResult.End);

                    var numberName = NativeNumberBookmarkName(formulaId);
                    if (bookmarks.Exists(numberName))
                    {
                        Bookmark? existing = null;
                        try
                        {
                            existing = bookmarks[numberName];
                            existing.Delete();
                        }
                        finally { Release(existing); }
                    }
                    rebuiltNumberBookmark = bookmarks.Add(
                        numberName,
                        completeNumberRange);
                }
                finally
                {
                    Release(rebuiltNumberBookmark);
                    Release(completeNumberRange);
                    Release(fieldResult);
                    Release(sequenceField);
                    Release(captionRange);
                    Release(captionBookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static Field? FindNativeEquationFieldInCaption(
        Range captionRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = captionRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(
                            code.Text,
                            nativeSequenceName))
                        continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private sealed class FormulaDocumentLocation
    {
        public FormulaDocumentLocation(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; }
        public int End { get; }
    }

    private sealed class NativeEquationCaptionEntry
    {
        public NativeEquationCaptionEntry(string formulaId, int position)
        {
            FormulaId = formulaId;
            Position = position;
        }

        public string FormulaId { get; }
        public int Position { get; }
    }

    private sealed class HeadingNumberAnchor
    {
        public HeadingNumberAnchor(int position, string numberText)
        {
            Position = position;
            NumberText = numberText;
        }

        public int Position { get; }
        public string NumberText { get; }
    }

    private sealed class HeadingParagraphCandidate
    {
        public HeadingParagraphCandidate(
            int position,
            int outlineLevel,
            string listNumber,
            string textNumber)
        {
            Position = position;
            OutlineLevel = outlineLevel;
            ListNumber = listNumber;
            TextNumber = textNumber;
        }

        public int Position { get; }
        public int OutlineLevel { get; }
        public string ListNumber { get; }
        public string TextNumber { get; }
        public string ExplicitNumber =>
            !string.IsNullOrWhiteSpace(ListNumber) ? ListNumber : TextNumber;
    }

    private static void UpdateNativeEquationSequenceFields(Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var format = ReadEquationNumberFormat(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchors(document, format.HeadingLevel)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();

        foreach (var caption in captions)
        {
            var scope = ResolveEquationNumberScope(
                caption.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
            localOrdinal++;
            ordinalByScope[scope.ScopePosition] = localOrdinal;
            UpdateNativeEquationCaptionNumber(
                document,
                caption.FormulaId,
                nativeSequenceName,
                localOrdinal,
                scope.Prefix);
        }
    }

    private static IReadOnlyList<NativeEquationCaptionEntry> GetNativeEquationCaptionEntries(
        Document document,
        string nativeSequenceName)
    {
        var result = new List<NativeEquationCaptionEntry>();
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? numberRange = null;
                Field? field = null;
                Range? fieldResult = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                        continue;
                    numberRange = bookmark.Range;
                    field = FindNativeEquationFieldAtRange(
                        document,
                        numberRange,
                        nativeSequenceName);
                    if (field is null) continue;
                    fieldResult = field.Result;
                    result.Add(new NativeEquationCaptionEntry(formulaId, fieldResult.Start));
                }
                finally
                {
                    Release(fieldResult);
                    Release(field);
                    Release(numberRange);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(item => item.Position).ToArray();
    }

    private static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchors(
        Document document,
        int targetLevel)
    {
        var candidates = new List<HeadingParagraphCandidate>();
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Paragraph? paragraph = null;
                Range? range = null;
                ListFormat? listFormat = null;
                Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    var outlineLevel = (int)paragraph.OutlineLevel;
                    if (outlineLevel < 1 || outlineLevel > targetLevel) continue;

                    range = paragraph.Range;
                    // Numbering tables and clipped native-caption paragraphs can
                    // inherit Heading 1 from the source formula paragraph. They
                    // are implementation scaffolds, never chapter boundaries.
                    if ((bool)range.get_Information(WdInformation.wdWithInTable))
                        continue;
                    try
                    {
                        frames = range.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch { }

                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = range.ListFormat;
                        listNumber = NormalizeHeadingNumberText(listFormat.ListString);
                    }
                    catch { }
                    candidates.Add(new HeadingParagraphCandidate(
                        range.Start,
                        outlineLevel,
                        listNumber,
                        ParseHeadingNumberFromText(range.Text, outlineLevel)));
                }
                finally
                {
                    Release(frames);
                    Release(listFormat);
                    Release(range);
                    Release(paragraph);
                }
            }
        }
        finally { Release(paragraphs); }

        var result = new List<HeadingNumberAnchor>();
        var usesExplicitNumbering = candidates.Any(candidate =>
            !string.IsNullOrWhiteSpace(candidate.ExplicitNumber));
        var counters = new int[10];
        foreach (var candidate in candidates)
        {
            var numberText = candidate.ExplicitNumber;
            if (string.IsNullOrWhiteSpace(numberText))
            {
                // In a manually/automatically numbered document, an unnumbered
                // Heading 1 is usually the document title, not "chapter 1".
                if (usesExplicitNumbering) continue;
                counters[candidate.OutlineLevel]++;
                for (var deeper = candidate.OutlineLevel + 1;
                     deeper < counters.Length;
                     deeper++)
                    counters[deeper] = 0;
                var parts = new List<string>();
                for (var level = 1; level <= candidate.OutlineLevel; level++)
                    parts.Add(Math.Max(1, counters[level]).ToString());
                numberText = string.Join(".", parts);
            }

            if (candidate.OutlineLevel < targetLevel)
            {
                var missingLevels = targetLevel - candidate.OutlineLevel;
                numberText += string.Concat(
                    Enumerable.Repeat(".0", missingLevels));
            }
            result.Add(new HeadingNumberAnchor(candidate.Position, numberText));
        }
        return result;
    }

    private static string NormalizeHeadingNumberText(string? value)
    {
        var text = (value ?? string.Empty)
            .Replace("\t", string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Trim();
        return text.TrimEnd(' ', '.', '-', '–', '—', '、', ')', '）');
    }

    private static string ParseHeadingNumberFromText(string? value, int outlineLevel)
    {
        var text = (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        // Support common manually typed headings such as "8. 多元微积分",
        // "8-1 小节" and "第 8 章" when Word's ListString is empty.
        var match = Regex.Match(
            text,
            @"^(?:第\s*)?(?<number>\d+(?:\s*[.．\-–—]\s*\d+)*)(?:\s*[章节篇部]|\s*[.．、:：)）\-–—]|\s+)",
            RegexOptions.CultureInvariant);
        if (!match.Success) return string.Empty;
        var number = Regex.Replace(
            match.Groups["number"].Value,
            @"\s*[.．\-–—]\s*",
            ".",
            RegexOptions.CultureInvariant);
        var parts = number
            .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToArray();
        if (parts.Length == 0 || parts.Length > Math.Max(1, outlineLevel))
            return string.Empty;
        return string.Join(".", parts);
    }

    private static (int ScopePosition, string Prefix) ResolveEquationNumberScope(
        int formulaPosition,
        EquationNumberFormat format,
        IReadOnlyList<HeadingNumberAnchor> anchors)
    {
        if (!format.UsesHeading) return (0, string.Empty);
        HeadingNumberAnchor? selected = null;
        for (var index = 0; index < anchors.Count; index++)
        {
            var anchor = anchors[index];
            if (anchor.Position > formulaPosition) break;
            selected = anchor;
        }
        var headingText = selected?.NumberText
            ?? string.Join(".", Enumerable.Repeat("0", format.HeadingLevel));
        return (
            selected?.Position ?? int.MinValue,
            headingText + format.Separator);
    }

    private static void UpdateNativeEquationCaptionNumber(
        Document document,
        string formulaId,
        string nativeSequenceName,
        int ordinal,
        string prefix)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Bookmark? captionBookmark = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Field? field = null;
        Range? code = null;
        Range? fieldResult = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefixRange = null;
        Range? refreshedNumberRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName) || !bookmarks.Exists(captionName)) return;
            numberBookmark = bookmarks[numberName];
            captionBookmark = bookmarks[captionName];
            numberRange = numberBookmark.Range;
            captionRange = captionBookmark.Range;
            field = FindNativeEquationFieldAtRange(document, numberRange, nativeSequenceName);
            if (field is null) return;

            code = field.Code;
            code.Text = $" SEQ {nativeSequenceName} \\r {ordinal} \\* ARABIC ";
            field.Update();
            fieldResult = field.Result;
            paragraphs = fieldResult.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            numberBookmark.Delete();
            Release(numberBookmark);
            numberBookmark = null;
            Release(code);
            code = field.Code;
            var fieldStart = Math.Max(paragraphRange.Start, code.Start - 1);
            prefixRange = document.Range(paragraphRange.Start, fieldStart);
            prefixRange.Text = prefix;

            Release(fieldResult);
            fieldResult = field.Result;
            refreshedNumberRange = document.Range(paragraphRange.Start, fieldResult.End);
            numberBookmark = bookmarks.Add(numberName, refreshedNumberRange);

            captionBookmark.Delete();
            Release(captionBookmark);
            captionBookmark = null;
            Release(captionRange);
            captionRange = paragraph.Range;
            captionBookmark = bookmarks.Add(captionName, captionRange);
            StyleNativeCaption(captionRange, refreshedNumberRange);
        }
        finally
        {
            Release(refreshedNumberRange);
            Release(prefixRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(captionRange);
            Release(numberRange);
            Release(captionBookmark);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static void UpdateFieldInBookmark(
        Document document,
        string bookmarkName,
        Func<string?, bool> predicate)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName)) return;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (predicate(code.Text))
                    {
                        field.Update();
                        NormalizeReferenceResult(field);
                    }
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AlignEquationNumberVertically(
        Range numberRange,
        float formulaHeightPoints,
        float formulaFontSizePoints)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = numberRange.Font;
            var numberFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);

            // The native caption target is deliberately white and one point.
            // Word propagates that appearance into REF results unless the
            // visible range is normalized after every field update. Normalize
            // the brackets and field result as one run so locale-specific body
            // fonts cannot put the digits and parentheses on different baselines.
            ApplyEquationNumberFont(
                font,
                numberFontSize,
                CalculateEquationNumberFontPosition(
                    formulaHeightPoints,
                    numberFontSize));
        }
        finally { Release(font); }
    }

    private static void ApplyEquationNumberFont(
        Microsoft.Office.Interop.Word.Font font,
        float size,
        int position)
    {
        font.Hidden = 0;
        font.Color = WdColor.wdColorAutomatic;
        font.Size = size;
        font.Position = position;
        try { font.Name = EquationNumberFontName; } catch { }
        try { font.NameAscii = EquationNumberFontName; } catch { }
        try { font.NameFarEast = EquationNumberFontName; } catch { }
        try { font.NameBi = EquationNumberFontName; } catch { }
        try { font.Bold = 0; } catch { }
        try { font.Italic = 0; } catch { }
        try { font.Superscript = 0; } catch { }
        try { font.Subscript = 0; } catch { }
        try { font.Scaling = 100; } catch { }
        try { font.Spacing = 0f; } catch { }
        try { font.Kerning = 0f; } catch { }
    }

    internal static IReadOnlyList<EquationReferenceTarget> GetEquationReferenceTargets(
        Document document)
    {
        Reconcile(document);
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var nativeFieldPositions = GetNativeEquationFieldPositions(document, nativeSequenceName);
        var nativeItems = document.GetCrossReferenceItems(WdCaptionLabelID.wdCaptionEquation) as Array;
        if (nativeItems is null || nativeItems.Length == 0)
            return Array.Empty<EquationReferenceTarget>();

        var targets = new List<EquationReferenceTarget>();
        void AddTarget(FormulaMetadata metadata, int position)
        {
            if (metadata.DisplayMode != "block" || !metadata.Numbered) return;
            if (!TryGetNativeCaptionInfo(
                    document,
                    metadata.FormulaId,
                    nativeSequenceName,
                    out var fieldPosition,
                    out var numberText))
                return;
            var nativeOrdinal = nativeFieldPositions.IndexOf(fieldPosition);
            if (nativeOrdinal < 0 || nativeOrdinal >= nativeItems.Length) return;
            var latex = string.Join(" ", metadata.Lines.Select(line => line.Latex))
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            if (latex.Length > 90) latex = latex.Substring(0, 87) + "…";
            targets.Add(new EquationReferenceTarget(
                metadata.FormulaId,
                nativeOrdinal + 1,
                numberText,
                latex,
                position));
        }

        var ommlFormulaIds = WordOmmlFormulaStore.FormulaIds(document);
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            for (var index = 1; index <= inlineShapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null) continue;
                    range = shape.Range;
                    AddTarget(metadata, range.Start);
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata is null) continue;
                    range = bookmark.Range;
                    AddTarget(metadata, range.Start);
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(inlineShapes); }
        return targets.OrderBy(target => target.Position).ToArray();
    }

    internal static void InsertEquationReference(
        Document document,
        Selection selection,
        EquationReferenceTarget target,
        EquationReferenceStyle style)
    {
        var prefix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "式（",
            EquationReferenceStyle.Parenthesized => "(",
            _ => string.Empty,
        };
        var suffix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "）",
            EquationReferenceStyle.Parenthesized => ")",
            _ => string.Empty,
        };

        Range? insertion = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        try
        {
            if (!string.IsNullOrEmpty(prefix)) selection.TypeText(prefix);
            insertion = selection.Range.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {NativeNumberBookmarkName(target.FormulaId)} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                insertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            NormalizeReferenceResult(field);
            result = field.Result;
            selection.SetRange(result.End, result.End);
            if (!string.IsNullOrEmpty(suffix)) selection.TypeText(suffix);
            UpdateNativeCrossReferences(document);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(fields);
            Release(insertion);
        }
    }

    internal static int FreezeFormulaCrossReferences(
        Document document,
        string formulaId)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        Fields? fields = null;
        var frozen = 0;
        try
        {
            fields = document.Fields;
            for (var index = fields.Count; index >= 1; index--)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var fieldCode = code.Text;
                    var matches = IsReferenceToBookmark(
                        fieldCode,
                        targetBookmarkName);
                    if (!matches
                        && TryResolveVisualTeXReferenceBookmark(
                            document,
                            fieldCode,
                            out var resolvedBookmark))
                    {
                        matches = string.Equals(
                            resolvedBookmark,
                            targetBookmarkName,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    if (!matches) continue;
                    NormalizeReferenceResult(field);
                    field.Unlink();
                    frozen++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        return frozen;
    }

    internal static int UpdateNativeCrossReferences(Document document)
    {
        Fields? fields = null;
        var updated = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsReferenceFieldCode(code.Text)) continue;
                    if (TryResolveVisualTeXReferenceBookmark(
                            document,
                            code.Text,
                            out var visualTeXBookmark)
                        && !IsReferenceToBookmark(code.Text, visualTeXBookmark))
                    {
                        code.Text = $" REF {visualTeXBookmark} \\h ";
                    }
                    // NormalizeReferenceResult performs the update itself after
                    // applying CHARFORMAT. Field.Type is not stable after Word
                    // rewrites a native cross-reference field, so the real REF
                    // field code is the authoritative discriminator.
                    NormalizeReferenceResult(field);
                    updated++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        return updated;
    }

    private static void NormalizeReferenceResult(Field field)
    {
        Range? code = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        try
        {
            var size = ResolveReferenceFontSize(field);
            code = field.Code;
            var codeText = code.Text ?? string.Empty;
            var normalizedCode = Regex.Replace(
                codeText,
                @"\\\*\s+MERGEFORMAT\b",
                string.Empty,
                RegexOptions.IgnoreCase);
            if (!Regex.IsMatch(
                    normalizedCode,
                    @"\\\*\s+CHARFORMAT\b",
                    RegexOptions.IgnoreCase))
            {
                normalizedCode = normalizedCode.TrimEnd() + " \\* CHARFORMAT ";
            }
            if (!string.Equals(codeText, normalizedCode, StringComparison.Ordinal))
                code.Text = normalizedCode;

            // CHARFORMAT makes Word use the field-code appearance instead of
            // copying the hidden one-point SEQ target appearance into the REF.
            codeFont = code.Font;
            ApplyEquationNumberFont(codeFont, size, position: 0);
            field.Update();

            result = field.Result;
            resultFont = result.Font;
            ApplyEquationNumberFont(resultFont, size, position: 0);
        }
        finally
        {
            Release(resultFont);
            Release(codeFont);
            Release(result);
            Release(code);
        }
    }

    private static float ResolveReferenceFontSize(Field field)
    {
        Range? result = null;
        Range? code = null;
        Range? paragraphRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Document? document = null;
        try
        {
            result = field.Result;
            var size = ReadUsableFontSize(result);
            if (size.HasValue) return size.Value;

            code = field.Code;
            size = ReadUsableFontSize(code);
            if (size.HasValue) return size.Value;

            paragraphs = result.Paragraphs;
            if (paragraphs.Count == 0) return 11f;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            document = result.Document;

            // Result.Start sits inside the field, so the immediately adjacent
            // character can still be a hidden field separator. Probe outside the
            // complete field code/result boundary, then the paragraph mark. This
            // recovers the surrounding body size even after Word copied the 1 pt
            // SEQ target appearance into both field code and result.
            var candidatePositions = new[]
            {
                code.Start - 2,
                code.Start - 1,
                result.End + 1,
                result.End + 2,
                paragraphRange.End - 1,
                paragraphRange.Start,
            };
            foreach (var position in candidatePositions.Distinct())
            {
                size = ReadUsableFontSizeAt(
                    document,
                    paragraphRange,
                    position);
                if (size.HasValue) return size.Value;
            }
            return 11f;
        }
        finally
        {
            Release(document);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(code);
            Release(result);
        }
    }

    private static float? ReadUsableFontSizeAt(
        Document document,
        Range paragraphRange,
        int position)
    {
        if (position < paragraphRange.Start || position >= paragraphRange.End)
            return null;
        Range? range = null;
        try
        {
            object start = position;
            object end = Math.Min(paragraphRange.End, position + 1);
            range = document.Range(ref start, ref end);
            return ReadUsableFontSize(range);
        }
        catch
        {
            return null;
        }
        finally { Release(range); }
    }

    private static float? ReadUsableFontSize(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            var size = font.Size;
            return float.IsNaN(size)
                || float.IsInfinity(size)
                || size <= 2f
                || size > 256f
                ? null
                : size;
        }
        catch
        {
            return null;
        }
        finally { Release(font); }
    }

    private static List<int> GetNativeEquationFieldPositions(
        Document document,
        string nativeSequenceName)
    {
        var positions = new List<int>();
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    positions.Add(result.Start);
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        positions.Sort();
        return positions;
    }

    private static bool TryGetNativeCaptionInfo(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out int position,
        out string numberText)
    {
        position = -1;
        numberText = string.Empty;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Field? field = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            var bookmarkName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            field = FindNativeEquationFieldAtRange(document, range, nativeSequenceName);
            if (field is null) return false;
            field.Update();
            result = field.Result;
            position = result.Start;
            numberText = (range.Text ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(numberText);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Field? FindNativeEquationFieldAtRange(
        Document document,
        Range targetRange,
        string nativeSequenceName)
    {
        // Word can invalidate a live Range RCW while the document-wide Fields
        // collection is materialized. Freeze the coordinates before enumerating
        // fields so later comparisons never re-enter a deleted COM proxy.
        var targetStart = targetRange.Start;
        var targetEnd = targetRange.End;
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    var overlaps = result.Start < targetEnd
                        && result.End > targetStart;
                    var sameCollapsedPosition = result.Start == targetStart
                        && result.End == targetEnd;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static string GetNativeEquationSequenceName(Document document)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        CaptionLabels? labels = null;
        CaptionLabel? label = null;
        try
        {
            application = document.Application;
            labels = application.CaptionLabels;
            label = labels[WdCaptionLabelID.wdCaptionEquation];
            var name = label.Name;
            if (string.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("Word built-in Equation caption label is unavailable.");
            return name;
        }
        finally
        {
            Release(label);
            Release(labels);
            Release(application);
        }
    }

    internal static string EquationBookmarkName(string formulaId) =>
        BookmarkName(EquationBookmarkPrefix, formulaId);

    internal static string NativeCaptionBookmarkName(string formulaId) =>
        BookmarkName(NativeCaptionBookmarkPrefix, formulaId);

    internal static string NativeNumberBookmarkName(string formulaId) =>
        BookmarkName(NativeNumberBookmarkPrefix, formulaId);

    private static string BookmarkName(string prefix, string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{prefix}{value:N}";
    }

    internal static bool TryFormulaIdFromEquationBookmark(
        string? bookmarkName,
        out string formulaId) =>
        TryFormulaIdFromBookmark(bookmarkName, EquationBookmarkPrefix, out formulaId);

    private static bool TryFormulaIdFromBookmark(
        string? bookmarkName,
        string prefix,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)) return false;
        var name = bookmarkName!;
        if (!name.StartsWith(prefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(name.Substring(prefix.Length), "N", out var value))
            return false;
        formulaId = value.ToString();
        return true;
    }

    internal static bool IsVisualTeXSequenceFieldCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"SEQ {LegacyEquationSequenceName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool IsNativeEquationSequenceFieldCode(
        string? code,
        string nativeSequenceName)
    {
        if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(nativeSequenceName))
            return false;
        return code!.IndexOf(
                   $"SEQ {nativeSequenceName}",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf(
                   $"SEQ \"{nativeSequenceName}\"",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool TryResolveVisualTeXReferenceBookmark(
        Document document,
        string? fieldCode,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(fieldCode)) return false;
        var match = Regex.Match(
            fieldCode!,
            @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
            RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        var targetName = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value
            : match.Groups["plain"].Value;
        if (targetName.StartsWith(NativeNumberBookmarkPrefix, StringComparison.Ordinal))
        {
            bookmarkName = targetName;
            return true;
        }

        Bookmarks? bookmarks = null;
        Bookmark? targetBookmark = null;
        Range? targetRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(targetName)) return false;
            targetBookmark = bookmarks[targetName];
            targetRange = targetBookmark.Range;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? candidateBookmark = null;
                Range? candidateRange = null;
                try
                {
                    candidateBookmark = bookmarks[index];
                    if (!candidateBookmark.Name.StartsWith(
                            NativeNumberBookmarkPrefix,
                            StringComparison.Ordinal))
                        continue;
                    candidateRange = candidateBookmark.Range;
                    var overlaps = candidateRange.Start < targetRange.End
                        && candidateRange.End > targetRange.Start;
                    var sameCollapsedPosition = candidateRange.Start == targetRange.Start
                        && candidateRange.End == targetRange.End;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    bookmarkName = candidateBookmark.Name;
                    return true;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidateBookmark);
                }
            }
            return false;
        }
        finally
        {
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
        }
    }

    private static bool IsReferenceFieldCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var trimmed = code!.TrimStart();
        return trimmed.StartsWith("REF ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReferenceToBookmark(string? code, string bookmarkName) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"REF {bookmarkName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static void DeleteBookmarkedRange(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            range.Delete();
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void DeleteBookmarkOnly(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
