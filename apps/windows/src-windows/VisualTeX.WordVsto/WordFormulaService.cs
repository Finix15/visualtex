using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    private const string InlineBaselineBookmarkPrefix = "VTBL_";
    private const string InlineMathGuard = "\u200B";
    private const string InlineBaselineSentinel = "\u2060";
    private const string LegacyInlineBaselineSentinel = "\u00A0";
    private const string BulkInlineFormulaPlaceholder = "\uE000";
    private const float ParagraphBeforeOleDisplaySpaceAfterPoints = 0f;
    private readonly Application _application;

    private sealed class WordViewState
    {
        internal int SelectionStart { get; set; }
        internal int SelectionEnd { get; set; }
        internal int? VerticalPercentScrolled { get; set; }
        internal int? HorizontalPercentScrolled { get; set; }
    }

    public WordFormulaService(Application application)
    {
        _application = application;
    }

    private static void TraceAcceptancePerformance(
        string operation,
        string stage,
        Stopwatch stopwatch,
        ref long checkpoint)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var elapsed = stopwatch.ElapsedMilliseconds;
        Console.WriteLine(
            $"    [perf] {operation}.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms total)");
        checkpoint = elapsed;
    }

    public OfficeSelection ReadSelection() => ReadSelection(null);

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? inlineShapes = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        Range? ommlEquationRange = null;
        var ownsSelection = providedSelection is null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = providedSelection ?? _application.Selection;
            range = selection.Range;
            inlineShapes = range.InlineShapes;
            FormulaMetadata? metadata = null;
            string? objectMode = null;
            if (inlineShapes.Count == 1)
            {
                shape = inlineShapes[1];
                metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                {
                    metadata.FontSizePt = FormulaFontSize.InferOleFontSize(
                        shape.Width,
                        shape.Height,
                        metadata);
                    objectMode = WordFormulaMetadataReader.IsNativeOle(shape)
                        ? FormulaOleContract.NativeOleMode
                        : FormulaOleContract.CrossPlatformPictureMode;
                }
            }
            if (metadata is null)
            {
                ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, range);
                if (ommlBookmark is not null)
                {
                    metadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
                    if (metadata is not null)
                    {
                        metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                            document,
                            ommlBookmark,
                            metadata);
                        metadata.FontSizePt = ReadOmmlFontSize(ommlBookmark, metadata);
                        objectMode = FormulaOleContract.WordOmmlMode;
                        // A double-click usually supplies only a collapsed caret
                        // or a small subrange inside the OMath. Word clips an
                        // OMath.Range obtained from such a probe, so carrying the
                        // raw selection into replacement can splice a new formula
                        // into the middle of the old one. Persist the bookmark-
                        // resolved complete equation range as the edit hint.
                        ommlEquationRange = WordOmmlFormulaStore.GetEquationRange(
                            ommlBookmark);
                    }
                }
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                // OLE keeps the exact source range as a fast edit hint. OMML
                // must carry the complete bookmark-resolved equation range;
                // Word clips ranges obtained from a caret inside an OMath.
                ObjectId = RangeReference(ommlEquationRange ?? range),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
                ObjectMode = objectMode,
            };
        }
        finally
        {
            Release(ommlEquationRange);
            Release(ommlBookmark);
            Release(shape);
            Release(inlineShapes);
            Release(range);
            if (ownsSelection) Release(selection);
            Release(document);
        }
    }

    public string ReadActiveDocumentId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return DocumentIdentity(document);
        }
        finally { Release(document); }
    }

    public bool IsSelectedNativeOle()
    {
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? format = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return false;
            shape = shapes[1];
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;
            format = shape.OLEFormat;
            return string.Equals(
                format.ProgID,
                FormulaOleContract.ProgId,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
        }
    }

    public void NormalizeTypingCaretAfterInlineFormula(Selection selection)
    {
        if (selection is null) return;
        Range? caret = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            caret = selection.Range;
            if (caret.Start != caret.End) return;
            paragraphs = caret.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (TryMoveCaretPastInlineBoundary(selection, caret.Start, shape))
                    return;
            }
        }
        catch
        {
            // A selection-change repair must never interrupt ordinary Word use.
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(caret);
        }
    }

    private bool TryMoveCaretPastInlineBoundary(
        Selection selection,
        int caretPosition,
        InlineShape shape)
    {
        if (!WordFormulaMetadataReader.IsNativeOle(shape)) return false;
        var metadata = WordFormulaMetadataReader.TryRead(shape);
        if (metadata is null
            || !string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return false;

        Range? formulaRange = null;
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            formulaRange = shape.Range;
            if (caretPosition < formulaRange.End) return false;
            document = formulaRange.Document;
            bookmarks = document.Bookmarks;
            var bookmarkName = InlineBaselineBookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(bookmarkName))
            {
                bookmark = bookmarks[bookmarkName];
                sentinel = bookmark.Range;
                if (IsUsableInlineBaselineSentinel(sentinel, formulaRange)
                    && caretPosition <= sentinel.End)
                {
                    ConfigureInlineBaselineSentinel(sentinel);
                    if (selection.Start != sentinel.End || selection.End != sentinel.End)
                        selection.SetRange(sentinel.End, sentinel.End);
                    font = selection.Font;
                    font.Position = 0;
                    font.Hidden = 0;
                    return true;
                }
            }

            if (caretPosition != formulaRange.End) return false;
            var target = EnsureInlineBaselineSentinel(formulaRange, metadata.FormulaId);
            selection.SetRange(target, target);
            font = selection.Font;
            font.Position = 0;
            font.Hidden = 0;
            return true;
        }
        finally
        {
            Release(font);
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
            Release(document);
            Release(formulaRange);
        }
    }

    private static bool TryResolveWordFontSize(float value, out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (float.IsNaN(value)
            || float.IsInfinity(value)
            || value < FormulaFontSize.MinimumPt
            || value > FormulaFontSize.MaximumPt)
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    public float ReadCurrentTypingFontSize()
    {
        Selection? selection = null;
        Range? selectionRange = null;
        Range? probeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            selection = _application.Selection;
            selectionRange = selection.Range;
            font = selection.Font;
            if (TryResolveWordFontSize(font.Size, out var selectedSize))
                return selectedSize;
            Release(font);
            font = null;

            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            // A collapsed insertion point can report Word's mixed-size sentinel.
            // Prefer the character immediately before the caret in the current
            // paragraph, then the next character, before falling back to the
            // paragraph run as a whole. This makes a new formula inherit the
            // actual surrounding body text instead of an arbitrary global size.
            if (selectionRange.Start > (paragraphRange?.Start ?? 0))
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start - 1, selectionRange.Start);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var previousSize))
                    return previousSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            var paragraphEnd = Math.Max(
                paragraphRange?.Start ?? selectionRange.Start,
                (paragraphRange?.End ?? selectionRange.End) - 1);
            if (selectionRange.Start < paragraphEnd)
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start, selectionRange.Start + 1);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var nextSize))
                    return nextSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        catch
        {
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probeRange);
            Release(selectionRange);
            Release(selection);
        }
    }

    public float? GetSelectedFormulaFontSize()
    {
        var selected = ReadSelection();
        return selected.Metadata is null
            ? null
            : FormulaFontSize.ResolveSemanticFontSize(selected.Metadata);
    }

    public float SetSelectedFormulaFontSize(double requestedFontSizePt)
    {
        var selected = ReadSelection();
        if (selected.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
            throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

        var target = FormulaFontSize.Normalize(requestedFontSizePt);
        Document? document = null;
        InlineShape? shape = null;
        Bookmark? bookmark = null;
        Range? equationRange = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Set Formula Font Size");
            var metadata = selected.Metadata;
            var sourceSemanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);

            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, selected.FormulaId!)
                    ?? throw new InvalidOperationException("The selected Word OMML formula no longer exists.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                metadata.FontSizePt = target;
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                var alignInline = ShouldAlignInline(equationRange, metadata);
                if (alignInline) metadata.DisplayMode = "inline";
                ApplyOmmlFontSize(equationRange, target);
                WordOmmlNativeSource.StampFingerprint(metadata, equationRange);
                WordOmmlFormulaStore.Save(document, metadata);
                if (alignInline)
                    RestoreTypingBaselineAfter(bookmark);
                else
                    TryReconcileOmml(document, bookmark, equationRange, metadata);
                return target;
            }

            shape = FindByFormulaId(document, selected.FormulaId!)
                ?? throw new InvalidOperationException("The selected Word formula no longer exists.");
            var alignOleInline = ShouldAlignInline(shape, metadata);
            if (alignOleInline) metadata.DisplayMode = "inline";
            var existingFontPosition = ReadDefinedShapeFontPosition(shape);
            metadata.FontSizePt = target;
            var size = FormulaFontSize.OleSizeAt(metadata, target);
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Width = size.Width;
            shape.Height = size.Height;
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            if (!WordFormulaMetadataReader.IsNativeOle(shape))
            {
                var encoded = FormulaMetadataCodec.Encode(metadata);
                shape.Title = encoded;
                shape.AlternativeText = encoded;
            }
            if (alignOleInline)
            {
                ApplyInlineBaseline(
                    shape,
                    shape.Height,
                    (float)(metadata.RenderHeightPx ?? 0),
                    metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                    existingFontPosition,
                    sourceSemanticFontSize,
                    target);
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                ResetShapeFontPosition(shape);
                TryReconcileShape(document, shape, metadata);
            }
            return target;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(equationRange);
            Release(bookmark);
            Release(shape);
            Release(document);
        }
    }

    public string DeleteSelectedFormula()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        Range? ommlRange = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Delete Formula");
            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word OMML formula no longer exists.");
                ommlRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
                ommlBookmark.Delete();
                ommlRange.Delete();
                WordOmmlFormulaStore.Delete(document, requiredFormulaId);
            }
            else
            {
                shape = FindByFormulaId(document, requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word formula no longer exists.");
                RemoveInlineBaselineSentinel(document, requiredFormulaId);
                shape.Delete();
            }
            WordEquationNumbering.TryReconcile(document);
            return requiredFormulaId;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(ommlRange);
            Release(ommlBookmark);
            Release(shape);
            Release(document);
        }
    }

    public int UpdateEquationNumbers()
    {
        Document? document = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Update Equation Numbers");
            return WordEquationNumbering.Reconcile(document);
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(document);
        }
    }

    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? oldShape = null;
        OLEFormat? format = null;
        object? oleObject = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        string? pngPath = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Export OLE Formula As Picture");
            oldShape = FindByFormulaId(document, requiredFormulaId)
                ?? throw new InvalidOperationException("The selected Word formula no longer exists.");
            var metadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? throw new InvalidDataException("The selected formula metadata is invalid.");
            format = oldShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected formula is already a picture.");
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            pngPath = OlePngPreviewExtractor.MaterializePng(oleObject, requiredFormulaId);

            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                pngPath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                oldWidth,
                oldHeight,
                pngPath,
                (float)(metadata.RenderHeightPx ?? 0),
                metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                metadata.DisplayMode == "inline");
            oldShape.Delete();
            TryReconcileShape(document, replacement, metadata);
            return requiredFormulaId;
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            if (pngPath is not null)
            {
                try { File.Delete(pngPath); } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oleObject);
            Release(format);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Inline Formula"
                    : "VisualTeX Insert Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            object link = false;
            object save = true;
            object rangeObject;
            if (session.DisplayMode == "inline")
            {
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            else
            {
                CompactParagraphBeforeOleDisplayFormula(document, insertion);
                insertion.InsertParagraphAfter();
                insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                paragraph = document.Paragraphs.Add(insertion);
                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraphRange = paragraph.Range;
                rangeObject = paragraphRange;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                TryReconcileShape(document, shape, metadata);
                if (session.Numbered)
                    MoveCaretToNormalTypingParagraphAfterNumberedDisplay(
                        document,
                        metadata.FormulaId);
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        Table? numberedTable = null;
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Native OLE Inline Formula"
                    : "VisualTeX Insert Native OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (session.DisplayMode == "inline")
            {
                shape = AddOleObject(document, insertion);
            }
            else
            {
                CompactParagraphBeforeOleDisplayFormula(document, insertion);
                if (session.Numbered)
                {
                    var tableInsertion = CreateNumberedDisplayTable(
                        document,
                        insertion,
                        out numberedTable);
                    Release(insertion);
                    insertion = tableInsertion;
                    shape = AddOleObject(document, insertion);
                }
                else
                {
                    insertion.InsertParagraphAfter();
                    insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                    paragraph = document.Paragraphs.Add(insertion);
                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphRange = paragraph.Range;
                    shape = AddOleObject(document, paragraphRange);
                }
            }
            InitializeOle(shape, metadata, emfPath, pngPath);
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                TryReconcileShape(document, shape, metadata);
                if (session.Numbered)
                    MoveCaretToNormalTypingParagraphAfterNumberedDisplay(
                        document,
                        metadata.FormulaId);
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            if (numberedTable is not null) TryDelete(numberedTable);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(numberedTable);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? equationRange = null;
        string sourceFingerprint = string.Empty;
        Bookmark? bookmark = null;
        Table? numberedTable = null;
        UndoRecord? undoRecord = null;
        var metadataSaved = false;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Word OMML Inline Formula"
                    : "VisualTeX Insert Word OMML Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (session.DisplayMode == "inline")
            {
                var placeholder = PrepareInlineBaselineSentinelBeforeInsert(
                    document,
                    insertion,
                    metadata.FormulaId);
                Release(insertion);
                insertion = placeholder;
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl,
                    display: false,
                    sourceFingerprint: out sourceFingerprint,
                    replaceTarget: true);
            }
            else
            {
                if (session.Numbered)
                {
                    var tableInsertion = CreateNumberedDisplayTable(
                        document,
                        insertion,
                        out numberedTable);
                    Release(insertion);
                    insertion = tableInsertion;
                }
                else
                {
                    insertion.InsertParagraphAfter();
                    insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                    paragraph = document.Paragraphs.Add(insertion);
                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphRange = paragraph.Range;
                    insertion.SetRange(paragraphRange.Start, paragraphRange.Start);
                }
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl,
                    display: true,
                    sourceFingerprint: out sourceFingerprint);
            }

            ApplyOmmlFontSize(equationRange, session.FontSizePt);
            metadata.NativeOmmlFingerprint = sourceFingerprint;
            bookmark = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
            WordOmmlFormulaStore.Save(document, metadata);
            metadataSaved = true;
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(bookmark);
            }
            else
            {
                TryReconcileOmml(document, bookmark!, equationRange, metadata);
                if (session.Numbered)
                    MoveCaretToNormalTypingParagraphAfterNumberedDisplay(
                        document,
                        metadata.FormulaId);
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(bookmark, deleteContents: true);
            if (bookmark is null) TryDelete(equationRange);
            if (numberedTable is not null) TryDelete(numberedTable);
            if (metadataSaved && document is not null)
            {
                try { WordOmmlFormulaStore.Delete(document, metadata.FormulaId); } catch { }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(bookmark);
            Release(equationRange);
            Release(numberedTable);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public WordBulkInsertResult InsertBulkDocument(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        Document? document = null;
        Selection? selection = null;
        Range? rollbackRange = null;
        UndoRecord? undoRecord = null;
        var insertedFormulaIds = new List<string>();
        var insertionStart = -1;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, expectedDocumentId);
            selection = _application.Selection;
            Range? sourceRange = null;
            try
            {
                sourceRange = ResolveSourceRange(document, sourceObjectId, selection);
                selection.SetRange(sourceRange.Start, sourceRange.End);
            }
            finally { Release(sourceRange); }
            if (selection.Range.Start != selection.Range.End)
                selection.Text = string.Empty;
            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            insertionStart = selection.Start;
            undoRecord = BeginUndoRecord("VisualTeX 批量导入 LaTeX / Markdown");

            for (var blockIndex = 0; blockIndex < source.Blocks.Count; blockIndex++)
            {
                var block = source.Blocks[blockIndex];
                var nextKind = blockIndex + 1 < source.Blocks.Count
                    ? source.Blocks[blockIndex + 1].Kind
                    : (WordBulkBlockKind?)null;
                if (block.Kind == WordBulkBlockKind.DisplayFormula)
                {
                    var formulaRun = block.Runs.Single(run => run.IsFormula);
                    if (!prepared.TryGetValue(formulaRun.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行间公式 {formulaRun.Id} 的渲染结果。");
                    InsertPreparedFormula(document, selection, formula, display: true);
                    insertedFormulaIds.Add(formula.Session.FormulaId);
                    continue;
                }

                EnsureWritableParagraph(selection);
                var paragraphStart = selection.Start;
                var pendingInlineFormulas = new List<(int Start, PreparedWordBulkFormula Formula)>();
                foreach (var run in block.Runs)
                {
                    if (!run.IsFormula)
                    {
                        InsertNativeTextRun(document, selection, run);
                        continue;
                    }
                    if (!prepared.TryGetValue(run.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行内公式 {run.Id} 的渲染结果。");

                    // Write the complete native paragraph before materializing
                    // inline formulas. Word keeps a caret collapsed at an OMML
                    // range end inside the math zone, so typing the following
                    // text immediately would absorb that text into <m:oMath>.
                    // Replacing one-character placeholders from right to left
                    // also ensures paragraph/list formatting is applied before
                    // OLE baseline offsets are calculated and persisted.
                    var placeholderStart = selection.Start;
                    selection.TypeText(BulkInlineFormulaPlaceholder);
                    pendingInlineFormulas.Add((placeholderStart, formula));
                }
                selection.TypeParagraph();
                var paragraphEnd = selection.Start;
                ApplyBulkParagraphFormatting(
                    document,
                    paragraphStart,
                    paragraphEnd,
                    block);

                for (var formulaIndex = pendingInlineFormulas.Count - 1;
                     formulaIndex >= 0;
                     formulaIndex--)
                {
                    var pending = pendingInlineFormulas[formulaIndex];
                    selection.SetRange(pending.Start, pending.Start + BulkInlineFormulaPlaceholder.Length);
                    selection.Text = string.Empty;
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);
                    InsertPreparedFormula(document, selection, pending.Formula, display: false);
                    insertedFormulaIds.Add(pending.Formula.Session.FormulaId);
                }

                MoveSelectionAfterBulkParagraph(document, selection, paragraphStart);
                ResetNextParagraphFormatting(selection, block.Kind, nextKind);
            }

            return new WordBulkInsertResult
            {
                BlockCount = source.Blocks.Count,
                FormulaCount = insertedFormulaIds.Count,
                FormulaIds = insertedFormulaIds,
            };
        }
        catch
        {
            if (document is not null)
            {
                foreach (var formulaId in insertedFormulaIds)
                {
                    try { RemoveInlineBaselineSentinel(document, formulaId); } catch { }
                    try { WordOmmlFormulaStore.Delete(document, formulaId); } catch { }
                }
                if (insertionStart >= 0)
                {
                    try
                    {
                        var contentEnd = document.Content.End;
                        var rollbackEnd = selection is null
                            ? insertionStart
                            : Math.Min(Math.Max(insertionStart, selection.Start), contentEnd);
                        rollbackRange = document.Range(insertionStart, rollbackEnd);
                        rollbackRange.Delete();
                    }
                    catch { }
                }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(rollbackRange);
            Release(selection);
            Release(document);
        }
    }

    private void InsertPreparedFormula(
        Document document,
        Selection selection,
        PreparedWordBulkFormula prepared,
        bool display)
    {
        var session = prepared.Session;
        session.DisplayMode = display ? "block" : "inline";
        session.Numbered = false;
        var metadata = session.ToMetadata();
        metadata.Validate();
        var nativeOmml = string.Equals(
            session.ObjectMode,
            FormulaOleContract.WordOmmlMode,
            StringComparison.Ordinal);
        Range? insertion = null;
        Range? equationRange = null;
        string sourceFingerprint = string.Empty;
        Bookmark? bookmark = null;
        InlineShape? shape = null;
        try
        {
            if (display)
            {
                if (!nativeOmml)
                {
                    Range? spacingAnchor = null;
                    try
                    {
                        spacingAnchor = selection.Range.Duplicate;
                        CompactParagraphBeforeOleDisplayFormula(document, spacingAnchor);
                    }
                    finally { Release(spacingAnchor); }
                }
                EnsureBlankDisplayParagraph(selection, preserveNativeOmmlSpacing: nativeOmml);
            }
            insertion = selection.Range.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (nativeOmml)
            {
                var mathMl = prepared.MathMl;
                if (string.IsNullOrWhiteSpace(mathMl))
                    throw new InvalidDataException(
                        $"公式 {metadata.FormulaId} 没有可用的 MathML。" );
                if (!display)
                {
                    var placeholder = PrepareInlineBaselineSentinelBeforeInsert(
                        document,
                        insertion,
                        metadata.FormulaId);
                    Release(insertion);
                    insertion = placeholder;
                }
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl!,
                    display,
                    sourceFingerprint: out sourceFingerprint,
                    replaceTarget: !display);
                ApplyOmmlFontSize(equationRange, session.FontSizePt);
                metadata.NativeOmmlFingerprint = sourceFingerprint;
                bookmark = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
                WordOmmlFormulaStore.Save(document, metadata);
                if (display)
                {
                    TryReconcileOmml(document, bookmark, equationRange, metadata);
                    MoveSelectionAfterDisplayFormula(selection, equationRange);
                }
                else
                {
                    RestoreTypingBaselineAfter(bookmark);
                }
                return;
            }

            if (!string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"批量导入不支持公式对象格式 {session.ObjectMode}。" );
            if (string.IsNullOrWhiteSpace(prepared.PngPath)
                || string.IsNullOrWhiteSpace(prepared.EmfPath))
                throw new InvalidDataException(
                    $"公式 {metadata.FormulaId} 没有可用的 OLE 预览。" );
            shape = AddOleObject(document, insertion);
            InitializeOle(shape, metadata, prepared.EmfPath!, prepared.PngPath!);
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                prepared.PngPath!,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                !display);
            if (display)
            {
                TryReconcileShape(document, shape, metadata);
                Range? shapeRange = null;
                try
                {
                    shapeRange = shape.Range;
                    MoveSelectionAfterDisplayFormula(selection, shapeRange);
                }
                finally { Release(shapeRange); }
            }
            else
            {
                RestoreTypingBaselineAfter(shape);
            }
        }
        catch
        {
            if (bookmark is not null) TryDelete(bookmark, deleteContents: true);
            else TryDelete(equationRange);
            TryDelete(shape);
            try { WordOmmlFormulaStore.Delete(document, metadata.FormulaId); } catch { }
            throw;
        }
        finally
        {
            Release(shape);
            Release(bookmark);
            Release(equationRange);
            Release(insertion);
        }
    }

    private static void MoveSelectionAfterBulkParagraph(
        Document document,
        Selection selection,
        int paragraphStart)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            anchor = document.Range(paragraphStart, paragraphStart);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Word 未能定位批量导入段落。");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            selection.SetRange(paragraphRange.End, paragraphRange.End);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static void EnsureWritableParagraph(Selection selection)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            if (!ContainsVisibleBodyText(range.Text))
                return;
            if (selection.Start >= range.End - 1)
                selection.TypeParagraph();
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void CompactParagraphBeforeOleDisplayFormula(
        Document document,
        Range insertion)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Min(Math.Max(insertion.Start, contentStart), contentEnd);
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A caret at the start of an empty paragraph belongs to that empty
            // paragraph, while the display formula will still be visually tied
            // to the preceding prose. Resolve the preceding paragraph mark in
            // that case so the local spacing adjustment is applied to the text
            // the reader actually sees above the equation.
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && position > contentStart)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                Release(anchor);
                anchor = null;

                anchor = document.Range(position - 1, position - 1);
                paragraphs = anchor.Paragraphs;
                if (paragraphs.Count == 0) return;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }

            if (!ContainsVisibleBodyText(paragraphRange.Text)) return;
            format = paragraph.Format;
            format.SpaceAfter = ParagraphBeforeOleDisplaySpaceAfterPoints;
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static void EnsureBlankDisplayParagraph(
        Selection selection,
        bool preserveNativeOmmlSpacing)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                range = paragraph.Range;
                if (ContainsVisibleBodyText(range.Text))
                    selection.TypeParagraph();
            }
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                selection.ParagraphFormat.SpaceBefore = 0;
                selection.ParagraphFormat.SpaceAfter = 0;
                selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void MoveSelectionAfterDisplayFormula(
        Selection selection,
        Range formulaRange)
    {
        selection.SetRange(formulaRange.End, formulaRange.End);
        selection.TypeParagraph();
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
    }

    private static void InsertNativeTextRun(
        Document document,
        Selection selection,
        WordBulkRun run)
    {
        if (string.IsNullOrEmpty(run.Text)) return;
        var start = selection.Start;
        selection.TypeText(run.Text);
        Range? inserted = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            inserted = document.Range(start, selection.Start);
            font = inserted.Font;
            font.Bold = run.Bold ? 1 : 0;
            font.Italic = run.Italic ? 1 : 0;
            font.StrikeThrough = run.Strike ? 1 : 0;
            font.Underline = run.Underline
                ? WdUnderline.wdUnderlineSingle
                : WdUnderline.wdUnderlineNone;
            if (run.Code)
            {
                font.Name = "Consolas";
                try { font.NameAscii = "Consolas"; } catch { }
                try { font.NameFarEast = "Microsoft YaHei UI"; } catch { }
            }
        }
        finally
        {
            Release(font);
            Release(inserted);
        }
    }

    private static void ApplyBulkParagraphFormatting(
        Document document,
        int start,
        int end,
        WordBulkBlock block)
    {
        if (end < start) return;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = document.Range(start, end);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            switch (block.Kind)
            {
                case WordBulkBlockKind.Heading:
                    object heading = block.Level switch
                    {
                        <= 1 => WdBuiltinStyle.wdStyleHeading1,
                        2 => WdBuiltinStyle.wdStyleHeading2,
                        3 => WdBuiltinStyle.wdStyleHeading3,
                        _ => WdBuiltinStyle.wdStyleHeading4,
                    };
                    try { range.set_Style(ref heading); } catch { }
                    break;
                case WordBulkBlockKind.Bullet:
                    listFormat = range.ListFormat;
                    listFormat.ApplyBulletDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Numbered:
                    listFormat = range.ListFormat;
                    listFormat.ApplyNumberDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Quote:
                    paragraph.LeftIndent = 18f;
                    paragraph.RightIndent = 9f;
                    font = range.Font;
                    font.Italic = 1;
                    break;
                case WordBulkBlockKind.Code:
                    font = range.Font;
                    font.Name = "Consolas";
                    try { font.NameAscii = "Consolas"; } catch { }
                    paragraph.LeftIndent = 18f;
                    paragraph.SpaceBefore = 3f;
                    paragraph.SpaceAfter = 3f;
                    break;
            }

        }
        finally
        {
            Release(font);
            Release(listFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    private static void ResetNextParagraphFormatting(
        Selection selection,
        WordBulkBlockKind current,
        WordBulkBlockKind? next)
    {
        var continuingList =
            current == WordBulkBlockKind.Bullet && next == WordBulkBlockKind.Bullet
            || current == WordBulkBlockKind.Numbered && next == WordBulkBlockKind.Numbered;
        if (continuingList) return;
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.RightIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
    }

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        Range? rollbackEquationRange = null;
        Bookmark? rollbackBookmark = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalMetadata = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var oldStart = -1;
        var removedOmml = false;
        var performanceWatch = Stopwatch.StartNew();
        long performanceCheckpoint = 0;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Native OLE Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            oldShape = FindByFormulaId(
                document,
                session.FormulaId,
                session.SourceObjectId);
            TraceAcceptancePerformance(
                "ReplaceOle",
                "locate-target",
                performanceWatch,
                ref performanceCheckpoint);
            float oldWidth;
            float oldHeight;
            if (oldShape is not null)
            {
                oldWidth = oldShape.Width;
                oldHeight = oldShape.Height;
                originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                    ?? session.OriginalMetadata;
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(document, session.FormulaId)
                    ?? throw new InvalidOperationException(
                        "The target Word formula no longer exists.");
                originalMetadata = WordOmmlFormulaStore.TryRead(document, oldBookmark)
                    ?? session.OriginalMetadata;
                oldWidth = (float)Math.Max(
                    12,
                    (originalMetadata?.RenderWidthPx ?? session.ExportResult?.Width ?? 200) * 0.75);
                oldHeight = (float)Math.Max(
                    12,
                    (originalMetadata?.RenderHeightPx ?? session.ExportResult?.Height ?? 60) * 0.75);
            }
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                originalFontSizePt: originalMetadata?.FontSizePt,
                originalRenderFontSizePt: originalMetadata?.RenderFontSizePt);

            // Reusing an inline OLE object preserves its previous COM extent.
            // When the edited formula becomes wider or taller, Word can then
            // scale the new preview back into the old canvas and make every
            // glyph look smaller. Recreate inline objects so the OLE server is
            // initialized with the new natural extent; block objects may still
            // update in place because their outer layout is intentionally
            // controlled by the host paragraph/table.
            if (oldShape is not null
                && session.DisplayMode != "inline"
                && TryUpdateOle(oldShape, metadata, emfPath, pngPath))
            {
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "update-native-object",
                    performanceWatch,
                    ref performanceCheckpoint);
                Configure(
                    oldShape,
                    metadata,
                    editedSize.Width,
                    editedSize.Height,
                    pngPath,
                    session.ExportResult?.Height ?? 0,
                    session.ExportResult?.Baseline,
                    session.DisplayMode == "inline");
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "configure",
                    performanceWatch,
                    ref performanceCheckpoint);
                if (session.DisplayMode == "inline")
                    RestoreTypingBaselineAfter(oldShape);
                else
                    TryReconcileShape(document, oldShape, metadata);
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "reconcile",
                    performanceWatch,
                    ref performanceCheckpoint);
                finalSelection = oldShape.Range.Duplicate;
                return Result(session, document);
            }

            oldRange = oldShape is not null
                ? oldShape.Range
                : WordOmmlFormulaStore.GetEquationRange(oldBookmark!);
            oldStart = oldRange.Start;
            if (oldShape is not null)
            {
                insertion = oldRange.Duplicate;
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                // Remove the native equation before creating the OLE object.
                // Inserting at OMath.End first allows Word to expand the live
                // math container around the OLE object, leaving the large OMML
                // selection frame and shifting the replacement horizontally.
                oldBookmark!.Delete();
                oldRange.Delete();
                WordOmmlFormulaStore.Delete(document, session.FormulaId);
                removedOmml = true;
                object insertionStart = oldStart;
                object insertionEnd = oldStart;
                insertion = document.Range(ref insertionStart, ref insertionEnd);
            }
            replacement = AddOleObject(document, insertion);
            InitializeOle(replacement, metadata, emfPath, pngPath);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (oldShape is not null)
                oldShape.Delete();
            if (session.DisplayMode == "block" && session.Numbered)
                NormalizeNumberedDisplayCell(replacement);
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(replacement);
                // OMML -> OLE conversion must leave the insertion point in the
                // ordinary zero-position text run after the formula. Selecting
                // the replacement again in finally restores the shape's negative
                // baseline onto Word's typing caret and contaminates following
                // prose. Existing OLE edits keep the historical object selection.
                finalSelection = oldShape is null
                    ? DuplicateCurrentSelectionRange()
                    : replacement.Range.Duplicate;
            }
            else
            {
                TryReconcileShape(document, replacement, metadata);
                finalSelection = replacement.Range.Duplicate;
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            if (removedOmml
                && document is not null
                && originalMetadata is not null
                && oldStart >= 0
                && !string.IsNullOrWhiteSpace(session.ExportResult?.MathMl))
            {
                try
                {
                    object restoreStart = oldStart;
                    object restoreEnd = oldStart;
                    string rollbackFingerprint = string.Empty;
                    var restoreInsertion = document.Range(ref restoreStart, ref restoreEnd);
                    try
                    {
                        rollbackEquationRange = WordOmmlConverter.Insert(
                            _application,
                            document,
                            restoreInsertion,
                            session.ExportResult!.MathMl!,
                            session.DisplayMode == "block",
                            sourceFingerprint: out rollbackFingerprint,
                            includeLeadingTab: false);
                    }
                    finally { Release(restoreInsertion); }
                    ApplyOmmlFontSize(
                        rollbackEquationRange,
                        FormulaFontSize.ResolveSemanticFontSize(originalMetadata));
                    originalMetadata.NativeOmmlFingerprint = rollbackFingerprint;
                    rollbackBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        rollbackEquationRange,
                        originalMetadata);
                    WordOmmlFormulaStore.Save(document, originalMetadata);
                    WordEquationNumbering.TryReconcile(document);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(rollbackBookmark);
            Release(rollbackEquationRange);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult ReplaceOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        Range? equationRange = null;
        string sourceFingerprint = string.Empty;
        Bookmark? replacement = null;
        Table? numberedTable = null;
        Paragraph? replacementParagraph = null;
        Range? replacementParagraphRange = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalOmmlMetadata = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var metadataSaved = false;
        var oldBookmarkRemoved = false;
        var oldNumberingArtifactsRemoved = false;
        var performanceWatch = Stopwatch.StartNew();
        long performanceCheckpoint = 0;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Word OMML Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            var sourceWasOmml = !string.IsNullOrWhiteSpace(
                session.OriginalMetadata?.NativeOmmlFingerprint);
            oldShape = sourceWasOmml
                ? null
                : FindByFormulaId(
                    document,
                    session.FormulaId,
                    session.SourceObjectId);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "locate-target",
                performanceWatch,
                ref performanceCheckpoint);
            if (oldShape is not null)
            {
                // Remove old equation-number bookmarks before inserting the
                // adjacent replacement. Word expands a trailing bookmark when
                // content is inserted at its edge; deleting that old bookmark
                // during reconciliation can otherwise delete the new table.
                WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                    document,
                    session.FormulaId);
                oldNumberingArtifactsRemoved = true;
                oldRange = oldShape.Range;
                insertion = oldRange.Duplicate;
                // Insert immediately before the source OLE in its existing
                // paragraph/cell. Creating another paragraph or numbered table
                // here nests layout containers during OLE -> OMML -> OLE
                // round-trips and leaves visible empty paragraph marks behind.
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(document, session.FormulaId)
                    ?? throw new InvalidOperationException(
                        "The target Word formula no longer exists.");
                originalOmmlMetadata = WordOmmlFormulaStore.TryRead(document, oldBookmark);
                // FormulaId/bookmark is the durable OMML identity. Never replace
                // an OMath range reconstructed from the editor-opening selection:
                // Word may clip OMath.Range to that caret/partial probe, which
                // leaves the old equation around the inserted replacement and
                // corrupts the paragraph layout. Resolve the complete equation
                // from its collapsed bookmark immediately before committing.
                oldRange = WordOmmlFormulaStore.GetEquationRange(oldBookmark);
                insertion = oldRange.Duplicate;
            }
            if (session.DisplayMode == "inline")
                PrepareInlineBaselineSentinelAfterFormula(
                    document,
                    oldRange,
                    metadata.FormulaId);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "resolve-source",
                performanceWatch,
                ref performanceCheckpoint);

            equationRange = WordOmmlConverter.Insert(
                _application,
                document,
                insertion,
                mathMl,
                session.DisplayMode == "block",
                sourceFingerprint: out sourceFingerprint,
                replaceTarget: oldShape is null);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "insert-native-omml",
                performanceWatch,
                ref performanceCheckpoint);
            ValidateInsertedOmml(equationRange);
            ApplyOmmlFontSize(equationRange, session.FontSizePt);
            if (oldBookmark is not null)
            {
                oldBookmark.Delete();
                oldBookmarkRemoved = true;
            }
            metadata.NativeOmmlFingerprint = sourceFingerprint;
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "stamp-fingerprint",
                performanceWatch,
                ref performanceCheckpoint);
            replacement = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
            WordOmmlFormulaStore.Save(document, metadata);
            metadataSaved = true;
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "save-metadata",
                performanceWatch,
                ref performanceCheckpoint);

            // Keep the source OLE until replacement and metadata are valid.
            if (oldShape is not null)
            {
                oldShape.Delete();
                RemoveInlineBaselineSentinel(document, session.FormulaId);
            }
            if (session.DisplayMode == "block" && session.Numbered)
            {
                // Turning an OMath into display form while its source OLE is
                // still present makes Word insert manual line-break runs on
                // both sides. Once the OLE is deleted those hidden breaks stay
                // in the formula cell, so the cell is centered but the formula
                // is not. Remove everything outside the replacement equation,
                // then recreate its collapsed anchor at the normalized edge.
                NormalizeNumberedDisplayCell(equationRange);
                Release(replacement);
                replacement = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata);
            }
            // The replacement range and metadata are already available here.
            // Re-reading both through the new bookmark only to set a temporary
            // caret adds a large COM round-trip and is immediately overwritten
            // by RestoreViewState's final formula selection. Keep the durable
            // inline text boundary normalized directly from the live range.
            if (session.DisplayMode == "inline")
                EnsureInlineBaselineSentinel(
                    equationRange,
                    metadata.FormulaId,
                    placeOutsideNativeMath: true);
            if (session.DisplayMode == "block")
                TryReconcileOmml(document, replacement!, equationRange, metadata);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "reconcile",
                performanceWatch,
                ref performanceCheckpoint);
            finalSelection = equationRange.Duplicate;
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement, deleteContents: true);
            if (replacement is null) TryDelete(equationRange);
            if (numberedTable is not null) TryDelete(numberedTable);
            else if (oldShape is not null) TryDelete(replacementParagraphRange);
            if (document is not null)
            {
                try
                {
                    if (oldBookmarkRemoved
                        && oldRange is not null
                        && originalOmmlMetadata is not null)
                    {
                        var restored = WordOmmlFormulaStore.Wrap(
                            document,
                            oldRange,
                            originalOmmlMetadata);
                        Release(restored);
                        WordOmmlFormulaStore.Save(document, originalOmmlMetadata);
                    }
                    else if (metadataSaved)
                    {
                        WordOmmlFormulaStore.Delete(document, metadata.FormulaId);
                    }
                    if (oldNumberingArtifactsRemoved && oldShape is not null)
                        WordEquationNumbering.TryReconcile(document);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(replacement);
            Release(equationRange);
            Release(replacementParagraphRange);
            Release(replacementParagraph);
            Release(numberedTable);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Replace Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            oldShape = FindByFormulaId(
                    document,
                    session.FormulaId,
                    session.SourceObjectId)
                ?? throw new InvalidOperationException("The target Word formula no longer exists.");
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? session.OriginalMetadata;
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                originalFontSizePt: originalMetadata?.FontSizePt,
                originalRenderFontSizePt: originalMetadata?.RenderFontSizePt);
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                imagePath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            oldShape.Delete();
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            else
                TryReconcileShape(document, replacement, metadata);
            finalSelection = replacement.Range.Duplicate;
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldShape);
            Release(document);
        }
    }

    private static InlineShape AddOleObject(Document document, Range range) =>
        document.InlineShapes.AddOLEObject(
            ClassType: FormulaOleContract.ProgId,
            LinkToFile: false,
            DisplayAsIcon: false,
            Range: range);

    private static void InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            if (oleObject is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "The inserted Word object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = WordOleObjectAccessor.GetRunningObject(format); }
            catch { return false; }
            if (oleObject is not IVisualTeXFormulaObject formula) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            return true;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private Range DuplicateCurrentSelectionRange()
    {
        Selection? selection = null;
        Range? range = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            return range.Duplicate;
        }
        finally
        {
            Release(range);
            Release(selection);
        }
    }

    private WordViewState CaptureViewState()
    {
        Selection? selection = null;
        Range? range = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            try { window = _application.ActiveWindow; } catch { }
            int? vertical = null;
            int? horizontal = null;
            try { vertical = window?.VerticalPercentScrolled; } catch { }
            try { horizontal = window?.HorizontalPercentScrolled; } catch { }
            return new WordViewState
            {
                SelectionStart = range.Start,
                SelectionEnd = range.End,
                VerticalPercentScrolled = vertical,
                HorizontalPercentScrolled = horizontal,
            };
        }
        finally
        {
            Release(window);
            Release(range);
            Release(selection);
        }
    }

    private void RestoreViewState(
        Document? document,
        WordViewState? state,
        Range? preferredSelection)
    {
        if (document is null || state is null) return;
        Selection? selection = null;
        Range? fallback = null;
        Range? content = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            if (preferredSelection is not null)
            {
                selection.SetRange(preferredSelection.Start, preferredSelection.End);
            }
            else
            {
                content = document.Content;
                var start = Math.Max(content.Start, Math.Min(state.SelectionStart, content.End));
                var end = Math.Max(start, Math.Min(state.SelectionEnd, content.End));
                object startValue = start;
                object endValue = end;
                fallback = document.Range(ref startValue, ref endValue);
                selection.SetRange(fallback.Start, fallback.End);
            }
            try { window = _application.ActiveWindow; } catch { }
            if (window is not null)
            {
                try
                {
                    if (state.HorizontalPercentScrolled.HasValue)
                        window.HorizontalPercentScrolled = state.HorizontalPercentScrolled.Value;
                }
                catch { }
                try
                {
                    if (state.VerticalPercentScrolled.HasValue)
                        window.VerticalPercentScrolled = state.VerticalPercentScrolled.Value;
                }
                catch { }
            }
        }
        catch { }
        finally
        {
            Release(window);
            Release(content);
            Release(fallback);
            Release(selection);
        }
    }

    private static InlineShape? FindByFormulaId(
        Document document,
        string formulaId,
        string? sourceObjectIdHint = null)
    {
        Range? hintedRange = null;
        Range? content = null;
        InlineShapes? hintedShapes = null;
        InlineShapes? shapes = null;
        try
        {
            if (TryParseRangeReference(sourceObjectIdHint, out var start, out var end))
            {
                try
                {
                    content = document.Content;
                    if (start >= content.Start && end >= start && end <= content.End)
                    {
                        object startValue = start;
                        object endValue = end;
                        hintedRange = document.Range(ref startValue, ref endValue);
                        hintedShapes = hintedRange.InlineShapes;
                        for (var index = 1; index <= hintedShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            try
                            {
                                candidate = hintedShapes[index];
                                var metadata = WordFormulaMetadataReader.TryRead(candidate);
                                if (string.Equals(
                                        metadata?.FormulaId,
                                        formulaId,
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    var result = candidate;
                                    candidate = null;
                                    return result;
                                }
                            }
                            finally { Release(candidate); }
                        }
                    }
                }
                catch
                {
                    // The document may have shifted while the editor was open.
                    // Fall back to the durable FormulaId scan below.
                }
                finally
                {
                    Release(hintedShapes);
                    hintedShapes = null;
                    Release(hintedRange);
                    hintedRange = null;
                    Release(content);
                    content = null;
                }
            }

            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        var result = shape;
                        shape = null;
                        return result;
                    }
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally
        {
            Release(shapes);
            Release(hintedShapes);
            Release(hintedRange);
            Release(content);
        }
    }

    private UndoRecord? BeginUndoRecord(string name)
    {
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = _application.UndoRecord;
            undoRecord.StartCustomRecord(name);
            return undoRecord;
        }
        catch
        {
            Release(undoRecord);
            return null;
        }
    }

    private static void EndUndoRecord(UndoRecord? undoRecord)
    {
        if (undoRecord is null) return;
        try { undoRecord.EndCustomRecord(); } catch { }
    }

    private void MoveCaretToNormalTypingParagraphAfterNumberedDisplay(
        Document document,
        string formulaId)
    {
        Range? caret = null;
        Selection? selection = null;
        try
        {
            caret = WordEquationNumbering
                .EnsureNormalTypingParagraphAfterNumberedDisplay(document, formulaId);
            if (caret is null) return;
            selection = _application.Selection;
            selection.SetRange(caret.Start, caret.End);
        }
        finally
        {
            Release(selection);
            Release(caret);
        }
    }

    private static void TryReconcileShape(
        Document document,
        InlineShape shape,
        FormulaMetadata metadata)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                ResetDisplayFormulaPosition(range);
            }
            WordEquationNumbering.TryReconcileFormula(
                document,
                range,
                shape.Height,
                metadata);
        }
        finally { Release(range); }
    }

    private static float ReadOmmlFontSize(
        Bookmark bookmark,
        FormulaMetadata metadata)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            font = range.Font;
            var size = font.Size;
            return size > 0 && !float.IsNaN(size) && !float.IsInfinity(size)
                ? FormulaFontSize.Normalize(size)
                : FormulaFontSize.ResolveSemanticFontSize(metadata);
        }
        catch
        {
            return FormulaFontSize.ResolveSemanticFontSize(metadata);
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private static void ApplyOmmlFontSize(Range equationRange, double fontSizePt)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            var normalized = FormulaFontSize.Normalize(fontSizePt);
            font = equationRange.Font;
            font.Position = 0;
            font.Size = normalized;
            try { font.SizeBi = normalized; } catch { }
        }
        finally { Release(font); }
    }

    private static void TryReconcileOmml(
        Document document,
        Bookmark bookmark,
        Range equationRange,
        FormulaMetadata metadata)
    {
        RemoveInlineBaselineSentinel(document, metadata.FormulaId);
        if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            ResetDisplayFormulaPosition(equationRange);
        if (!metadata.Numbered) return;
        // The exact equation range is already available. Re-reading it through
        // the bookmark and enumerating all document OMaths only to estimate a
        // height made this local operation scale with total formula count.
        var height = (float)Math.Max(
            11,
            FormulaFontSize.ResolveSemanticFontSize(metadata) * 1.55);
        WordEquationNumbering.TryReconcileFormula(
            document,
            equationRange,
            height,
            metadata);
    }

    private static void Configure(
        InlineShape shape,
        FormulaMetadata metadata,
        float maxWidth,
        float maxHeight,
        string imagePath,
        float exportedHeight,
        float? exportedBaseline,
        bool alignInline)
    {
        using var image = Image.FromFile(imagePath);
        var ratio = image.Width / (float)Math.Max(1, image.Height);
        // maxWidth/maxHeight are already the SVG's physical size after the
        // 96 dpi CSS-pixel to 72 dpi Word-point conversion. A 12 pt minimum
        // width scales narrow inline formulas (notably x) far above their
        // semantic font size. Keep only a one-point safety floor.
        var width = Math.Max(1f, maxWidth);
        var height = width / ratio;
        if (maxHeight > 0 && height > maxHeight)
        {
            height = maxHeight;
            width = height * ratio;
        }
        // An OLE object is initially created with the placeholder preview's 4:1
        // aspect ratio. Setting only Width while aspect-ratio locking is enabled
        // therefore distorts the real formula. Apply both natural dimensions
        // explicitly, then lock the resolved ratio for later user resizing.
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (!WordFormulaMetadataReader.IsNativeOle(shape))
        {
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encoded;
            shape.AlternativeText = encoded;
        }
        if (alignInline)
            ApplyInlineBaseline(
                shape,
                shape.Height,
                exportedHeight,
                exportedBaseline,
                FormulaFontSize.ResolveSemanticFontSize(metadata));
        else
            ResetDisplayFormulaPosition(shape);
    }

    private static bool ShouldAlignInline(InlineShape shape, FormulaMetadata metadata)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return ShouldAlignInline(range, metadata);
        }
        finally { Release(range); }
    }

    private static bool ShouldAlignInline(Range formulaRange, FormulaMetadata metadata)
    {
        if (string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return true;
        return HasVisibleSurroundingText(formulaRange);
    }

    private static bool HasVisibleSurroundingText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start > paragraphRange.Start)
            {
                before = paragraphRange.Duplicate;
                before.SetRange(paragraphRange.Start, formulaRange.Start);
                if (ContainsVisibleBodyText(before.Text)) return true;
            }
            if (formulaRange.End < paragraphRange.End)
            {
                after = paragraphRange.Duplicate;
                after.SetRange(formulaRange.End, paragraphRange.End);
                if (ContainsVisibleBodyText(after.Text)) return true;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsVisibleBodyText(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var character in value!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a' or '\u0001' or '\u200B')
                continue;
            if (!char.IsWhiteSpace(character)) return true;
        }
        return false;
    }

    private static float? ReadDefinedShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            font = range.Font;
            var position = font.Position;
            return position == (int)WdConstants.wdUndefined
                || position < -256
                || position > 256
                    ? null
                    : position;
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private static void ResetShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetRangeFontPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetDisplayFormulaPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(Range formulaRange)
    {
        ResetRangeFontPosition(formulaRange);
        ResetParagraphTypingPosition(formulaRange);
    }

    private static void ResetParagraphTypingPosition(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? nextCharacter = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End > paragraphRange.Start)
            {
                paragraphMark = paragraphRange.Duplicate;
                paragraphMark.SetRange(paragraphRange.End - 1, paragraphRange.End);
                ResetRangeFontPosition(paragraphMark);
            }

            if (formulaRange.End >= paragraphRange.End) return;
            nextCharacter = paragraphRange.Duplicate;
            nextCharacter.SetRange(
                formulaRange.End,
                Math.Min(formulaRange.End + 1, paragraphRange.End));
            if (nextCharacter.Text is "\v" or "\r" or "\n")
                ResetRangeFontPosition(nextCharacter);
        }
        catch
        {
            // Baseline restoration is best-effort and must not invalidate the
            // formula that has already been inserted or resized.
        }
        finally
        {
            Release(nextCharacter);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static string InlineBaselineBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX formulaId must be a UUID.");
        return InlineBaselineBookmarkPrefix + parsed.ToString("N");
    }


    private Range PrepareInlineBaselineSentinelBeforeInsert(
        Document document,
        Range insertionRange,
        string formulaId)
    {
        RemoveInlineBaselineSentinel(document, formulaId);
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? placeholders = null;
        Range? formulaPlaceholder = null;
        Range? sentinel = null;
        try
        {
            placeholders = insertionRange.Duplicate;
            placeholders.Collapse(WdCollapseDirection.wdCollapseEnd);
            var start = placeholders.Start;
            // Word absorbs the first ordinary character immediately after a
            // newly materialized inline OMath. Keep one sacrificial guard space
            // before the durable VTBL sentinel so the caret can sit in a proven
            // non-math text run.
            placeholders.Text = BulkInlineFormulaPlaceholder
                + InlineMathGuard
                + InlineBaselineSentinel;
            formulaPlaceholder = document.Range(start, start + BulkInlineFormulaPlaceholder.Length);
            sentinel = document.Range(
                formulaPlaceholder.End + InlineMathGuard.Length,
                formulaPlaceholder.End + InlineMathGuard.Length + InlineBaselineSentinel.Length);
            ConfigureInlineBaselineSentinel(sentinel);
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add(InlineBaselineBookmarkName(formulaId), sentinel);
            var result = formulaPlaceholder;
            formulaPlaceholder = null;
            return result;
        }
        finally
        {
            Release(sentinel);
            Release(formulaPlaceholder);
            Release(placeholders);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private int PrepareInlineBaselineSentinelAfterFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                sentinel = bookmark.Range;
                if (IsUsableInlineBaselineSentinel(sentinel, formulaRange))
                {
                    ConfigureInlineBaselineSentinel(sentinel);
                    return sentinel.End;
                }
                bookmark.Delete();
                if (string.Equals(sentinel.Text, InlineBaselineSentinel, StringComparison.Ordinal)
                    || string.Equals(sentinel.Text, LegacyInlineBaselineSentinel, StringComparison.Ordinal))
                    sentinel.Delete();
                Release(sentinel);
                sentinel = null;
                Release(bookmark);
                bookmark = null;
            }

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidOperationException("Word could not locate the inline formula paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var insertionPosition = Math.Max(formulaRange.End, paragraphRange.Start);
            var finalPosition = Math.Max(insertionPosition, paragraphRange.End - 1);
            for (var position = insertionPosition; position <= finalPosition; position++)
            {
                Release(probe);
                object probeStart = position;
                object probeEnd = Math.Min(position + 1, paragraphRange.End);
                probe = document.Range(ref probeStart, ref probeEnd);
                if (RangeContainsMath(probe)) continue;
                insertionPosition = position;
                break;
            }

            object sentinelStart = insertionPosition;
            object sentinelEnd = insertionPosition;
            sentinel = document.Range(ref sentinelStart, ref sentinelEnd);
            sentinel.Text = InlineMathGuard + InlineBaselineSentinel;
            sentinel.SetRange(
                insertionPosition + InlineMathGuard.Length,
                insertionPosition + InlineMathGuard.Length + InlineBaselineSentinel.Length);
            ConfigureInlineBaselineSentinel(sentinel);
            bookmark = bookmarks.Add(name, sentinel);
            return sentinel.End;
        }
        finally
        {
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static bool IsUsableInlineBaselineSentinel(
        Range sentinel,
        Range formulaRange)
    {
        if (!string.Equals(sentinel.Text, InlineBaselineSentinel, StringComparison.Ordinal))
            return false;
        // The durable VTBL bookmark is the hard formula/text boundary. Word can
        // over-report adjacent prose through OMath.Range, so coordinate
        // adjacency is more reliable than character-level OMaths.Count here.
        return sentinel.Start >= formulaRange.End
            && sentinel.Start <= formulaRange.End + 8;
    }

    private static bool RangeContainsMath(Range range)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    private int EnsureInlineBaselineSentinel(
        Range formulaRange,
        string formulaId,
        bool placeOutsideNativeMath = false)
    {
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            bookmarks = document.Bookmarks;
            var name = InlineBaselineBookmarkName(formulaId);
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                sentinel = bookmark.Range;
                if (IsUsableInlineBaselineSentinel(sentinel, formulaRange))
                {
                    ResetParagraphTypingPosition(formulaRange);
                    ConfigureInlineBaselineSentinel(sentinel);
                    return sentinel.End;
                }
            }
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
            Release(document);
        }

        document = _application.ActiveDocument
            ?? throw new InvalidOperationException("No active Word document.");
        try
        {
            var result = PrepareInlineBaselineSentinelAfterFormula(
                document,
                formulaRange,
                formulaId);
            if (placeOutsideNativeMath)
                ResetParagraphTypingPosition(formulaRange);
            return result;
        }
        finally { Release(document); }
    }

    private static void RemoveInlineBaselineSentinel(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            sentinel = bookmark.Range;
            bookmark.Delete();
            if (string.Equals(sentinel.Text, InlineBaselineSentinel, StringComparison.Ordinal)
                || string.Equals(sentinel.Text, LegacyInlineBaselineSentinel, StringComparison.Ordinal))
                sentinel.Delete();
        }
        catch
        {
            // A stale or externally edited sentinel must never block formula work.
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }


    private static void ConfigureInlineBaselineSentinel(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            font.Position = 0;
            font.Hidden = 1;
            font.Name = "Calibri";
        }
        finally { Release(font); }
    }

    private static void ResetRangeFontPosition(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            font.Position = 0;
        }
        finally { Release(font); }
    }

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        double semanticFontSizePoints) =>
        ApplyInlineBaseline(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition: null,
            sourceSemanticFontSizePoints: semanticFontSizePoints,
            targetSemanticFontSizePoints: semanticFontSizePoints);

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            font = range.Font;
            font.Position = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline,
                existingFontPosition,
                sourceSemanticFontSizePoints,
                targetSemanticFontSizePoints);
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private void RestoreTypingBaselineAfter(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            var metadata = WordFormulaMetadataReader.TryRead(shape);
            var caretPosition = metadata is null
                ? range.End
                : EnsureInlineBaselineSentinel(range, metadata.FormulaId);
            RestoreTypingBaselineAfter(range, caretPosition);
        }
        finally { Release(range); }
    }

    private void RestoreTypingBaselineAfter(Bookmark bookmark)
    {
        Range? range = null;
        Document? document = null;
        try
        {
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            document = range.Document;
            var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
            var caretPosition = metadata is null
                ? range.End
                : EnsureInlineBaselineSentinel(
                    range,
                    metadata.FormulaId,
                    placeOutsideNativeMath: true);
            RestoreTypingBaselineAfter(range, caretPosition);
        }
        finally
        {
            Release(document);
            Release(range);
        }
    }

    private void RestoreTypingBaselineAfter(Range formulaRange) =>
        RestoreTypingBaselineAfter(formulaRange, null);

    private void RestoreTypingBaselineAfter(Range formulaRange, int? caretPosition)
    {
        Range? caret = null;
        Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            ResetParagraphTypingPosition(formulaRange);
            caret = formulaRange.Duplicate;
            if (caretPosition.HasValue)
                caret.SetRange(caretPosition.Value, caretPosition.Value);
            else
                caret.Collapse(WdCollapseDirection.wdCollapseEnd);
            try
            {
                font = caret.Font;
                font.Position = 0;
                font.Hidden = 0;
            }
            catch
            {
                // A collapsed range immediately after a locked hidden content
                // control can reject direct font writes. The structural caret
                // position is authoritative; formatting reset is best-effort.
            }
            finally
            {
                Release(font);
                font = null;
            }

            selection = _application.Selection;
            selection.SetRange(caret.Start, caret.End);
            try
            {
                font = selection.Font;
                font.Position = 0;
                font.Hidden = 0;
            }
            catch
            {
                // Keep the caret outside the formula even if Word refuses to
                // mutate the insertion-point font at this boundary.
            }
        }
        finally
        {
            Release(font);
            Release(selection);
            Release(caret);
        }
    }

    private static OfficeObjectResult Result(OfficeSessionDocument session, Document document) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(document),
            ObjectId = session.FormulaId,
        };

    private static string RangeReference(Range range) =>
        $"{RangeReferencePrefix}{range.Start}:{range.End}";

    private static Range? TryResolveOmmlRangeReference(
        Document document,
        string? sourceObjectId)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return null;
        Range? content = null;
        Range? candidate = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            content = document.Content;
            if (start < content.Start || end < start || end > content.End)
                return null;
            object startValue = start;
            object endValue = end;
            candidate = document.Range(ref startValue, ref endValue);
            maths = candidate.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            var result = math.Range.Duplicate;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(math);
            Release(maths);
            Release(candidate);
            Release(content);
        }
    }

    private static Range ResolveSourceRange(
        Document document,
        string? sourceObjectId,
        Selection selection)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return selection.Range.Duplicate;
        Range? content = null;
        try
        {
            content = document.Content;
            if (start < 0 || end < start || end > content.End)
                throw new InvalidOperationException(
                    "The Word insertion range selected when the formula editor opened is no longer valid.");
            object startValue = start;
            object endValue = end;
            return document.Range(ref startValue, ref endValue);
        }
        finally { Release(content); }
    }

    private static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0 || separator >= payload.Length - 1) return false;
        return int.TryParse(payload.Substring(0, separator), out start)
            && int.TryParse(payload.Substring(separator + 1), out end);
    }

    private static void EnsureSourceDocument(
        Document document,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(document);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active Word document changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Document document)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(document.FullName)) return document.FullName;
        }
        catch { }
        return document.Name;
    }

    private static void EnsureWritable(Document document)
    {
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("The active Word document is read-only.");
    }

    private static bool HasLeadingTab(Document document, Range formulaRange)
    {
        if (formulaRange.Start <= 0) return false;
        Range? preceding = null;
        try
        {
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return true;
            if (!string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                || formulaRange.Start <= 1)
                return false;
            preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
    }

    private static Range CreateNumberedDisplayTable(
        Document document,
        Range anchor,
        out Table table)
    {
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Borders? borders = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        try
        {
            anchor.InsertParagraphAfter();
            anchor.Collapse(WdCollapseDirection.wdCollapseEnd);
            table = document.Tables.Add(anchor, 1, 3);
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            borders = table.Borders;
            borders.Enable = 0;
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

            centerCell = table.Cell(1, 2);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerCellRange = centerCell.Range;
            var insertion = centerCellRange.Duplicate;
            insertion.End = Math.Max(insertion.Start, insertion.End - 1);
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            return insertion;
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(borders);
            Release(centerCellRange);
            Release(centerCell);
        }
    }

    private static void NormalizeNumberedDisplayCell(Range formulaRange)
    {
        Document? document = null;
        Table? table = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Range? cellRange = null;
        Range? character = null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                || formulaRange.Tables.Count == 0)
                return;
            document = formulaRange.Document;
            table = formulaRange.Tables[1];
            columns = table.Columns;
            if (columns.Count < 3) return;
            centerCell = table.Cell(1, 2);
            cellRange = centerCell.Range;

            // A display OMath inserted next to the source OLE can leave one
            // manual line break on each side. Delete only those exact control
            // characters, scanning backwards so Word's shifting ranges cannot
            // expand across and remove the replacement formula object.
            for (var position = cellRange.End - 2;
                 position >= cellRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start
                    && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(
                    ref characterStart,
                    ref characterEnd);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
                Release(character);
                character = null;
            }
        }
        finally
        {
            Release(character);
            Release(cellRange);
            Release(centerCell);
            Release(columns);
            Release(table);
            Release(document);
        }
    }

    private static void NormalizeNumberedDisplayCell(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            NormalizeNumberedDisplayCell(range);
        }
        finally { Release(range); }
    }

    private static void ValidateInsertedOmml(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create exactly one native OMML equation.");
            math = maths[1];
            mathRange = math.Range;
            var wordOpenXml = mathRange.WordOpenXML;
            if (mathRange.End <= mathRange.Start
                || string.IsNullOrWhiteSpace(wordOpenXml)
                || wordOpenXml.IndexOf("oMath", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Word returned an empty native OMML equation.");
            // Validate the structure after Word has imported and normalized the
            // OMML. Pre-insertion XML checks cannot catch empty slots that Word
            // introduces while materializing the native equation tree.
            WordOmmlConverter.ValidateMaterializedOmml(wordOpenXml);
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void TryDelete(InlineShape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void TryDelete(Table? table)
    {
        if (table is null) return;
        try { table.Delete(); } catch { }
    }

    private static void TryDelete(Bookmark? bookmark, bool deleteContents)
    {
        if (bookmark is null) return;
        Range? range = null;
        try
        {
            if (deleteContents) range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            bookmark.Delete();
            if (deleteContents) range?.Delete();
        }
        catch { }
        finally { Release(range); }
    }

    private static void TryDelete(Range? range)
    {
        if (range is null) return;
        try { range.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
