using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

[ComImport]
[Guid("00000112-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleObjectNative
{
    [PreserveSig] int SetClientSite(IntPtr clientSite);
    [PreserveSig] int GetClientSite(out IntPtr clientSite);
    [PreserveSig] int SetHostNames(IntPtr containerApp, IntPtr containerObject);
    [PreserveSig] int Close(uint saveOption);
    [PreserveSig] int SetMoniker(uint whichMoniker, IntPtr moniker);
    [PreserveSig] int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker);
    [PreserveSig] int InitFromData(IntPtr dataObject, int creation, uint reserved);
    [PreserveSig] int GetClipboardData(uint reserved, out IntPtr dataObject);
    [PreserveSig] int DoVerb(
        int verb,
        IntPtr message,
        IntPtr activeSite,
        int index,
        IntPtr parentWindow,
        IntPtr positionRect);
    [PreserveSig] int EnumVerbs(out IntPtr enumerator);
    [PreserveSig] int Update();
    [PreserveSig] int IsUpToDate();
    [PreserveSig] int GetUserClassId(out Guid classId);
    [PreserveSig] int GetUserType(uint formOfType, out IntPtr userType);
    [PreserveSig] int SetExtent(uint drawAspect, ref OleSize size);
    [PreserveSig] int GetExtent(uint drawAspect, out OleSize size);
}

[StructLayout(LayoutKind.Sequential)]
internal struct OleSize
{
    internal int Cx;
    internal int Cy;
}

internal sealed class PowerPointFormulaService
{
    private const string FormulaIdTag = "VisualTeXFormulaId";
    private const string MetadataTag = "VisualTeXMetadata";
    private const string SlideReferencePrefix = "visualtex-ppt-vsto-slide:";
    private readonly Application _application;

    public PowerPointFormulaService(Application application)
    {
        _application = application;
    }

    public OfficeSelection ReadSelection() => ReadSelection(null);

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Selection? selection = null;
        ShapeRange? range = null;
        Shape? shape = null;
        var ownsSelection = providedSelection is null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = (Slide)view.Slide;
            selection = providedSelection ?? window.Selection;
            FormulaMetadata? metadata = null;
            string? objectMode = null;
            string? objectId = SlideReference(slide);
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                range = selection.ShapeRange;
                if (range.Count == 1)
                {
                    shape = range[1];
                    objectId = shape.Name;
                    metadata = ReadMetadata(shape);
                    if (metadata is not null)
                    {
                        metadata.FontSizePt = FormulaFontSize.InferOleFontSize(
                            shape.Width,
                            shape.Height,
                            metadata);
                        objectMode = IsNativeOle(shape)
                            ? "nativeOle"
                            : "crossPlatformPicture";
                    }
                }
            }
            return new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = DocumentIdentity(presentation),
                ObjectId = objectId,
                ReadOnly = presentation.ReadOnly == MsoTriState.msoTrue,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
                ObjectMode = objectMode,
            };
        }
        finally
        {
            Release(shape);
            Release(range);
            if (ownsSelection) Release(selection);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public float ReadCurrentTypingFontSize()
    {
        object? mathRange = null;
        object? font = null;
        try
        {
            if (TryGetSelectedMathRange(out mathRange))
            {
                dynamic range = mathRange!;
                font = range.Font;
                var size = Convert.ToDouble(((dynamic)font).Size);
                return FormulaFontSize.Normalize(size, 20f);
            }

            var window = _application.ActiveWindow;
            if (window is null) return 20f;
            Selection? selection = null;
            try
            {
                selection = window.Selection;
                dynamic selected = selection;
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    object textRange = selected.TextRange2;
                    try
                    {
                        font = ((dynamic)textRange).Font;
                        var size = Convert.ToDouble(((dynamic)font).Size);
                        return FormulaFontSize.Normalize(size, 20f);
                    }
                    finally { Release(textRange); }
                }
            }
            finally { Release(selection); Release(window); }
            return 20f;
        }
        catch { return 20f; }
        finally
        {
            Release(font);
            Release(mathRange);
        }
    }

    public float? GetSelectedFormulaFontSize()
    {
        try
        {
            var selected = ReadSelection();
            if (selected.Metadata is not null)
                return FormulaFontSize.ResolveSemanticFontSize(selected.Metadata);
        }
        catch { }

        object? mathRange = null;
        object? font = null;
        try
        {
            if (!TryGetSelectedMathRange(out mathRange)) return null;
            dynamic range = mathRange!;
            font = range.Font;
            var size = Convert.ToDouble(((dynamic)font).Size);
            return FormulaFontSize.Normalize(size, 18f);
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(mathRange);
        }
    }

    public float SetSelectedFormulaFontSize(double requestedFontSizePt)
    {
        var target = FormulaFontSize.Normalize(requestedFontSizePt);
        var selected = ReadSelection();
        if (selected.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
        {
            if (SetSelectedMathZoneFontSize(target)) return target;
            throw new InvalidOperationException("请先选择一个 VisualTeX 公式或 PowerPoint 原生公式。");
        }

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, shape) = FindFormula(presentation, selected.FormulaId!, selected.ObjectId);
            if (slide is null || shape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");

            var metadata = selected.Metadata;
            var currentFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
            var fontScale = target / Math.Max(0.5f, currentFontSize);
            var size = ScaleCurrentShapeSize(
                shape.Width,
                shape.Height,
                fontScale,
                600f,
                400f);
            metadata.FontSizePt = target;
            var centerX = shape.Left + shape.Width / 2f;
            var centerY = shape.Top + shape.Height / 2f;
            if (IsNativeOle(shape))
                ApplyOleSizeAndRefresh(shape, size.Width, size.Height);
            else
            {
                shape.LockAspectRatio = MsoTriState.msoFalse;
                shape.Width = size.Width;
                shape.Height = size.Height;
                shape.LockAspectRatio = MsoTriState.msoTrue;
            }
            shape.Left = centerX - size.Width / 2f;
            shape.Top = centerY - size.Height / 2f;
            Configure(shape, metadata);
            return target;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(presentation);
        }
    }

    private bool SetSelectedMathZoneFontSize(float target)
    {
        object? mathRange = null;
        object? font = null;
        try
        {
            if (!TryGetSelectedMathRange(out mathRange)) return false;
            StartNewUndoEntry();
            dynamic range = mathRange!;
            font = range.Font;
            ((dynamic)font).Size = target;
            return true;
        }
        finally
        {
            StartNewUndoEntry();
            Release(font);
            Release(mathRange);
        }
    }

    private bool TryGetSelectedMathRange(out object? mathRange)
    {
        mathRange = null;
        DocumentWindow? window = null;
        Selection? selection = null;
        ShapeRange? shapes = null;
        Shape? shape = null;
        object? textRange = null;
        try
        {
            window = _application.ActiveWindow;
            if (window is null) return false;
            selection = window.Selection;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                dynamic selected = selection;
                textRange = selected.TextRange2;
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                shapes = selection.ShapeRange;
                if (shapes.Count != 1) return false;
                shape = shapes[1];
                if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
                dynamic textFrame = shape.TextFrame2;
                textRange = textFrame.TextRange;
                Release(textFrame);
            }
            else
            {
                return false;
            }

            dynamic range = textRange!;
            object candidate;
            try { candidate = range.MathZones(Type.Missing, Type.Missing); }
            catch (COMException) { candidate = range.MathZones(); }
            var length = Convert.ToInt32(((dynamic)candidate).Length);
            if (length <= 0)
            {
                Release(candidate);
                return false;
            }
            mathRange = candidate;
            return true;
        }
        catch
        {
            Release(mathRange);
            mathRange = null;
            return false;
        }
        finally
        {
            Release(textRange);
            Release(shape);
            Release(shapes);
            Release(selection);
            Release(window);
        }
    }

    public string DeleteSelectedFormula()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, shape) = FindFormula(presentation, requiredFormulaId, selected.ObjectId);
            if (slide is null || shape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");
            shape.Delete();
            return requiredFormulaId;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(presentation);
        }
    }

    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        OLEFormat? format = null;
        object? oleObject = null;
        string? pngPath = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, oldShape) = FindFormula(presentation, requiredFormulaId, selected.ObjectId);
            if (slide is null || oldShape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");
            var metadata = ReadMetadata(oldShape)
                ?? throw new InvalidDataException("The selected formula metadata is invalid.");
            format = oldShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected formula is already a picture.");
            oleObject = format.Object;
            pngPath = OlePngPreviewExtractor.MaterializePng(oleObject, requiredFormulaId);

            var left = oldShape.Left;
            var top = oldShape.Top;
            var width = oldShape.Width;
            var height = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            replacement = slide.Shapes.AddPicture(
                pngPath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                left,
                top,
                width,
                height);
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            oldShape.Delete();
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
            StartNewUndoEntry();
            Release(oleObject);
            Release(format);
            Release(replacement);
            Release(oldShape);
            Release(slide);
            Release(presentation);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = ResolveTargetSlide(presentation, session.SourceObjectId, view);
            var width = Math.Max(12f, (session.ExportResult?.Width ?? 240) * 0.75f);
            var height = Math.Max(12f, (session.ExportResult?.Height ?? 80) * 0.75f);
            var scale = Math.Min(1f, Math.Min(600f / width, 400f / height));
            width *= scale;
            height *= scale;
            var left = Math.Max(0f, (presentation.PageSetup.SlideWidth - width) / 2f);
            var top = Math.Max(0f, (presentation.PageSetup.SlideHeight - height) / 2f);
            shape = slide.Shapes.AddPicture(
                imagePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                left,
                top,
                width,
                height);
            Configure(shape, metadata);
            return Result(session, presentation, shape.Name);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = ResolveTargetSlide(presentation, session.SourceObjectId, view);
            var width = Math.Max(12f, (session.ExportResult?.Width ?? 240) * 0.75f);
            var height = Math.Max(12f, (session.ExportResult?.Height ?? 80) * 0.75f);
            var scale = Math.Min(1f, Math.Min(600f / width, 400f / height));
            width *= scale;
            height *= scale;
            var left = Math.Max(0f, (presentation.PageSetup.SlideWidth - width) / 2f);
            var top = Math.Max(0f, (presentation.PageSetup.SlideHeight - height) / 2f);
            shape = AddOleObject(slide, left, top, width, height);
            InitializeOle(shape, metadata, emfPath, pngPath);
            ApplyOleSizeAndRefresh(shape, width, height);
            RestoreOlePosition(shape, left, top);
            Configure(shape, metadata);
            return Result(session, presentation, shape.Name);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            (slide, oldShape) = FindFormula(
                presentation,
                session.FormulaId,
                session.SourceObjectId);
            if (slide is null || oldShape is null)
                throw new InvalidOperationException("The target PowerPoint formula no longer exists.");

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = ReadMetadata(oldShape) ?? session.OriginalMetadata;
            var convertingPictureToOle = !IsNativeOle(oldShape);
            var editedSize = convertingPictureToOle
                && FormulaContentEquivalent(originalMetadata, metadata)
                    ? (Width: oldWidth, Height: oldHeight)
                    : OfficeFormulaSizing.EditedSize(
                        oldWidth,
                        oldHeight,
                        originalMetadata?.RenderWidthPx,
                        originalMetadata?.RenderHeightPx,
                        session.ExportResult?.Width ?? oldWidth / 0.75f,
                        session.ExportResult?.Height ?? oldHeight / 0.75f,
                        600f,
                        400f,
                        originalMetadata?.FontSizePt,
                        originalMetadata?.RenderFontSizePt);
            var newLeft = left + (oldWidth - editedSize.Width) / 2f;
            var newTop = top + (oldHeight - editedSize.Height) / 2f;

            if (TryUpdateOle(oldShape, metadata, emfPath, pngPath))
            {
                ApplyOleSizeAndRefresh(oldShape, editedSize.Width, editedSize.Height);
                RestoreOlePosition(oldShape, newLeft, newTop);
                Configure(oldShape, metadata);
                return Result(session, presentation, oldShape.Name);
            }

            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            replacement = AddOleObject(
                slide,
                newLeft,
                newTop,
                editedSize.Width,
                editedSize.Height);
            InitializeOle(replacement, metadata, emfPath, pngPath);
            ApplyOleSizeAndRefresh(replacement, editedSize.Width, editedSize.Height);
            RestoreOlePosition(replacement, newLeft, newTop);
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            oldShape.Delete();
            return Result(session, presentation, replacement.Name);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(replacement);
            Release(oldShape);
            Release(slide);
            Release(presentation);
        }
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            (slide, oldShape) = FindFormula(
                presentation,
                session.FormulaId,
                session.SourceObjectId);
            if (slide is null || oldShape is null)
                throw new InvalidOperationException("The target PowerPoint formula no longer exists.");

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            var originalMetadata = ReadMetadata(oldShape) ?? session.OriginalMetadata;
            var convertingOleToPicture = IsNativeOle(oldShape);
            var editedSize = convertingOleToPicture
                && FormulaContentEquivalent(originalMetadata, metadata)
                    ? (Width: oldWidth, Height: oldHeight)
                    : OfficeFormulaSizing.EditedSize(
                        oldWidth,
                        oldHeight,
                        originalMetadata?.RenderWidthPx,
                        originalMetadata?.RenderHeightPx,
                        session.ExportResult?.Width ?? oldWidth / 0.75f,
                        session.ExportResult?.Height ?? oldHeight / 0.75f,
                        600f,
                        400f,
                        originalMetadata?.FontSizePt,
                        originalMetadata?.RenderFontSizePt);
            var newLeft = left + (oldWidth - editedSize.Width) / 2f;
            var newTop = top + (oldHeight - editedSize.Height) / 2f;

            replacement = slide.Shapes.AddPicture(
                imagePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                newLeft,
                newTop,
                editedSize.Width,
                editedSize.Height);
            // PowerPoint 2021 can ignore one of the AddPicture dimensions for
            // SVG and immediately restore the file's intrinsic aspect ratio.
            // Reapply the exact formula box after the SVG shape exists so a
            // lossless OLE→picture conversion keeps the user's physical size.
            ApplyPictureSize(replacement, editedSize.Width, editedSize.Height);
            replacement.Left = newLeft;
            replacement.Top = newTop;
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            oldShape.Delete();
            return Result(session, presentation, replacement.Name);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(replacement);
            Release(oldShape);
            Release(slide);
            Release(presentation);
        }
    }

    private void StartNewUndoEntry()
    {
        try { _application.StartNewUndoEntry(); } catch { }
    }

    private static Shape AddOleObject(
        Slide slide,
        float left,
        float top,
        float width,
        float height) =>
        slide.Shapes.AddOLEObject(
            left,
            top,
            width,
            height,
            FormulaOleContract.ProgId,
            string.Empty,
            MsoTriState.msoFalse,
            string.Empty,
            0,
            string.Empty,
            MsoTriState.msoFalse);

    private static void InitializeOle(
        Shape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = format.Object;
            if (oleObject is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "The inserted PowerPoint object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        Shape shape,
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
            try { oleObject = format.Object; }
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

    private static (Slide? Slide, Shape? Shape) FindFormula(
        Presentation presentation,
        string formulaId,
        string? preferredObjectId)
    {
        Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (var slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        Shape? shape = null;
                        try
                        {
                            shape = shapes[shapeIndex];
                            var metadata = ReadMetadata(shape);
                            var preferredMatch = !string.IsNullOrEmpty(preferredObjectId)
                                && shape.Name == preferredObjectId;
                            var metadataMatch = metadata?.FormulaId == formulaId
                                || shape.Name == $"VisualTeX_{formulaId}";
                            if (metadataMatch || (preferredMatch && metadata?.FormulaId == formulaId))
                            {
                                var foundSlide = slide;
                                var foundShape = shape;
                                slide = null;
                                shape = null;
                                return (foundSlide, foundShape);
                            }
                        }
                        finally { Release(shape); }
                    }
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }
            return (null, null);
        }
        finally { Release(slides); }
    }

    private static FormulaMetadata? ReadMetadata(Shape shape)
    {
        if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
            and not MsoShapeType.msoLinkedOLEObject)
            return ReadPictureMetadata(shape);

        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return ReadPictureMetadata(shape); }
            string? progId;
            try { progId = format.ProgID; }
            catch { return ReadPictureMetadata(shape); }
            if (!string.Equals(
                    progId,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                return ReadPictureMetadata(shape);
            try { oleObject = format.Object; }
            catch { return null; }
            return oleObject is IVisualTeXFormulaObject formula
                ? FormulaOleInterop.ReadMetadata(formula)
                : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static FormulaMetadata? ReadPictureMetadata(Shape shape)
    {
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            string? encoded = null;
            try { encoded = tags[MetadataTag]; } catch { }
            FormulaMetadata? metadata = FormulaMetadataCodec.Decode(encoded);
            if (metadata is not null) return metadata;
            try { encoded = shape.AlternativeText; } catch { encoded = null; }
            return FormulaMetadataCodec.Decode(encoded);
        }
        finally { Release(tags); }
    }

    private static bool FormulaContentEquivalent(
        FormulaMetadata? original,
        FormulaMetadata current)
    {
        if (original is null) return false;
        return string.Equals(
                NormalizeFormulaText(original.Latex),
                NormalizeFormulaText(current.Latex),
                StringComparison.Ordinal)
            && string.Equals(
                original.DisplayMode,
                current.DisplayMode,
                StringComparison.Ordinal);
    }

    private static string NormalizeFormulaText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace("\r", "\n")
            .Trim();

    private static bool IsNativeOle(Shape shape)
    {
        if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
            and not MsoShapeType.msoLinkedOLEObject)
            return false;

        OLEFormat? format = null;
        try
        {
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
        finally { Release(format); }
    }

    private static (float Width, float Height) ScaleCurrentShapeSize(
        float currentWidth,
        float currentHeight,
        float requestedScale,
        float maximumWidth,
        float maximumHeight)
    {
        var scale = requestedScale;
        if (float.IsNaN(scale) || float.IsInfinity(scale) || scale <= 0)
            scale = 1f;
        scale = Math.Max(0.1f, Math.Min(10f, scale));
        var width = Math.Max(1f, currentWidth) * scale;
        var height = Math.Max(1f, currentHeight) * scale;
        var fit = Math.Min(
            1f,
            Math.Min(
                maximumWidth > 0 ? maximumWidth / width : 1f,
                maximumHeight > 0 ? maximumHeight / height : 1f));
        if (!float.IsNaN(fit) && !float.IsInfinity(fit) && fit > 0 && fit < 1f)
        {
            width *= fit;
            height *= fit;
        }
        return (Math.Max(1f, width), Math.Max(1f, height));
    }

    private static void ApplyPictureSize(Shape shape, float width, float height)
    {
        shape.LockAspectRatio = MsoTriState.msoFalse;
        shape.Width = Math.Max(1f, width);
        shape.Height = Math.Max(1f, height);
        shape.LockAspectRatio = MsoTriState.msoTrue;
    }

    private static void ApplyOleSizeAndRefresh(Shape shape, float width, float height)
    {
        // PowerPoint 2021 may resize the outer Shape first, then restore the OLE
        // server's previous HIMETRIC extent when the data-change notification
        // arrives. Keep both sides synchronized, then reapply the host box after
        // the server notification so width and height remain uniformly scaled.
        ApplyPictureSize(shape, width, height);
        SetOleServerExtent(shape, width, height);
        ApplyPictureSize(shape, width, height);
        // Do not invoke an OLE verb here. PowerPoint's DoVerb API accepts only
        // host verb indexes (0..n), not OLEIVERB_SHOW (-1), and the primary
        // verb would activate the editor. The LocalServer data-change
        // notification plus CF_ENHMETAFILE/CF_METAFILEPICT presentations own
        // the preview refresh instead.
    }

    private static void SetOleServerExtent(Shape shape, float width, float height)
    {
        const uint dvaspectContent = 1;
        const double himetricPerPoint = 2540.0 / 72.0;
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                return;
            oleObject = format.Object;
            if (oleObject is not IOleObjectNative nativeOle)
                throw new InvalidCastException(
                    "The VisualTeX OLE object does not expose IOleObject.");
            var size = new OleSize
            {
                Cx = Math.Max(1, (int)Math.Round(Math.Max(1f, width) * himetricPerPoint)),
                Cy = Math.Max(1, (int)Math.Round(Math.Max(1f, height) * himetricPerPoint)),
            };
            var result = nativeOle.SetExtent(dvaspectContent, ref size);
            if (result < 0)
                Marshal.ThrowExceptionForHR(result);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static void RestoreOlePosition(Shape shape, float left, float top)
    {
        // PowerPoint can reset a newly initialized OLE object's position while
        // it synchronizes the server extent and cached presentation. Position
        // is therefore the final geometry operation, after width and height.
        shape.Left = left;
        shape.Top = top;
    }

    private static void Configure(Shape shape, FormulaMetadata metadata)
    {
        shape.LockAspectRatio = MsoTriState.msoTrue;
        shape.Name = $"VisualTeX_{metadata.FormulaId}";
        if (IsNativeOle(shape)) return;

        var encoded = FormulaMetadataCodec.Encode(metadata);
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            tags.Add(FormulaIdTag, metadata.FormulaId);
            tags.Add(MetadataTag, encoded);
        }
        finally { Release(tags); }
        shape.AlternativeText = encoded;
    }

    private static void TryApplyRotation(Shape shape, float rotation)
    {
        if (Math.Abs(rotation) < 0.01f) return;
        try { shape.Rotation = rotation; } catch { }
    }

    private static void MoveToZOrder(Shape shape, int target)
    {
        for (var attempts = 0; attempts < 512 && shape.ZOrderPosition > target; attempts++)
            shape.ZOrder(MsoZOrderCmd.msoSendBackward);
    }

    private static OfficeObjectResult Result(
        OfficeSessionDocument session,
        Presentation presentation,
        string objectId) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(presentation),
            ObjectId = objectId,
        };

    private static string SlideReference(Slide slide) =>
        $"{SlideReferencePrefix}{slide.SlideID}:{slide.SlideIndex}";

    private static Slide ResolveTargetSlide(
        Presentation presentation,
        string? sourceObjectId,
        View view)
    {
        if (TryParseSlideReference(sourceObjectId, out var slideId))
        {
            Slides? slides = null;
            try
            {
                slides = presentation.Slides;
                for (var index = 1; index <= slides.Count; index++)
                {
                    Slide? candidate = null;
                    try
                    {
                        candidate = slides[index];
                        if (candidate.SlideID != slideId) continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            finally { Release(slides); }
            throw new InvalidOperationException(
                "The PowerPoint slide selected when the formula editor opened no longer exists.");
        }
        return (Slide)view.Slide;
    }

    private static bool TryParseSlideReference(string? value, out int slideId)
    {
        slideId = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(SlideReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(SlideReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator >= 0) payload = payload.Substring(0, separator);
        return int.TryParse(payload, out slideId) && slideId > 0;
    }

    private static void EnsureSourceDocument(
        Presentation presentation,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(presentation);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active PowerPoint presentation changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Presentation presentation)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(presentation.FullName)) return presentation.FullName;
        }
        catch { }
        return presentation.Name;
    }

    private static void EnsureWritable(Presentation presentation)
    {
        if (presentation.ReadOnly == MsoTriState.msoTrue)
            throw new UnauthorizedAccessException("The active PowerPoint presentation is read-only.");
    }

    private void EnsureNotSlideShow()
    {
        SlideShowWindows? windows = null;
        try
        {
            windows = _application.SlideShowWindows;
            if (windows.Count > 0)
                throw new InvalidOperationException("PowerPoint slide show mode does not allow formula editing.");
        }
        finally { Release(windows); }
    }

    private static void TryDelete(Shape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
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
