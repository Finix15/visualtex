using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WinForms = System.Windows.Forms;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

[ComImport]
[Guid("00000112-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface GeometryOleObjectNative
{
    [PreserveSig] int SetClientSite(IntPtr clientSite);
    [PreserveSig] int GetClientSite(out IntPtr clientSite);
    [PreserveSig] int SetHostNames(IntPtr containerApp, IntPtr containerObject);
    [PreserveSig] int Close(uint saveOption);
    [PreserveSig] int SetMoniker(uint whichMoniker, IntPtr moniker);
    [PreserveSig] int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker);
    [PreserveSig] int InitFromData(IntPtr dataObject, int creation, uint reserved);
    [PreserveSig] int GetClipboardData(uint reserved, out IntPtr dataObject);
    [PreserveSig] int DoVerb(int verb, IntPtr message, IntPtr activeSite, int index, IntPtr parentWindow, IntPtr positionRect);
    [PreserveSig] int EnumVerbs(out IntPtr enumerator);
    [PreserveSig] int Update();
    [PreserveSig] int IsUpToDate();
    [PreserveSig] int GetUserClassId(out Guid classId);
    [PreserveSig] int GetUserType(uint formOfType, out IntPtr userType);
    [PreserveSig] int SetExtent(uint drawAspect, ref GeometryOleSize size);
    [PreserveSig] int GetExtent(uint drawAspect, out GeometryOleSize size);
}

[StructLayout(LayoutKind.Sequential)]
internal struct GeometryOleSize
{
    internal int Cx;
    internal int Cy;
}

internal static partial class Program
{
    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern IntPtr SetEnhMetaFileBits(uint bufferSize, byte[] data);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern bool PlayEnhMetaFile(IntPtr hdc, IntPtr metafile, ref NativeRect bounds);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern bool DeleteEnhMetaFile(IntPtr metafile);

    private const string PowerPointGeometryLatex = @"x=\frac{-b\pm \sqrt{b^{2}-4ac}}{2a}";
    private const float PowerPointGeometryRenderWidth = 509.49335f;
    private const float PowerPointGeometryRenderHeight = 144.72f;
    private const float PowerPointGeometryFontSize = 40f;

    private sealed class PowerPointGeometryFixture
    {
        internal PowerPointGeometryFixture(
            string latex,
            string svg,
            string? pngBase64,
            float renderWidth,
            float renderHeight,
            float baseline,
            float fontSizePt)
        {
            Latex = latex;
            Svg = svg;
            PngBase64 = pngBase64;
            RenderWidth = renderWidth;
            RenderHeight = renderHeight;
            Baseline = baseline;
            FontSizePt = fontSizePt;
        }

        internal string Latex { get; }
        internal string Svg { get; }
        internal string? PngBase64 { get; }
        internal float RenderWidth { get; }
        internal float RenderHeight { get; }
        internal float Baseline { get; }
        internal float FontSizePt { get; }
    }

    private static PowerPointGeometryFixture LoadPowerPointGeometryFixtureFromEnvironment()
    {
        var sessionPath = Environment.GetEnvironmentVariable("VISUALTEX_POWERPOINT_REAL_SESSION_PATH");
        if (string.IsNullOrWhiteSpace(sessionPath))
        {
            return new PowerPointGeometryFixture(
                PowerPointGeometryLatex,
                CreateSvg(PowerPointGeometryRenderWidth, PowerPointGeometryRenderHeight),
                null,
                PowerPointGeometryRenderWidth,
                PowerPointGeometryRenderHeight,
                97.6f,
                PowerPointGeometryFontSize);
        }
        if (!File.Exists(sessionPath))
            throw new FileNotFoundException("Real PowerPoint Session fixture does not exist.", sessionPath);

        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(sessionPath));
        var root = document.RootElement;
        var export = root.GetProperty("exportResult");
        var svg = export.GetProperty("svg").GetString();
        if (string.IsNullOrWhiteSpace(svg))
            throw new InvalidDataException("Real PowerPoint Session has no SVG export.");
        var latex = root.GetProperty("lines")[0].GetProperty("latex").GetString()
            ?? PowerPointGeometryLatex;
        var pngBase64 = export.TryGetProperty("pngBase64", out var png)
            ? png.GetString()
            : null;
        var fixture = new PowerPointGeometryFixture(
            latex,
            svg!,
            pngBase64,
            export.GetProperty("width").GetSingle(),
            export.GetProperty("height").GetSingle(),
            export.GetProperty("baseline").GetSingle(),
            root.GetProperty("fontSizePt").GetSingle());
        Console.WriteLine(
            $"Using real PowerPoint Session SVG: {sessionPath}; " +
            $"{fixture.RenderWidth:F3}x{fixture.RenderHeight:F3}px, {fixture.FontSizePt:F1}pt.");
        return fixture;
    }

    private static void RunPowerPointCurrentOleCacheProbe()
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? sourcePresentation = null;
        PowerPoint.Slide? sourceSlide = null;
        PowerPoint.Shapes? sourceShapes = null;
        PowerPoint.Shape? sourceShape = null;
        PowerPoint.Presentation? tempPresentation = null;
        PowerPoint.Slide? tempSlide = null;
        PowerPoint.ShapeRange? pastedRange = null;
        PowerPoint.Shape? tempShape = null;
        PowerPoint.Tags? tags = null;
        PowerPoint.OLEFormat? format = null;
        object? oleObject = null;
        string? fixtureSvgPath = null;
        try
        {
            application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            sourcePresentation = application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            sourceSlide = (PowerPoint.Slide)application.ActiveWindow.View.Slide;
            sourceShapes = sourceSlide.Shapes;
            for (var index = 1; index <= sourceShapes.Count; index++)
            {
                var candidate = sourceShapes[index];
                if (candidate.Name.StartsWith("VisualTeX_", StringComparison.OrdinalIgnoreCase))
                {
                    sourceShape = candidate;
                    break;
                }
                Release(candidate);
            }
            if (sourceShape is null)
                throw new InvalidOperationException("No VisualTeX shape exists on the active PowerPoint slide.");

            Console.WriteLine(
                $"Source OLE before copy: {sourceShape.Name} {sourceShape.Width:F2}x{sourceShape.Height:F2} pt " +
                $"at ({sourceShape.Left:F2},{sourceShape.Top:F2}).");
            sourceShape.Copy();

            tempPresentation = application.Presentations.Add(MsoTriState.msoTrue);
            tempSlide = tempPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            pastedRange = tempSlide.Shapes.Paste();
            tempShape = pastedRange[1];
            Console.WriteLine(
                $"Temporary pasted OLE: {tempShape.Width:F2}x{tempShape.Height:F2} pt " +
                $"at ({tempShape.Left:F2},{tempShape.Top:F2}).");

            string? encoded = null;
            try
            {
                tags = tempShape.Tags;
                encoded = tags["VisualTeXMetadata"];
            }
            catch { }
            var metadata = FormulaMetadataCodec.Decode(encoded)
                ?? throw new InvalidDataException("Copied VisualTeX OLE has no readable metadata tag.");
            var fixture = LoadPowerPointGeometryFixtureFromEnvironment();
            fixtureSvgPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                $"geometry-repair-{Guid.NewGuid():N}.svg");
            Directory.CreateDirectory(Path.GetDirectoryName(fixtureSvgPath)!);
            File.WriteAllText(fixtureSvgPath, fixture.Svg);

            var originalCenterX = tempShape.Left + tempShape.Width / 2f;
            var originalCenterY = tempShape.Top + tempShape.Height / 2f;
            tempShape.Select();
            WinForms.Application.DoEvents();
            var service = new PowerPointFormulaService(application);
            var selected = service.ReadSelection();
            if (selected.Metadata is null)
                throw new InvalidDataException("Current-source service could not read copied bad OLE metadata.");
            var persistedFont = FormulaFontSize.ResolveSemanticFontSize(metadata);
            var readFont = FormulaFontSize.ResolveSemanticFontSize(selected.Metadata);
            AssertNear(persistedFont, readFont, 0.1f,
                "Corrupted PowerPoint OLE geometry polluted the semantic font size.");
            Console.WriteLine(
                $"Current-source ReadSelection kept semantic font at {readFont:F1} pt despite " +
                $"{tempShape.Width:F2}x{tempShape.Height:F2} pt corrupted host geometry.");

            var semanticScale = FormulaFontSize.ResolveSemanticFontSize(metadata)
                / Math.Max(0.5f, FormulaFontSize.ResolveRenderFontSize(metadata));
            var expectedWidth = fixture.RenderWidth * 0.75f * semanticScale;
            var expectedHeight = fixture.RenderHeight * 0.75f * semanticScale;
            var toPicture = CreatePowerPointGeometrySession(
                "edit",
                FormulaOleContract.CrossPlatformPictureMode,
                selected.Metadata.FormulaId,
                tempShape.Name,
                selected.Metadata,
                fixture);
            var pictureResult = service.Replace(toPicture, fixtureSvgPath);
            Release(tempShape);
            tempShape = tempSlide.Shapes[pictureResult.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(tempShape),
                "Corrupted OLE recovery did not produce an SVG picture.");
            AssertNear(expectedWidth, tempShape.Width, 1.5f,
                "Corrupted OLE width was propagated instead of recovering metadata geometry.");
            AssertNear(expectedHeight, tempShape.Height, 1.5f,
                "Corrupted OLE height was propagated instead of recovering metadata geometry.");
            AssertNear(originalCenterX, tempShape.Left + tempShape.Width / 2f, 1.0f,
                "Corrupted OLE recovery did not preserve the formula center X.");
            AssertNear(originalCenterY, tempShape.Top + tempShape.Height / 2f, 1.0f,
                "Corrupted OLE recovery did not preserve the formula center Y.");
            Console.WriteLine(
                $"Corrupted OLE self-recovery passed: {sourceShape.Width:F2}x{sourceShape.Height:F2} pt -> " +
                $"{tempShape.Width:F2}x{tempShape.Height:F2} pt SVG, semantic font stayed {readFont:F1} pt.");
        }
        finally
        {
            Release(oleObject);
            Release(format);
            Release(tags);
            Release(tempShape);
            Release(pastedRange);
            Release(tempSlide);
            if (tempPresentation is not null)
            {
                try { tempPresentation.Close(); } catch { }
            }
            Release(tempPresentation);
            Release(sourceShape);
            Release(sourceShapes);
            Release(sourceSlide);
            Release(sourcePresentation);
            Release(application);
            if (!string.IsNullOrWhiteSpace(fixtureSvgPath))
            {
                try { File.Delete(fixtureSvgPath); } catch { }
            }
            ForceComCleanup();
        }
    }

    private static void RunPowerPointCurrentGeometryProbe()
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slides? slides = null;
        try
        {
            application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            presentation = application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            Console.WriteLine(
                $"PowerPoint current geometry probe: {presentation.Name}, slides={presentation.Slides.Count}, " +
                $"page={presentation.PageSetup.SlideWidth:F2}x{presentation.PageSetup.SlideHeight:F2} pt.");
            slides = presentation.Slides;
            for (var slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                PowerPoint.Slide? slide = null;
                PowerPoint.Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        PowerPoint.Shape? shape = null;
                        PowerPoint.Tags? tags = null;
                        PowerPoint.OLEFormat? format = null;
                        object? oleObject = null;
                        try
                        {
                            shape = shapes[shapeIndex];
                            string? encoded = null;
                            try
                            {
                                tags = shape.Tags;
                                encoded = tags["VisualTeXMetadata"];
                            }
                            catch { }
                            var metadata = FormulaMetadataCodec.Decode(encoded);
                            var visualTeXName = shape.Name.StartsWith("VisualTeX_", StringComparison.OrdinalIgnoreCase);
                            if (metadata is null && !visualTeXName) continue;

                            string? progId = null;
                            var extentText = "n/a";
                            if (shape.Type is MsoShapeType.msoEmbeddedOLEObject or MsoShapeType.msoLinkedOLEObject)
                            {
                                try
                                {
                                    format = shape.OLEFormat;
                                    progId = format.ProgID;
                                    oleObject = format.Object;
                                    if (oleObject is GeometryOleObjectNative nativeOle)
                                    {
                                        var result = nativeOle.GetExtent(1, out var extent);
                                        if (result >= 0)
                                        {
                                            var widthPt = extent.Cx * 72f / 2540f;
                                            var heightPt = extent.Cy * 72f / 2540f;
                                            extentText = $"{widthPt:F2}x{heightPt:F2}pt ({extent.Cx}x{extent.Cy} HIMETRIC)";
                                        }
                                        else
                                        {
                                            extentText = $"HRESULT=0x{result:X8}";
                                        }
                                    }
                                }
                                catch (Exception error)
                                {
                                    extentText = "error=" + error.Message;
                                }
                            }

                            Console.WriteLine(
                                $"  slide={slideIndex} shape={shapeIndex} name={shape.Name} type={shape.Type} " +
                                $"left={shape.Left:F2} top={shape.Top:F2} size={shape.Width:F2}x{shape.Height:F2}pt " +
                                $"progId={progId ?? "n/a"} extent={extentText} " +
                                $"render={metadata?.RenderWidthPx:F3}x{metadata?.RenderHeightPx:F3}px " +
                                $"font={metadata?.FontSizePt:F1} formulaId={metadata?.FormulaId ?? "n/a"}");
                        }
                        finally
                        {
                            Release(oleObject);
                            Release(format);
                            Release(tags);
                            Release(shape);
                        }
                    }
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }
        }
        finally
        {
            Release(slides);
            Release(presentation);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPointOleSvgGeometryAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var officeTempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(officeTempRoot);
        var svgPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        PowerPoint.Application? application = null;
        var ownsApplication = false;
        Process? testOleServerProcess = null;
        object? oleServerKeepAlive = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        try
        {
            var testOleServerPath = Environment.GetEnvironmentVariable("VISUALTEX_TEST_OLE_SERVER_PATH");
            if (!string.IsNullOrWhiteSpace(testOleServerPath))
            {
                testOleServerProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = testOleServerPath,
                    Arguments = "-Embedding",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }) ?? throw new InvalidOperationException("Failed to start the test VisualTeX OLE server.");
                Thread.Sleep(500);
                if (testOleServerProcess.HasExited)
                    throw new InvalidOperationException("The test VisualTeX OLE server exited before COM activation.");
                var oleServerType = Type.GetTypeFromProgID(FormulaOleContract.ProgId, throwOnError: true)
                    ?? throw new InvalidOperationException("VisualTeX OLE server class is not registered.");
                oleServerKeepAlive = Activator.CreateInstance(oleServerType)
                    ?? throw new InvalidOperationException("VisualTeX OLE server keep-alive could not be created.");
                Console.WriteLine(
                    $"  Test OLE server pinned: pid={testOleServerProcess.Id}, path={testOleServerPath}");
            }

            var fixture = LoadPowerPointGeometryFixtureFromEnvironment();
            File.WriteAllText(svgPath, fixture.Svg);
            if (!string.IsNullOrWhiteSpace(fixture.PngBase64))
                File.WriteAllBytes(pngPath, Convert.FromBase64String(fixture.PngBase64));
            else
                WriteAcceptancePng(
                    pngPath,
                    fixture.Latex,
                    Math.Max(32, (int)Math.Ceiling(fixture.RenderWidth * 2)),
                    Math.Max(24, (int)Math.Ceiling(fixture.RenderHeight * 2)));
            emfPath = CreatePowerPointAcceptanceEmf(
                svgPath,
                fixture.RenderWidth,
                fixture.RenderHeight);
            AssertEmfPhysicalFrame(
                emfPath,
                fixture.RenderWidth,
                fixture.RenderHeight);
            var retainedEmfPath = Path.Combine(artifactRoot, "real-session-preview.emf");
            File.Copy(emfPath, retainedEmfPath, overwrite: true);
            var directEmfPngPath = Path.Combine(artifactRoot, "real-session-emf-direct.png");
            RenderEmfDirectly(
                emfPath,
                directEmfPngPath,
                Math.Max(32, (int)Math.Ceiling(fixture.RenderWidth)),
                Math.Max(24, (int)Math.Ceiling(fixture.RenderHeight)));
            var directEmfInk = AnalyzeVisibleDarkPixels(directEmfPngPath);
            Console.WriteLine(
                $"  Direct EMF replay ink: {directEmfInk.Width}x{directEmfInk.Height}/{directEmfInk.Count} " +
                $"on {directEmfInk.ImageWidth}x{directEmfInk.ImageHeight}.");
            var win32EmfPngPath = Path.Combine(artifactRoot, "real-session-emf-playenhmetafile.png");
            RenderEmfWithPlayEnhMetaFile(
                emfPath,
                win32EmfPngPath,
                Math.Max(32, (int)Math.Ceiling(fixture.RenderWidth)),
                Math.Max(24, (int)Math.Ceiling(fixture.RenderHeight)));
            var win32EmfInk = AnalyzeVisibleDarkPixels(win32EmfPngPath);
            Console.WriteLine(
                $"  Win32 PlayEnhMetaFile ink: {win32EmfInk.Width}x{win32EmfInk.Height}/{win32EmfInk.Count} " +
                $"on {win32EmfInk.ImageWidth}x{win32EmfInk.ImageHeight}.");

            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {
                application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
                ownsApplication = true;
            }
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var postedUiWork = new Queue<Action>();
            var delayedUiWork = new Queue<(Action Operation, int DelayMilliseconds)>();
            var captureOleStages = true;
            var oleStageOrdinal = 0;
            var sawHiddenCreatedStage = false;
            var sawHiddenInitializedStage = false;
            var sawVisibleFinalizedStage = false;
            var service = new PowerPointFormulaService(
                application,
                operation => postedUiWork.Enqueue(operation),
                (operation, delayMilliseconds) => delayedUiWork.Enqueue((operation, delayMilliseconds)),
                (stage, stageShape) =>
                {
                    if (!captureOleStages) return;
                    oleStageOrdinal++;
                    var visible = stageShape.Visible;
                    var stageExtent = ReadOleExtentPoints(stageShape);
                    if (string.Equals(stage, "created", StringComparison.Ordinal))
                    {
                        AssertEqual(MsoTriState.msoFalse, visible,
                            "New PowerPoint OLE became visible before initialization.");
                        sawHiddenCreatedStage = true;
                    }
                    else if (string.Equals(stage, "initialized", StringComparison.Ordinal))
                    {
                        AssertEqual(MsoTriState.msoFalse, visible,
                            "PowerPoint exposed the transient OLE presentation reflow during initialization.");
                        sawHiddenInitializedStage = true;
                    }
                    else if (string.Equals(stage, "finalized", StringComparison.Ordinal))
                    {
                        AssertEqual(MsoTriState.msoTrue, visible,
                            "PowerPoint OLE did not become visible after final geometry restoration.");
                        sawVisibleFinalizedStage = true;
                    }

                    if (visible == MsoTriState.msoFalse)
                    {
                        Console.WriteLine(
                            $"  OLE synchronous stage {oleStageOrdinal:D2} {stage}: hidden, " +
                            $"shape={stageShape.Width:F2}x{stageShape.Height:F2}pt, server={stageExtent}.");
                    }
                    else
                    {
                        var stagePath = Path.Combine(
                            artifactRoot,
                            $"real-session-ole-stage-{oleStageOrdinal:D2}-{stage}.png");
                        stageShape.Export(stagePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                        var stageInk = AnalyzeVisibleDarkPixels(stagePath, margin: 4);
                        Console.WriteLine(
                            $"  OLE synchronous stage {oleStageOrdinal:D2} {stage}: visible, " +
                            $"shape={stageShape.Width:F2}x{stageShape.Height:F2}pt, " +
                            $"server={stageExtent}, ink={stageInk.Width}x{stageInk.Height}/{stageInk.Count}, " +
                            $"canvas={stageInk.ImageWidth}x{stageInk.ImageHeight}.");
                    }
                    if (string.Equals(stage, "finalized", StringComparison.Ordinal))
                        captureOleStages = false;
                });
            void DrainPostedUiWork()
            {
                while (postedUiWork.Count > 0)
                    postedUiWork.Dequeue()();
            }
            void DrainDelayedUiWork()
            {
                while (delayedUiWork.Count > 0)
                    delayedUiWork.Dequeue().Operation();
            }
            var formulaId = Guid.NewGuid().ToString();

            var create = CreatePowerPointGeometrySession(
                "create",
                FormulaOleContract.CrossPlatformPictureMode,
                formulaId,
                null,
                null,
                fixture);
            var inserted = service.Insert(create, svgPath);
            shape = slide.Shapes[inserted.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                "PowerPoint geometry fixture did not start as an SVG picture.");
            var baselineLeft = shape.Left;
            var baselineTop = shape.Top;
            var baselineWidth = shape.Width;
            var baselineHeight = shape.Height;
            var baselineAspect = baselineWidth / baselineHeight;
            var sourceExportPath = Path.Combine(artifactRoot, "real-session-source.png");
            shape.Export(sourceExportPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var sourceInk = AnalyzeVisibleDarkPixels(sourceExportPath);
            AssertTrue(sourceInk.Count > 0, "Real Session source SVG exported no visible formula pixels.");
            Console.WriteLine(
                $"  Real Session SVG ink: {sourceInk.Width}x{sourceInk.Height}/{sourceInk.Count} " +
                $"on {sourceInk.ImageWidth}x{sourceInk.ImageHeight}.");
            AssertNear(fixture.RenderWidth * 0.75f, baselineWidth, 1.5f,
                "PowerPoint geometry fixture width does not match its 96-dpi natural size.");
            AssertNear(fixture.RenderHeight * 0.75f, baselineHeight, 1.5f,
                "PowerPoint geometry fixture height does not match its 96-dpi natural size.");

            for (var iteration = 1; iteration <= 20; iteration++)
            {
                shape.Select(MsoTriState.msoFalse);
                WinForms.Application.DoEvents();
                var metadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"Picture metadata disappeared before iteration {iteration}.");
                var toOle = CreatePowerPointGeometrySession(
                    "edit",
                    FormulaOleContract.NativeOleMode,
                    formulaId,
                    shape.Name,
                    metadata,
                    fixture);
                var oleResult = service.ReplaceOle(toOle, pngPath, emfPath);
                Release(shape);
                shape = slide.Shapes[oleResult.ObjectId];
                if (iteration == 1)
                {
                    AssertTrue(sawHiddenCreatedStage,
                        "PowerPoint OLE regression did not observe the hidden creation stage.");
                    AssertTrue(sawHiddenInitializedStage,
                        "PowerPoint OLE regression did not observe the hidden initialized stage.");
                    AssertTrue(sawVisibleFinalizedStage,
                        "PowerPoint OLE regression did not observe the visible finalized stage.");
                    // Reproduce the two real host geometries observed in the
                    // installed PowerPoint add-in after its OLE callback unwound.
                    var centerX = baselineLeft + baselineWidth / 2f;
                    var centerY = baselineTop + baselineHeight / 2f;
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = 546.375f;
                    shape.Height = 210.875f;
                    shape.Left = centerX - shape.Width / 2f;
                    shape.Top = centerY - shape.Height / 2f;
                    AssertTrue(postedUiWork.Count > 0,
                        "PowerPoint OLE conversion did not schedule deferred geometry repair.");
                    postedUiWork.Dequeue()();
                    AssertNear(baselineWidth, shape.Width, 1.5f,
                        "First deferred OLE geometry repair did not restore width.");
                    AssertNear(baselineHeight, shape.Height, 1.5f,
                        "First deferred OLE geometry repair did not restore height.");
                    AssertNear(baselineLeft, shape.Left, 0.5f,
                        "First deferred OLE geometry repair did not restore left position.");
                    AssertNear(baselineTop, shape.Top, 0.5f,
                        "First deferred OLE geometry repair did not restore top position.");

                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = 1504.38f;
                    shape.Height = 719.50f;
                    shape.Left = centerX - shape.Width / 2f;
                    shape.Top = centerY - shape.Height / 2f;
                    DrainPostedUiWork();
                    AssertNear(baselineWidth, shape.Width, 1.5f,
                        "Second deferred OLE geometry repair did not restore width.");
                    AssertNear(baselineHeight, shape.Height, 1.5f,
                        "Second deferred OLE geometry repair did not restore height.");
                    AssertNear(baselineLeft, shape.Left, 0.5f,
                        "Second deferred OLE geometry repair did not restore left position.");
                    AssertNear(baselineTop, shape.Top, 0.5f,
                        "Second deferred OLE geometry repair did not restore top position.");

                    // The installed 175%-DPI PowerPoint build reproduced a later
                    // container reflow after both immediate UI posts had already
                    // completed: 99.29x32.43 pt became 136.25x57.50 pt while the
                    // OLE server extent remained near the original natural size.
                    // Simulate that late replay and require a delayed UI repair to
                    // restore both the box and its center.
                    AssertTrue(delayedUiWork.Count >= 3,
                        "PowerPoint OLE conversion did not schedule delayed geometry rechecks.");
                    shape.LockAspectRatio = MsoTriState.msoFalse;
                    shape.Width = 136.25f;
                    shape.Height = 57.50f;
                    shape.Left = centerX - shape.Width / 2f;
                    shape.Top = centerY - shape.Height / 2f;
                    DrainDelayedUiWork();
                    AssertNear(baselineWidth, shape.Width, 1.5f,
                        "Delayed OLE reflow repair did not restore width.");
                    AssertNear(baselineHeight, shape.Height, 1.5f,
                        "Delayed OLE reflow repair did not restore height.");
                    AssertNear(baselineLeft, shape.Left, 0.5f,
                        "Delayed OLE reflow repair did not restore left position.");
                    AssertNear(baselineTop, shape.Top, 0.5f,
                        "Delayed OLE reflow repair did not restore top position.");
                }
                else
                {
                    DrainPostedUiWork();
                    DrainDelayedUiWork();
                }
                WaitForPowerPointOleSettle();
                Release(shape);
                shape = slide.Shapes[oleResult.ObjectId];
                AssertPowerPointOle(shape, $"Iteration {iteration} picture -> OLE failed.");
                AssertNear(baselineLeft, shape.Left, 0.5f,
                    $"Iteration {iteration} picture -> OLE left position drifted.");
                AssertNear(baselineTop, shape.Top, 0.5f,
                    $"Iteration {iteration} picture -> OLE top position drifted.");
                AssertNear(baselineWidth, shape.Width, 1.5f,
                    $"Iteration {iteration} picture -> OLE width drifted.");
                AssertNear(baselineHeight, shape.Height, 1.5f,
                    $"Iteration {iteration} picture -> OLE height drifted.");
                AssertNear(baselineAspect, shape.Width / shape.Height, 0.02f,
                    $"Iteration {iteration} picture -> OLE aspect ratio drifted.");
                AssertOleServerExtent(shape, baselineWidth, baselineHeight, iteration);
                if (iteration == 1)
                {
                    var oleExportPath = Path.Combine(artifactRoot, "real-session-ole.png");
                    shape.Export(oleExportPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    var oleInk = AnalyzeVisibleDarkPixels(oleExportPath, margin: 4);
                    AssertTrue(oleInk.Count > 0, "Real Session OLE exported no visible formula pixels inside the OLE border.");
                    var widthDelta = Math.Abs(oleInk.Width - sourceInk.Width) / (double)Math.Max(1, sourceInk.Width);
                    var heightDelta = Math.Abs(oleInk.Height - sourceInk.Height) / (double)Math.Max(1, sourceInk.Height);
                    var sourceInkRatio = sourceInk.Width / (double)Math.Max(1, sourceInk.Height);
                    var oleInkRatio = oleInk.Width / (double)Math.Max(1, oleInk.Height);
                    var ratioDelta = Math.Abs(oleInkRatio - sourceInkRatio) / Math.Max(0.001, sourceInkRatio);
                    var countRatio = oleInk.Count / (double)Math.Max(1, sourceInk.Count);
                    Console.WriteLine(
                        $"  Real Session visual ink: SVG={sourceInk.Width}x{sourceInk.Height}/{sourceInk.Count}, " +
                        $"OLE={oleInk.Width}x{oleInk.Height}/{oleInk.Count}, countRatio={countRatio:F3}, " +
                        $"ratio={sourceInkRatio:F4}->{oleInkRatio:F4}; " +
                        $"OLE canvas={oleInk.ImageWidth}x{oleInk.ImageHeight}.");
                    if (widthDelta > 0.08 || heightDelta > 0.08 || ratioDelta > 0.05
                        || countRatio < 0.55 || countRatio > 1.8)
                        throw new InvalidOperationException(
                            "Real MathJax SVG -> OLE changed or clipped the formula presentation.");

                    for (var settleSample = 1; settleSample <= 40; settleSample++)
                    {
                        WinForms.Application.DoEvents();
                        Thread.Sleep(250);
                        var settledWidth = shape.Width;
                        var settledHeight = shape.Height;
                        if (settleSample == 1 || settleSample % 4 == 0)
                        {
                            Console.WriteLine(
                                $"  OLE delayed-settle +{settleSample * 0.25:F2}s: " +
                                $"{settledWidth:F2}x{settledHeight:F2} pt at ({shape.Left:F2},{shape.Top:F2}).");
                        }
                        if (Math.Abs(settledWidth - baselineWidth) > 1.5f
                            || Math.Abs(settledHeight - baselineHeight) > 1.5f)
                        {
                            throw new InvalidOperationException(
                                $"PowerPoint asynchronously resized the OLE host after conversion: " +
                                $"expected {baselineWidth:F2}x{baselineHeight:F2} pt, " +
                                $"observed {settledWidth:F2}x{settledHeight:F2} pt after {settleSample * 0.25:F2}s.");
                        }
                    }
                }

                shape.Select(MsoTriState.msoFalse);
                WinForms.Application.DoEvents();
                metadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"OLE metadata disappeared before iteration {iteration}.");
                var toPicture = CreatePowerPointGeometrySession(
                    "edit",
                    FormulaOleContract.CrossPlatformPictureMode,
                    formulaId,
                    shape.Name,
                    metadata,
                    fixture);
                var pictureResult = service.Replace(toPicture, svgPath);
                Release(shape);
                shape = slide.Shapes[pictureResult.ObjectId];
                WinForms.Application.DoEvents();
                Thread.Sleep(120);
                Release(shape);
                shape = slide.Shapes[pictureResult.ObjectId];
                AssertTrue(IsPowerPointEditablePictureShape(shape),
                    $"Iteration {iteration} OLE -> picture failed.");
                AssertNear(baselineLeft, shape.Left, 0.5f,
                    $"Iteration {iteration} OLE -> picture left position drifted.");
                AssertNear(baselineTop, shape.Top, 0.5f,
                    $"Iteration {iteration} OLE -> picture top position drifted.");
                AssertNear(baselineWidth, shape.Width, 1.5f,
                    $"Iteration {iteration} OLE -> picture width drifted.");
                AssertNear(baselineHeight, shape.Height, 1.5f,
                    $"Iteration {iteration} OLE -> picture height drifted.");
                AssertNear(baselineAspect, shape.Width / shape.Height, 0.02f,
                    $"Iteration {iteration} OLE -> picture aspect ratio drifted.");
            }

            // Persist a final real-MathJax OLE and verify PowerPoint does not fall
            // back to the initial placeholder after save/reopen.
            var finalPictureMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("Picture metadata disappeared before save/reopen validation.");
            var finalOleSession = CreatePowerPointGeometrySession(
                "edit",
                FormulaOleContract.NativeOleMode,
                formulaId,
                shape.Name,
                finalPictureMetadata,
                fixture);
            var finalOleResult = service.ReplaceOle(finalOleSession, pngPath, emfPath);
            Release(shape);
            shape = slide.Shapes[finalOleResult.ObjectId];
            WaitForPowerPointOleSettle();
            Release(shape);
            shape = slide.Shapes[finalOleResult.ObjectId];
            AssertPowerPointOle(shape, "Final save/reopen fixture did not become OLE.");
            AssertNear(baselineLeft, shape.Left, 0.5f, "Final OLE left position drifted before save.");
            AssertNear(baselineTop, shape.Top, 0.5f, "Final OLE top position drifted before save.");
            AssertNear(baselineWidth, shape.Width, 1.5f, "Final OLE width drifted before save.");
            AssertNear(baselineHeight, shape.Height, 1.5f, "Final OLE height drifted before save.");

            var output = Path.Combine(artifactRoot, "powerpoint-ole-svg-geometry-stable.pptx");
            presentation.SaveAs(
                output,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);

            Release(shape);
            shape = null;
            Release(slide);
            slide = null;
            presentation.Close();
            Release(presentation);
            presentation = null;

            presentation = application.Presentations.Open(
                output,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue);
            slide = presentation.Slides[1];
            shape = slide.Shapes[finalOleResult.ObjectId];
            AssertPowerPointOle(shape, "Reopened real-MathJax formula is no longer OLE.");
            AssertNear(baselineLeft, shape.Left, 0.5f, "Reopened OLE left position drifted.");
            AssertNear(baselineTop, shape.Top, 0.5f, "Reopened OLE top position drifted.");
            AssertNear(baselineWidth, shape.Width, 1.5f, "Reopened OLE width drifted.");
            AssertNear(baselineHeight, shape.Height, 1.5f, "Reopened OLE height drifted.");
            AssertNear(baselineAspect, shape.Width / shape.Height, 0.02f,
                "Reopened OLE aspect ratio drifted.");
            var reopenedOleExportPath = Path.Combine(artifactRoot, "real-session-ole-reopened.png");
            shape.Export(reopenedOleExportPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var reopenedOleInk = AnalyzeVisibleDarkPixels(reopenedOleExportPath, margin: 4);
            AssertTrue(reopenedOleInk.Count > 0,
                "Reopened real-MathJax OLE exported no visible formula pixels inside the OLE border.");
            var reopenedWidthDelta = Math.Abs(reopenedOleInk.Width - sourceInk.Width)
                / (double)Math.Max(1, sourceInk.Width);
            var reopenedHeightDelta = Math.Abs(reopenedOleInk.Height - sourceInk.Height)
                / (double)Math.Max(1, sourceInk.Height);
            var reopenedCountRatio = reopenedOleInk.Count / (double)Math.Max(1, sourceInk.Count);
            Console.WriteLine(
                $"  Reopened real Session visual ink: SVG={sourceInk.Width}x{sourceInk.Height}/{sourceInk.Count}, " +
                $"OLE={reopenedOleInk.Width}x{reopenedOleInk.Height}/{reopenedOleInk.Count}, " +
                $"countRatio={reopenedCountRatio:F3}.");
            if (reopenedWidthDelta > 0.08 || reopenedHeightDelta > 0.08
                || reopenedCountRatio < 0.55 || reopenedCountRatio > 1.8)
                throw new InvalidOperationException(
                    "Saved/reopened real MathJax OLE changed or clipped the formula presentation.");

            Console.WriteLine(
                $"PowerPoint OLE/SVG geometry acceptance passed: real Session formula stayed at " +
                $"({baselineLeft:0.0},{baselineTop:0.0}) with {baselineWidth:0.0}x{baselineHeight:0.0} pt " +
                "across 20 round trips and a saved/reopened OLE retained the full formula presentation; " +
                "no threshold reset or geometry self-heal was used.");
        }
        finally
        {
            Release(shape);
            Release(slide);
            Release(oleServerKeepAlive);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null && ownsApplication)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            if (testOleServerProcess is not null)
            {
                try
                {
                    if (!testOleServerProcess.HasExited)
                        testOleServerProcess.Kill();
                }
                catch { }
                testOleServerProcess.Dispose();
            }
            if (!string.IsNullOrWhiteSpace(emfPath))
            {
                try { File.Delete(emfPath); } catch { }
            }
            try { File.Delete(svgPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            ForceComCleanup();
        }
    }

    private static void RunPowerPointInstalledOlePresentationAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "installed-addin-source.svg");
        var sourcePngPath = Path.Combine(artifactRoot, "installed-addin-source.png");
        var olePngPath = Path.Combine(artifactRoot, "installed-addin-ole.png");
        PowerPoint.Application? application = null;
        var ownsApplication = false;
        Process? testOleServerProcess = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        object? oleServerKeepAlive = null;
        COMAddIns? addIns = null;
        COMAddIn? installedAddIn = null;
        try
        {
            var testOleServerPath = Environment.GetEnvironmentVariable("VISUALTEX_TEST_OLE_SERVER_PATH");
            if (!string.IsNullOrWhiteSpace(testOleServerPath))
            {
                testOleServerProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = testOleServerPath,
                    Arguments = "-Embedding",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }) ?? throw new InvalidOperationException("Failed to start the test VisualTeX OLE server.");
                Thread.Sleep(500);
                if (testOleServerProcess.HasExited)
                    throw new InvalidOperationException("The test VisualTeX OLE server exited before COM activation.");
                Console.WriteLine(
                    $"Test OLE server started in-process: pid={testOleServerProcess.Id}, path={testOleServerPath}");
            }

            var oleServerType = Type.GetTypeFromProgID(FormulaOleContract.ProgId, throwOnError: true)
                ?? throw new InvalidOperationException("VisualTeX OLE server class is not registered.");
            oleServerKeepAlive = Activator.CreateInstance(oleServerType)
                ?? throw new InvalidOperationException("VisualTeX OLE server keep-alive could not be created.");

            var formulaId = Guid.NewGuid().ToString();
            var realFixture = LoadPowerPointGeometryFixtureFromEnvironment();
            var fixture = CreatePowerPointGeometrySession(
                "create",
                FormulaOleContract.CrossPlatformPictureMode,
                formulaId,
                null,
                null,
                realFixture);
            var metadata = fixture.ToMetadata();
            var encoded = FormulaMetadataCodec.Encode(metadata);
            File.WriteAllText(svgPath, realFixture.Svg);

            try
            {
                application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {
                application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
                ownsApplication = true;
            }
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var baselineLeft = 300f;
            var baselineTop = 200f;
            var baselineWidth = realFixture.RenderWidth * 0.75f;
            var baselineHeight = realFixture.RenderHeight * 0.75f;
            var baselineRatio = baselineWidth / baselineHeight;

            shape = slide.Shapes.AddPicture(
                svgPath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                baselineLeft,
                baselineTop,
                baselineWidth,
                baselineHeight);
            shape.LockAspectRatio = MsoTriState.msoFalse;
            shape.Width = baselineWidth;
            shape.Height = baselineHeight;
            shape.Left = baselineLeft;
            shape.Top = baselineTop;
            shape.LockAspectRatio = MsoTriState.msoTrue;
            shape.Name = $"VisualTeX_{formulaId}";
            shape.AlternativeText = encoded;
            PowerPoint.Tags? tags = null;
            try
            {
                tags = shape.Tags;
                tags.Add("VisualTeXFormulaId", formulaId);
                tags.Add("VisualTeXMetadata", encoded);
            }
            finally { Release(tags); }

            shape.Export(sourcePngPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var sourceInk = AnalyzeDarkPixels(sourcePngPath);
            AssertTrue(sourceInk.Count > 0, "Source SVG export contains no formula pixels.");

            addIns = application.COMAddIns;
            object addInIndex = "VisualTeX.PowerPointVsto";
            installedAddIn = addIns.Item(ref addInIndex);
            if (!installedAddIn.Connect)
            {
                installedAddIn.Connect = true;
                for (var index = 0; index < 20 && installedAddIn.Object is null; index++)
                {
                    WinForms.Application.DoEvents();
                    Thread.Sleep(100);
                }
            }
            dynamic callbacks = installedAddIn.Object
                ?? throw new InvalidOperationException("Installed PowerPoint add-in automation object is unavailable.");

            for (var iteration = 1; iteration <= 10; iteration++)
            {
                shape.Select();
                callbacks.OnConvertSelected(null);
                Release(shape);
                shape = WaitForInstalledPowerPointFormulaShape(
                    slide,
                    formulaId,
                    MsoShapeType.msoEmbeddedOLEObject,
                    TimeSpan.FromSeconds(15));
                Thread.Sleep(250);
                WinForms.Application.DoEvents();

                if (iteration == 1)
                {
                    var convertedMetadata = DecodePowerPointMetadata(shape)
                        ?? throw new InvalidDataException("Installed OLE lost VisualTeX metadata.");
                    PowerPoint.OLEFormat? diagnosticFormat = null;
                    object? diagnosticOleObject = null;
                    try
                    {
                        diagnosticFormat = shape.OLEFormat;
                        diagnosticOleObject = diagnosticFormat.Object;
                        var serverExtentText = "unavailable";
                        if (diagnosticOleObject is GeometryOleObjectNative nativeOle
                            && nativeOle.GetExtent(1, out var extent) >= 0)
                        {
                            serverExtentText = string.Format(
                                System.Globalization.CultureInfo.InvariantCulture,
                                "{0:F2}x{1:F2}pt",
                                extent.Cx * 72.0 / 2540.0,
                                extent.Cy * 72.0 / 2540.0);
                        }
                        Console.WriteLine(
                            $"  Installed first OLE diagnostics: shape={shape.Width:F2}x{shape.Height:F2}pt, " +
                            $"serverExtent={serverExtentText}, metadataRender=" +
                            $"{convertedMetadata.RenderWidthPx:F3}x{convertedMetadata.RenderHeightPx:F3}px, " +
                            $"font={convertedMetadata.FontSizePt:F1}/{convertedMetadata.RenderFontSizePt:F1}pt.");
                    }
                    finally
                    {
                        Release(diagnosticOleObject);
                        Release(diagnosticFormat);
                    }

                    shape.Export(olePngPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    var oleInk = AnalyzeDarkPixels(olePngPath);
                    AssertTrue(oleInk.Count > 0, "OLE export contains no formula pixels.");
                    var sourceInkRatio = sourceInk.Width / (double)Math.Max(1, sourceInk.Height);
                    var oleInkRatio = oleInk.Width / (double)Math.Max(1, oleInk.Height);
                    Console.WriteLine(
                        $"  Installed first visual ink: SVG={sourceInk.Width}x{sourceInk.Height}/{sourceInk.Count} " +
                        $"ratio={sourceInkRatio:F4}; OLE={oleInk.Width}x{oleInk.Height}/{oleInk.Count} " +
                        $"ratio={oleInkRatio:F4}.");
                    if (Math.Abs(sourceInkRatio - oleInkRatio) / sourceInkRatio > 0.03)
                        throw new InvalidOperationException(
                            $"SVG -> OLE visually distorted the formula. Source ink ratio={sourceInkRatio:F4}, " +
                            $"OLE ink ratio={oleInkRatio:F4}.");

                    for (var settleSample = 1; settleSample <= 40; settleSample++)
                    {
                        WinForms.Application.DoEvents();
                        Thread.Sleep(250);
                        if (settleSample == 1 || settleSample % 4 == 0)
                        {
                            Console.WriteLine(
                                $"  Installed add-in delayed-settle +{settleSample * 0.25:F2}s: " +
                                $"L={shape.Left:F2} T={shape.Top:F2} W={shape.Width:F2} H={shape.Height:F2}");
                        }
                        AssertInstalledPowerPointGeometry(
                            shape,
                            baselineLeft,
                            baselineTop,
                            baselineWidth,
                            baselineHeight,
                            baselineRatio,
                            $"Installed add-in delayed-settle +{settleSample * 0.25:F2}s");
                    }
                }

                AssertInstalledPowerPointGeometry(
                    shape,
                    baselineLeft,
                    baselineTop,
                    baselineWidth,
                    baselineHeight,
                    baselineRatio,
                    $"Iteration {iteration} SVG -> OLE");

                shape.Select();
                callbacks.OnExportSelectedAsPicture(null);
                Release(shape);
                shape = WaitForInstalledPowerPointFormulaShape(
                    slide,
                    formulaId,
                    (MsoShapeType)28,
                    TimeSpan.FromSeconds(15));
                Thread.Sleep(250);
                WinForms.Application.DoEvents();
                AssertInstalledPowerPointGeometry(
                    shape,
                    baselineLeft,
                    baselineTop,
                    baselineWidth,
                    baselineHeight,
                    baselineRatio,
                    $"Iteration {iteration} OLE -> SVG");
            }

            shape.Select();
            callbacks.OnConvertSelected(null);
            Release(shape);
            shape = WaitForInstalledPowerPointFormulaShape(
                slide,
                formulaId,
                MsoShapeType.msoEmbeddedOLEObject,
                TimeSpan.FromSeconds(15));
            Thread.Sleep(250);
            WinForms.Application.DoEvents();
            AssertInstalledPowerPointGeometry(
                shape,
                baselineLeft,
                baselineTop,
                baselineWidth,
                baselineHeight,
                baselineRatio,
                "Save/reopen source OLE");

            var persistedPath = Path.Combine(artifactRoot, "installed-addin-ole-persistence.pptx");
            presentation.SaveAs(
                persistedPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);
            Release(shape);
            shape = null;
            Release(slide);
            slide = null;
            presentation.Close();
            Release(presentation);
            presentation = application.Presentations.Open(
                persistedPath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue);
            slide = presentation.Slides[1];
            application.ActiveWindow.View.GotoSlide(1);
            shape = WaitForInstalledPowerPointFormulaShape(
                slide,
                formulaId,
                MsoShapeType.msoEmbeddedOLEObject,
                TimeSpan.FromSeconds(15));
            Thread.Sleep(500);
            WinForms.Application.DoEvents();
            AssertInstalledPowerPointGeometry(
                shape,
                baselineLeft,
                baselineTop,
                baselineWidth,
                baselineHeight,
                baselineRatio,
                "Reopened persisted OLE");

            var reopenedOlePngPath = Path.Combine(artifactRoot, "installed-addin-reopened-ole.png");
            shape.Export(reopenedOlePngPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var reopenedInk = AnalyzeDarkPixels(reopenedOlePngPath);
            AssertTrue(reopenedInk.Count > 0, "Reopened OLE export contains no formula pixels.");
            var reopenedSourceInkRatio = sourceInk.Width / (double)Math.Max(1, sourceInk.Height);
            var reopenedOleInkRatio = reopenedInk.Width / (double)Math.Max(1, reopenedInk.Height);
            if (Math.Abs(reopenedSourceInkRatio - reopenedOleInkRatio) / reopenedSourceInkRatio > 0.03)
                throw new InvalidOperationException(
                    $"Reopened OLE visually distorted the formula. Source ink ratio={reopenedSourceInkRatio:F4}, " +
                    $"OLE ink ratio={reopenedOleInkRatio:F4}.");

            shape.Select();
            callbacks.OnExportSelectedAsPicture(null);
            Release(shape);
            shape = WaitForInstalledPowerPointFormulaShape(
                slide,
                formulaId,
                (MsoShapeType)28,
                TimeSpan.FromSeconds(15));
            Thread.Sleep(250);
            WinForms.Application.DoEvents();
            AssertInstalledPowerPointGeometry(
                shape,
                baselineLeft,
                baselineTop,
                baselineWidth,
                baselineHeight,
                baselineRatio,
                "Reopened OLE -> SVG");

            Console.WriteLine(
                "Installed PowerPoint add-in OLE presentation acceptance passed: 10 SVG/OLE round trips " +
                "kept Left/Top/Width/Height/aspect exactly stable, the first OLE visual ink ratio matched the SVG, " +
                "and a saved/reopened OLE preserved geometry, visual aspect, and OLE -> SVG conversion.");
        }
        finally
        {
            Release(shape);
            Release(oleServerKeepAlive);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            Release(installedAddIn);
            Release(addIns);
            if (ownsApplication && application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            if (testOleServerProcess is not null)
            {
                try
                {
                    if (!testOleServerProcess.HasExited)
                        testOleServerProcess.Kill();
                }
                catch { }
                testOleServerProcess.Dispose();
            }
            ForceComCleanup();
        }
    }

    private static PowerPoint.Shape WaitForInstalledPowerPointFormulaShape(
        PowerPoint.Slide slide,
        string formulaId,
        MsoShapeType expectedType,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            PowerPoint.Shapes? shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape? candidate = null;
                    try
                    {
                        candidate = shapes[index];
                        try
                        {
                            if (candidate.Type != expectedType) continue;
                            var candidateMetadata = DecodePowerPointMetadata(candidate);
                            if (!string.Equals(candidateMetadata?.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                                continue;
                        }
                        catch (COMException error) when ((uint)error.HResult == 0x800A01A8u)
                        {
                            // PowerPoint can leave a just-deleted Shape RCW visible
                            // for one enumeration turn while replacing SVG/OLE.
                            continue;
                        }
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            finally { Release(shapes); }
            Thread.Sleep(100);
        }
        throw new TimeoutException(
            $"Timed out waiting for PowerPoint formula {formulaId} with shape type {expectedType}.");
    }

    private static void AssertInstalledPowerPointGeometry(
        PowerPoint.Shape shape,
        float expectedLeft,
        float expectedTop,
        float expectedWidth,
        float expectedHeight,
        float expectedRatio,
        string stage)
    {
        Console.WriteLine(
            $"  {stage}: L={shape.Left:F3} T={shape.Top:F3} W={shape.Width:F3} H={shape.Height:F3} " +
            $"ratio={shape.Width / shape.Height:F5}");
        AssertNear(expectedLeft, shape.Left, 0.2f, $"{stage} moved horizontally.");
        AssertNear(expectedTop, shape.Top, 0.2f, $"{stage} moved vertically.");
        AssertNear(expectedWidth, shape.Width, 0.2f, $"{stage} changed width.");
        AssertNear(expectedHeight, shape.Height, 0.2f, $"{stage} changed height.");
        AssertNear(expectedRatio, shape.Width / shape.Height, 0.01f, $"{stage} changed aspect ratio.");
    }

    private static OfficeSessionDocument CreatePowerPointGeometrySession(
        string mode,
        string objectMode,
        string formulaId,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata,
        PowerPointGeometryFixture? fixture = null)
    {
        fixture ??= new PowerPointGeometryFixture(
            PowerPointGeometryLatex,
            CreateSvg(PowerPointGeometryRenderWidth, PowerPointGeometryRenderHeight),
            null,
            PowerPointGeometryRenderWidth,
            PowerPointGeometryRenderHeight,
            97.6f,
            PowerPointGeometryFontSize);
        var svg = fixture.Svg;
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString(),
            Mode = mode,
            Host = "powerpoint",
            FormulaId = formulaId,
            SourceDocumentId = null,
            SourceObjectId = sourceObjectId,
            Title = "PowerPoint geometry acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = fixture.Latex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = objectMode,
            Numbered = false,
            FontSizePt = fixture.FontSizePt,
            OriginalMetadata = originalMetadata,
            ExportResult = new OfficeExportDocument
            {
                Svg = svg,
                SvgBase64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)),
                Width = fixture.RenderWidth,
                Height = fixture.RenderHeight,
                Baseline = fixture.Baseline,
            },
        };
    }

    private static void RenderEmfDirectly(
        string emfPath,
        string pngPath,
        int width,
        int height)
    {
        using var metafile = new System.Drawing.Imaging.Metafile(emfPath);
        using var bitmap = new System.Drawing.Bitmap(
            width,
            height,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(System.Drawing.Color.White);
        graphics.DrawImage(metafile, new System.Drawing.Rectangle(0, 0, width, height));
        bitmap.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
    }

    private static void RenderEmfWithPlayEnhMetaFile(
        string emfPath,
        string pngPath,
        int width,
        int height)
    {
        var bytes = File.ReadAllBytes(emfPath);
        var metafile = SetEnhMetaFileBits((uint)bytes.Length, bytes);
        if (metafile == IntPtr.Zero)
            throw new InvalidOperationException($"SetEnhMetaFileBits failed with {Marshal.GetLastWin32Error()}.");
        try
        {
            using var bitmap = new System.Drawing.Bitmap(
                width,
                height,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var graphics = System.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(System.Drawing.Color.White);
            var hdc = graphics.GetHdc();
            try
            {
                var bounds = new NativeRect { Left = 0, Top = 0, Right = width, Bottom = height };
                if (!PlayEnhMetaFile(hdc, metafile, ref bounds))
                    throw new InvalidOperationException($"PlayEnhMetaFile failed with {Marshal.GetLastWin32Error()}.");
            }
            finally { graphics.ReleaseHdc(hdc); }
            bitmap.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
        }
        finally { DeleteEnhMetaFile(metafile); }
    }

    private static (int Count, int Width, int Height, int ImageWidth, int ImageHeight)
        AnalyzeVisibleDarkPixels(string path, int margin = 0)
    {
        using var bitmap = new System.Drawing.Bitmap(path);
        margin = Math.Max(0, Math.Min(margin, Math.Min(bitmap.Width, bitmap.Height) / 3));
        var count = 0;
        var minimumX = bitmap.Width;
        var minimumY = bitmap.Height;
        var maximumX = -1;
        var maximumY = -1;
        for (var y = margin; y < bitmap.Height - margin; y++)
        for (var x = margin; x < bitmap.Width - margin; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.A < 16) continue;
            if (pixel.R >= 160 && pixel.G >= 160 && pixel.B >= 160) continue;
            count++;
            minimumX = Math.Min(minimumX, x);
            minimumY = Math.Min(minimumY, y);
            maximumX = Math.Max(maximumX, x);
            maximumY = Math.Max(maximumY, y);
        }
        return maximumX < minimumX || maximumY < minimumY
            ? (0, 0, 0, bitmap.Width, bitmap.Height)
            : (count, maximumX - minimumX + 1, maximumY - minimumY + 1, bitmap.Width, bitmap.Height);
    }

    private static void WaitForPowerPointOleSettle()
    {
        for (var index = 0; index < 5; index++)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(80);
        }
    }

    private static string ReadOleExtentPoints(PowerPoint.Shape shape)
    {
        PowerPoint.OLEFormat? format = null;
        object? value = null;
        try
        {
            format = shape.OLEFormat;
            value = format.Object;
            if (value is not GeometryOleObjectNative nativeOle) return "n/a";
            var result = nativeOle.GetExtent(1, out var extent);
            if (result < 0) return $"hr=0x{result:X8}";
            var width = extent.Cx * 72f / 2540f;
            var height = extent.Cy * 72f / 2540f;
            return $"{width:F2}x{height:F2}pt";
        }
        catch (Exception error)
        {
            return error.GetType().Name;
        }
        finally
        {
            Release(value);
            Release(format);
        }
    }

    private static void AssertOleServerExtent(
        PowerPoint.Shape shape,
        float expectedWidth,
        float expectedHeight,
        int iteration)
    {
        PowerPoint.OLEFormat? format = null;
        object? value = null;
        try
        {
            format = shape.OLEFormat;
            value = format.Object;
            if (value is not GeometryOleObjectNative nativeOle)
                throw new InvalidOperationException("VisualTeX OLE server does not expose IOleObject.");
            var result = nativeOle.GetExtent(1, out var extent);
            if (result < 0) Marshal.ThrowExceptionForHR(result);
            var width = extent.Cx * 72f / 2540f;
            var height = extent.Cy * 72f / 2540f;
            AssertNear(expectedWidth, width, 1.5f,
                $"OLE server extent width drifted at iteration {iteration}.");
            AssertNear(expectedHeight, height, 1.5f,
                $"OLE server extent height drifted at iteration {iteration}.");
        }
        finally
        {
            Release(value);
            Release(format);
        }
    }

    private static void AssertEmfPhysicalFrame(
        string emfPath,
        float widthPixels,
        float heightPixels)
    {
        var bytes = File.ReadAllBytes(emfPath);
        if (bytes.Length < 40)
            throw new InvalidDataException("PowerPoint geometry EMF header is truncated.");
        var frameWidth = BitConverter.ToInt32(bytes, 32) - BitConverter.ToInt32(bytes, 24);
        var frameHeight = BitConverter.ToInt32(bytes, 36) - BitConverter.ToInt32(bytes, 28);
        var widthPt = frameWidth * 72f / 2540f;
        var heightPt = frameHeight * 72f / 2540f;
        AssertNear(widthPixels * 0.75f, widthPt, 1.5f,
            "PowerPoint geometry EMF physical width is DPI-dependent.");
        AssertNear(heightPixels * 0.75f, heightPt, 1.5f,
            "PowerPoint geometry EMF physical height is DPI-dependent.");
    }
}
