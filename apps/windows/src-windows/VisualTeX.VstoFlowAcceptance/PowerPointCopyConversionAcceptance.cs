using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string PowerPointAcceptanceLatex = @"\frac{a}{b}+\sqrt{x}";
    private const string PowerPointAcceptanceMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
        "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo><msqrt><mi>x</mi></msqrt></math>";

    private static void RunPowerPointCopyConversionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide1 = null;
        PowerPoint.Slide? slide2 = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.Shape? copiedShape = null;
        PowerPoint.ShapeRange? pastedRange = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var officeTempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(officeTempRoot);
        var svgPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        var presentationPath = Path.Combine(artifactRoot, "powerpoint-copy-conversion.pptx");
        var ommlSnapshotPath = Path.Combine(artifactRoot, "powerpoint-native-omml-snapshot.pptx");

        File.WriteAllText(svgPath, CreateSvg(220, 72));
        WriteAcceptancePng(pngPath, PowerPointAcceptanceLatex, 440, 144);

        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide1 = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            slide2 = presentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var service = new PowerPointFormulaService(application);
            emfPath = CreatePowerPointAcceptanceEmf(svgPath, 220, 72);

            var formulaId = Guid.NewGuid().ToString();
            var createPicture = CreatePowerPointAcceptanceSession(
                mode: "create",
                objectMode: "crossPlatformPicture",
                formulaId: formulaId,
                sourceObjectId: null,
                originalMetadata: null);
            var pictureResult = service.Insert(createPicture, svgPath);
            shape = slide1.Shapes[pictureResult.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "Initial PowerPoint VisualTeX formula was not a picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Initial PowerPoint picture FormulaId mismatch.");

            // Cross-slide copy is the hard case because PowerPoint shape names are
            // only unique per slide. The persisted SlideID:Shape.Id owner token must
            // still distinguish the copied picture from its source.
            shape.Copy();
            application.ActiveWindow.View.GotoSlide(2);
            pastedRange = slide2.Shapes.Paste();
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedPictureSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedPictureSelection.FormulaId), "Copied picture was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedPictureSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied picture reused the source FormulaId.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Reading the copied picture changed the source FormulaId.");
            var copiedPictureId = copiedPictureSelection.FormulaId!;
            var sourcePictureWidth = shape.Width;
            var copiedPictureWidth = copiedShape.Width;
            service.SetSelectedFormulaFontSize(30);
            AssertNear(sourcePictureWidth, shape.Width, 0.5f, "Formatting the copied picture resized the source picture.");
            AssertTrue(Math.Abs(copiedShape.Width - copiedPictureWidth) > 0.5f, "Formatting the copied picture did not target the copy.");
            var copiedPictureMetadataAfterFormatting = DecodePowerPointMetadata(copiedShape);
            AssertEqual(copiedPictureId, copiedPictureMetadataAfterFormatting?.FormulaId, "Copied picture lost its independent FormulaId after formatting.");
            AssertNear(30f, (float)(copiedPictureMetadataAfterFormatting?.FontSizePt ?? 0), 0.1f, "Copied picture metadata did not persist the requested font size.");
            Console.WriteLine("PowerPoint picture copy identity passed: cross-slide copy received an independent FormulaId and formatting targeted only the copy.");

            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);

            // Picture -> OLE.
            var sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint picture metadata disappeared before OLE conversion.");
            var pictureToOle = CreatePowerPointAcceptanceSession(
                "edit",
                "nativeOle",
                formulaId,
                shape.Name,
                sourceMetadata);
            var oleResult = service.ReplaceOle(pictureToOle, pngPath, emfPath);
            Release(shape);
            shape = slide1.Shapes[oleResult.ObjectId];
            AssertPowerPointOle(shape, "Picture -> OLE conversion did not create a VisualTeX OLE object.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Picture -> OLE changed FormulaId.");

            // Copy the OLE on the same slide. The copy inherits both embedded
            // metadata and shape tags, but the owner token must force a new id.
            shape.Copy();
            pastedRange = slide1.Shapes.Paste();
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedOleSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedOleSelection.FormulaId), "Copied OLE was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedOleSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied OLE reused the source FormulaId.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Reading the copied OLE changed the source FormulaId.");
            var copiedOleId = copiedOleSelection.FormulaId!;
            var sourceOleWidth = shape.Width;
            var copiedOleWidth = copiedShape.Width;
            service.SetSelectedFormulaFontSize(32);
            AssertNear(sourceOleWidth, shape.Width, 0.5f, "Formatting the copied OLE resized the source OLE.");
            AssertTrue(Math.Abs(copiedShape.Width - copiedOleWidth) > 0.5f, "Formatting the copied OLE did not target the copy.");
            var copiedOleMetadataAfterFormatting = DecodePowerPointMetadata(copiedShape);
            AssertEqual(copiedOleId, copiedOleMetadataAfterFormatting?.FormulaId, "Copied OLE lost its independent FormulaId after formatting.");
            AssertNear(32f, (float)(copiedOleMetadataAfterFormatting?.FontSizePt ?? 0), 0.1f, "Copied OLE metadata did not persist the requested font size.");
            Console.WriteLine("PowerPoint OLE copy identity passed: copied OLE received an independent FormulaId and editing geometry stayed isolated.");

            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            shape.Select(MsoTriState.msoTrue);

            // OLE -> picture.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint OLE metadata disappeared before picture conversion.");
            var oleToPicture = CreatePowerPointAcceptanceSession(
                "edit",
                "crossPlatformPicture",
                formulaId,
                shape.Name,
                sourceMetadata);
            var backToPicture = service.Replace(oleToPicture, svgPath);
            Release(shape);
            shape = slide1.Shapes[backToPicture.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "OLE -> picture conversion did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "OLE -> picture changed FormulaId.");

            // Picture -> native Office Math / OMML.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint picture metadata disappeared before OMML conversion.");
            var pictureToOmml = CreatePowerPointAcceptanceSession(
                "edit",
                "wordOmml",
                formulaId,
                shape.Name,
                sourceMetadata);
            var ommlResult = service.ReplaceOmml(pictureToOmml);
            Release(shape);
            shape = slide1.Shapes[ommlResult.ObjectId];
            AssertTrue(IsPowerPointNativeEquation(shape), "Picture -> OMML did not create native PowerPoint Office Math.");
            shape.Select(MsoTriState.msoTrue);
            var nativeSelection = service.ReadSelection();
            AssertEqual("wordOmml", nativeSelection.ObjectMode, "Native PowerPoint equation was not recognized as OMML mode.");
            AssertEqual(formulaId, nativeSelection.FormulaId, "Picture -> OMML changed FormulaId.");
            AssertContains(nativeSelection.Metadata?.Latex, @"\frac", "Native OMML did not round-trip the fraction back to LaTeX.");
            AssertContains(nativeSelection.Metadata?.Latex, @"\sqrt", "Native OMML did not round-trip the radical back to LaTeX.");

            presentation.SaveCopyAs(
                ommlSnapshotPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);
            AssertPowerPointPptxContainsNativeOmml(ommlSnapshotPath);

            // OMML copy gets an independent identity as well.
            shape.Copy();
            pastedRange = slide1.Shapes.Paste();
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedOmmlSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedOmmlSelection.FormulaId), "Copied OMML was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedOmmlSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied OMML reused the source FormulaId.");
            AssertTrue(IsPowerPointNativeEquation(copiedShape), "Copied OMML ceased to be native Office Math.");
            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            shape.Select(MsoTriState.msoTrue);

            // OMML -> OLE.
            sourceMetadata = service.ReadSelection().Metadata
                ?? throw new InvalidDataException("PowerPoint OMML metadata disappeared before OLE conversion.");
            var ommlToOle = CreatePowerPointAcceptanceSession(
                "edit",
                "nativeOle",
                formulaId,
                shape.Name,
                sourceMetadata);
            var ommlOleResult = service.ReplaceOle(ommlToOle, pngPath, emfPath);
            Release(shape);
            shape = slide1.Shapes[ommlOleResult.ObjectId];
            AssertPowerPointOle(shape, "OMML -> OLE conversion did not create a VisualTeX OLE object.");

            // OLE -> OMML.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint OLE metadata disappeared before OMML conversion.");
            var oleToOmml = CreatePowerPointAcceptanceSession(
                "edit",
                "wordOmml",
                formulaId,
                shape.Name,
                sourceMetadata);
            var secondOmml = service.ReplaceOmml(oleToOmml);
            Release(shape);
            shape = slide1.Shapes[secondOmml.ObjectId];
            AssertTrue(IsPowerPointNativeEquation(shape), "OLE -> OMML conversion did not create native Office Math.");

            // OMML -> picture.
            shape.Select(MsoTriState.msoTrue);
            sourceMetadata = service.ReadSelection().Metadata
                ?? throw new InvalidDataException("PowerPoint OMML metadata disappeared before picture conversion.");
            var ommlToPicture = CreatePowerPointAcceptanceSession(
                "edit",
                "crossPlatformPicture",
                formulaId,
                shape.Name,
                sourceMetadata);
            var finalPicture = service.Replace(ommlToPicture, svgPath);
            Release(shape);
            shape = slide1.Shapes[finalPicture.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "OMML -> picture conversion did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Three-way conversion changed the original FormulaId.");

            // Stress the direct PowerPoint MathML importer. This specifically
            // guards against the old transient Word automation path, which could
            // fail intermittently even when a single conversion succeeded.
            var ommlWriteMilliseconds = new List<double>();
            for (var iteration = 1; iteration <= 25; iteration++)
            {
                sourceMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"PowerPoint stress picture metadata disappeared at iteration {iteration}.");
                var stressToOmml = CreatePowerPointAcceptanceSession(
                    "edit",
                    "wordOmml",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var stressOmmlResult = service.ReplaceOmml(stressToOmml);
                stopwatch.Stop();
                ommlWriteMilliseconds.Add(stopwatch.Elapsed.TotalMilliseconds);
                Release(shape);
                shape = slide1.Shapes[stressOmmlResult.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape), $"Stress picture -> OMML failed at iteration {iteration}.");
                shape.Select(MsoTriState.msoTrue);
                var stressRead = service.ReadSelection();
                AssertContains(stressRead.Metadata?.Latex, @"\frac", $"Stress OMML fraction readback failed at iteration {iteration}.");
                AssertContains(stressRead.Metadata?.Latex, @"\sqrt", $"Stress OMML radical readback failed at iteration {iteration}.");

                sourceMetadata = stressRead.Metadata
                    ?? throw new InvalidDataException($"PowerPoint stress OMML metadata disappeared at iteration {iteration}.");
                var stressToPicture = CreatePowerPointAcceptanceSession(
                    "edit",
                    "crossPlatformPicture",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var stressPictureResult = service.Replace(stressToPicture, svgPath);
                Release(shape);
                shape = slide1.Shapes[stressPictureResult.ObjectId];
                AssertTrue(IsPowerPointEditablePictureShape(shape), $"Stress OMML -> picture failed at iteration {iteration}.");
            }
            var orderedOmmlWrites = ommlWriteMilliseconds.OrderBy(value => value).ToArray();
            var ommlP50 = orderedOmmlWrites[orderedOmmlWrites.Length / 2];
            var ommlMax = orderedOmmlWrites[orderedOmmlWrites.Length - 1];
            Console.WriteLine(
                $"PowerPoint direct MathML stress passed: 25/25 picture -> OMML -> readback -> picture cycles; " +
                $"OMML write p50={ommlP50:F1} ms, max={ommlMax:F1} ms.");

            // Exercise the actual Ribbon -> Session -> converter -> PowerPoint path
            // for the new OMML mode and both existing conversion callbacks.
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            application.ActiveWindow.Activate();
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);
            var existing = SnapshotSessionIds();
            var converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "wordOmml",
                () => addIn.OnConvertSelectedOmml(new object()),
                TimeSpan.FromSeconds(45),
                out var ribbonPictureToOmmlElapsed,
                () => addIn.DiagnosticLastError);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon picture-to-OMML conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointNativeEquation(shape), "Ribbon picture -> OMML did not create native Office Math.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "nativeOle",
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OMML-to-OLE conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertPowerPointOle(shape, "Ribbon OMML -> OLE did not create a VisualTeX OLE object.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "wordOmml",
                () => addIn.OnConvertSelectedOmml(new object()),
                TimeSpan.FromSeconds(45),
                out var ribbonOleToOmmlElapsed);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OLE-to-OMML conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointNativeEquation(shape), "Ribbon OLE -> OMML did not create native Office Math.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "crossPlatformPicture",
                () => addIn.OnExportSelectedAsPicture(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OMML-to-picture conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "Ribbon OMML -> picture did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Ribbon conversion chain changed the original FormulaId.");
            Console.WriteLine(
                "PowerPoint Ribbon conversion path passed: SVG -> OMML -> OLE -> OMML -> SVG completed through VisualTeX Sessions. " +
                $"SVG->OMML={ribbonPictureToOmmlElapsed.TotalMilliseconds:F0} ms, " +
                $"OLE->OMML={ribbonOleToOmmlElapsed.TotalMilliseconds:F0} ms.");

            presentation.SaveAs(
                presentationPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);
            Console.WriteLine("PowerPoint three-way conversion passed: picture <-> OLE, picture <-> OMML, and OLE <-> OMML all completed with native Office Math round-trip verification.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(copiedShape);
            Release(pastedRange);
            Release(shape);
            Release(slide2);
            Release(slide1);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            if (!string.IsNullOrWhiteSpace(emfPath))
            {
                try { File.Delete(emfPath); } catch { }
            }
            try { File.Delete(svgPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreatePowerPointAcceptanceSession(
        string mode,
        string objectMode,
        string formulaId,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString(),
            Mode = mode,
            Host = "powerpoint",
            FormulaId = formulaId,
            SourceDocumentId = null,
            SourceObjectId = sourceObjectId,
            Title = "PowerPoint copy/conversion acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = PowerPointAcceptanceLatex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = objectMode,
            Numbered = false,
            FontSizePt = 20,
            OriginalMetadata = originalMetadata,
            ExportResult = new OfficeExportDocument
            {
                Svg = CreateSvg(220, 72),
                SvgBase64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(CreateSvg(220, 72))),
                MathMl = PowerPointAcceptanceMathMl,
                PngBase64 = null,
                Width = 220,
                Height = 72,
                Baseline = 54,
            },
        };
    }

    private static void WriteAcceptancePng(
        string path,
        string text,
        int width,
        int height)
    {
        using var bitmap = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(System.Drawing.Color.Transparent);
        using var font = new System.Drawing.Font("Cambria Math", 30f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
        graphics.DrawString(text, font, brush, 6, 18);
        bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }

    private static string CreatePowerPointAcceptanceEmf(
        string svgPath,
        float width,
        float height)
    {
        var type = typeof(PowerPointFormulaService).Assembly.GetType(
            "VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview",
            throwOnError: true)
            ?? throw new InvalidOperationException("PowerPoint OfficeOlePreview type is unavailable.");
        var method = type.GetMethod(
            "CreateVectorEmfFromSvg",
            BindingFlags.Public | BindingFlags.Static)
            ?? throw new MissingMethodException(type.FullName, "CreateVectorEmfFromSvg");
        return (string)(method.Invoke(null, new object[] { svgPath, width, height })
            ?? throw new InvalidOperationException("PowerPoint EMF preview generation returned null."));
    }

    private static void AssertPowerPointOle(PowerPoint.Shape shape, string message)
    {
        PowerPoint.OLEFormat? format = null;
        try
        {
            if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
                and not MsoShapeType.msoLinkedOLEObject)
                throw new InvalidDataException(message);
            format = shape.OLEFormat;
            AssertEqual(
                FormulaOleContract.ProgId,
                format.ProgID,
                message);
        }
        finally { Release(format); }
    }

    private static void AssertPowerPointPptxContainsNativeOmml(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var xml = string.Join(
            "\n",
            archive.Entries
                .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                    && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .Select(entry =>
                {
                    using var stream = entry.Open();
                    using var reader = new StreamReader(stream);
                    return reader.ReadToEnd();
                }));
        AssertTrue(xml.IndexOf("<a14:m", StringComparison.Ordinal) >= 0, "Saved PowerPoint did not contain an a14:m native math wrapper.");
        AssertTrue(xml.IndexOf("<m:oMath", StringComparison.Ordinal) >= 0, "Saved PowerPoint did not contain OMML oMath markup.");
        AssertTrue(xml.IndexOf("<m:f>", StringComparison.Ordinal) >= 0, "Saved PowerPoint OMML did not contain the expected fraction structure.");
        AssertTrue(xml.IndexOf("<m:rad>", StringComparison.Ordinal) >= 0, "Saved PowerPoint OMML did not contain the expected radical structure.");
    }

    private static bool IsPowerPointNativeEquation(PowerPoint.Shape shape)
    {
        object? textFrame = null;
        object? range = null;
        object? mathZones = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            try { mathZones = ((dynamic)range).MathZones(); }
            catch { mathZones = ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
            return Convert.ToInt32(((dynamic)mathZones).Length) > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    private static void AssertContains(string? value, string expected, string message)
    {
        if (value is null || value.IndexOf(expected, StringComparison.Ordinal) < 0)
            throw new InvalidDataException($"{message} Actual: {value ?? "<null>"}");
    }
}
