using System.Diagnostics;
using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNumberedFormulaPerformanceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(tempRoot);
        var svgPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? copyDocument = null;
        var originalNumberFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 240, 72);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 72);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(document, "continuous");
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            var insertTimings = new List<long>();

            for (var index = 1; index <= 6; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var range = application.Selection.Range;
                var formulaId = Guid.NewGuid().ToString("D");
                var session = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(range.Start, range.End),
                    originalMetadata: null,
                    latex: $"x_{{{index}}}={index}");
                Release(range);

                var watch = Stopwatch.StartNew();
                service.InsertOle(session, pngPath, emfPath);
                watch.Stop();
                insertTimings.Add(watch.ElapsedMilliseconds);
                formulaIds.Add(formulaId);
            }

            AssertEqual(6, document.Tables.Count,
                "Numbered OLE performance fixture did not create six equation tables.");
            AssertNumberedFormulaArtifacts(document, formulaIds);
            AssertVisibleEquationNumbers(document, formulaIds, 1);

            Word.InlineShape? editShape = null;
            Word.Range? editRange = null;
            try
            {
                editShape = FindNumberedOleByFormulaId(document, formulaIds[2]);
                var originalMetadata = WordFormulaMetadataReader.TryRead(editShape)
                    ?? throw new InvalidOperationException("Numbered OLE edit metadata is missing.");
                editRange = editShape.Range;
                var editSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaIds[2],
                    document.FullName,
                    WordRangeReference(editRange.Start, editRange.End),
                    originalMetadata,
                    @"x_3=33");
                var watch = Stopwatch.StartNew();
                service.ReplaceOle(editSession, pngPath, emfPath);
                watch.Stop();
                Console.WriteLine($"    [perf] numbered OLE edit: {watch.ElapsedMilliseconds}ms");
                if (watch.ElapsedMilliseconds > 1500)
                    throw new InvalidDataException(
                        $"Numbered OLE edit still took {watch.ElapsedMilliseconds}ms.");
            }
            finally
            {
                Release(editRange);
                Release(editShape);
            }

            Word.InlineShape? ommlSourceShape = null;
            Word.Range? ommlSourceRange = null;
            Word.Bookmark? ommlBookmark = null;
            Word.Range? ommlRange = null;
            try
            {
                var formulaId = formulaIds[3];
                ommlSourceShape = FindNumberedOleByFormulaId(document, formulaId);
                var sourceMetadata = WordFormulaMetadataReader.TryRead(ommlSourceShape)
                    ?? throw new InvalidOperationException("Numbered OLE->OMML source metadata is missing.");
                ommlSourceRange = ommlSourceShape.Range;
                var convertSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(ommlSourceRange.Start, ommlSourceRange.End),
                    sourceMetadata,
                    @"y_4=4");
                convertSession.ObjectMode = FormulaOleContract.WordOmmlMode;
                service.ReplaceOmml(convertSession, PerformanceMathMl(4, 1));

                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException("Converted numbered OMML bookmark is missing.");
                var ommlMetadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark)
                    ?? throw new InvalidOperationException("Converted numbered OMML metadata is missing.");
                ommlRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
                var editSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(ommlRange.Start, ommlRange.End),
                    ommlMetadata,
                    @"y_4=44");
                editSession.ObjectMode = FormulaOleContract.WordOmmlMode;

                var watch = Stopwatch.StartNew();
                service.ReplaceOmml(editSession, PerformanceMathMl(4, 2));
                watch.Stop();
                Console.WriteLine($"    [perf] numbered OMML edit: {watch.ElapsedMilliseconds}ms");
                if (watch.ElapsedMilliseconds > 1500)
                    throw new InvalidDataException(
                        $"Numbered OMML edit still took {watch.ElapsedMilliseconds}ms.");
            }
            finally
            {
                Release(ommlRange);
                Release(ommlBookmark);
                Release(ommlSourceRange);
                Release(ommlSourceShape);
            }

            Word.Table? sourceTable = null;
            Word.Range? sourceTableRange = null;
            Word.Table? copySourceTable = null;
            Word.Range? copySourceRange = null;
            Word.InlineShape? copiedShape = null;
            Word.Range? copiedShapeRange = null;
            var copiedFormulaIds = new List<string>();
            try
            {
                sourceTable = document.Tables[1];
                sourceTableRange = sourceTable.Range.Duplicate;
                sourceTableRange.Copy();

                copyDocument = application.Documents.Add();
                copyDocument.Activate();
                WordEquationNumbering.SetEquationNumberFormatPreference(copyDocument, "continuous");
                application.Selection.Paste();
                AssertEqual(1, copyDocument.Tables.Count,
                    "Numbered OLE copy performance fixture could not seed the copy document.");

                copiedShape = copyDocument.Tables[1].Cell(1, 2).Range.InlineShapes[1];
                copiedShapeRange = copiedShape.Range;
                copiedShapeRange.Select();
                var firstSelection = service.ReadSelection();
                if (string.IsNullOrWhiteSpace(firstSelection.FormulaId)
                    || firstSelection.Metadata is null)
                    throw new InvalidOperationException(
                        "Seeded numbered OLE formula was not recognized in the copy document.");
                // A table copied into a different document carries the visible
                // number cell but not VisualTeX's hidden caption paragraph. Seed
                // the destination document with one complete local numbering
                // scaffold before measuring same-document duplication.
                WordEquationNumbering.ReconcileFormula(
                    copyDocument,
                    copiedShapeRange,
                    copiedShape.Height,
                    firstSelection.Metadata,
                    numberingOrderMayHaveChanged: true);
                copiedFormulaIds.Add(firstSelection.FormulaId!);
                Release(copiedShapeRange);
                copiedShapeRange = null;
                Release(copiedShape);
                copiedShape = null;

                long finalPasteMs = 0;
                long finalRepairMs = 0;
                for (var copyIndex = 2; copyIndex <= 7; copyIndex++)
                {
                    Release(copySourceRange);
                    copySourceRange = null;
                    Release(copySourceTable);
                    copySourceTable = copyDocument.Tables[copyIndex - 1];
                    copySourceRange = copySourceTable.Range.Duplicate;
                    copySourceRange.Copy();
                    application.Selection.EndKey(Word.WdUnits.wdStory);
                    application.Selection.TypeParagraph();

                    var pasteWatch = Stopwatch.StartNew();
                    application.Selection.Paste();
                    pasteWatch.Stop();
                    AssertEqual(copyIndex, copyDocument.Tables.Count,
                        $"Numbered OLE copy performance fixture did not paste table {copyIndex}.");

                    Release(copiedShapeRange);
                    copiedShapeRange = null;
                    Release(copiedShape);
                    copiedShape = copyDocument.Tables[copyIndex].Cell(1, 2).Range.InlineShapes[1];
                    copiedShapeRange = copiedShape.Range;
                    copiedShapeRange.Select();
                    var repairWatch = Stopwatch.StartNew();
                    var copiedSelection = service.ReadSelection();
                    repairWatch.Stop();
                    if (string.IsNullOrWhiteSpace(copiedSelection.FormulaId)
                        || copiedFormulaIds.Any(id => string.Equals(
                            id,
                            copiedSelection.FormulaId,
                            StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException(
                            $"Copied numbered OLE formula {copyIndex} did not receive an independent FormulaId.");
                    copiedFormulaIds.Add(copiedSelection.FormulaId!);
                    if (copyIndex == 7)
                    {
                        finalPasteMs = pasteWatch.ElapsedMilliseconds;
                        finalRepairMs = repairWatch.ElapsedMilliseconds;
                    }
                }

                Console.WriteLine(
                    $"    [perf] seventh numbered OLE table paste: {finalPasteMs}ms; "
                    + $"identity/number repair: {finalRepairMs}ms");
                if (finalRepairMs > 1500)
                    throw new InvalidDataException(
                        $"Numbered OLE copy identity/number repair still took {finalRepairMs}ms at seven formulas.");
                AssertNumberedFormulaArtifacts(copyDocument, copiedFormulaIds);
                AssertVisibleEquationNumbers(copyDocument, copiedFormulaIds, 1);
            }
            finally
            {
                Release(copiedShapeRange);
                Release(copiedShape);
                Release(copySourceRange);
                Release(copySourceTable);
                Release(sourceTableRange);
                Release(sourceTable);
                if (copyDocument is not null)
                {
                    try { copyDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Release(copyDocument);
                    copyDocument = null;
                    document.Activate();
                }
            }

            AssertNumberedFormulaArtifacts(document, formulaIds);
            AssertVisibleEquationNumbers(document, formulaIds, 1);
            Console.WriteLine(
                $"    [perf] sixth numbered OLE insert: {insertTimings[5]}ms "
                + $"(all inserts: {string.Join(",", insertTimings)}ms)");
            if (insertTimings[5] > 1500)
                throw new InvalidDataException(
                    $"The sixth numbered OLE insertion still took {insertTimings[5]}ms.");

            document.SaveAs2(
                Path.Combine(artifactRoot, "word-numbered-formula-performance.docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "Word numbered formula performance acceptance passed: six-formula OLE insert, "
                + "OLE edit, OMML edit and copied-table identity repair stayed on localized numbering paths.");
        }
        finally
        {
            if (copyDocument is not null)
            {
                try { copyDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(copyDocument);
            }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(originalNumberFormat);
            ForceComCleanup();
            foreach (var path in new[] { svgPath, pngPath, emfPath })
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { File.Delete(path!); } catch { }
            }
        }
    }

    private static OfficeSessionDocument CreateNumberedPerformanceSession(
        string mode,
        string formulaId,
        string sourceDocumentId,
        string sourceObjectId,
        FormulaMetadata? originalMetadata,
        string latex)
    {
        var session = CreateOleFontSession(
            "word",
            mode,
            formulaId,
            sourceDocumentId,
            sourceObjectId,
            originalMetadata,
            "times",
            "songti");
        session.Numbered = true;
        session.Title = "Numbered formula performance acceptance";
        session.Lines = new List<FormulaLine>
        {
            new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
        };
        return session;
    }

    private static string PerformanceMathMl(int index, int round) =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + $"<msub><mi>y</mi><mn>{index}</mn></msub><mo>=</mo><mn>{index * 10 + round}</mn>"
        + "</math>";

    private static Word.InlineShape FindNumberedOleByFormulaId(
        Word.Document document,
        string formulaId)
    {
        var shapes = document.InlineShapes;
        try
        {
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata?.Numbered == true
                        && string.Equals(
                            metadata.FormulaId,
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
            throw new InvalidOperationException($"Numbered OLE formula {formulaId} was not found.");
        }
        finally { Release(shapes); }
    }

    private static void AssertNumberedFormulaArtifacts(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        foreach (var formulaId in formulaIds)
        {
            AssertTrue(
                WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(document, formulaId),
                $"Numbered formula {formulaId} lost its numbering artifacts.");
        }
    }

    private static void AssertVisibleEquationNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        int firstNumber)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 0; index < formulaIds.Count; index++)
            {
                var bookmarkName = WordEquationNumbering.EquationBookmarkName(formulaIds[index]);
                if (!bookmarks.Exists(bookmarkName))
                    throw new InvalidOperationException($"Visible number bookmark {bookmarkName} is missing.");
                Release(bookmark);
                bookmark = bookmarks[bookmarkName];
                var text = (bookmark.Range.Text ?? string.Empty).Trim();
                AssertEqual(
                    $"({firstNumber + index})",
                    text,
                    $"Visible numbered formula {index + 1} is stale.");
            }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }
}
