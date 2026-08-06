using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBatchEquationNumbering(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(
            artifactRoot,
            "word-batch-equation-numbering.docx");
        DeleteBulkPerformanceArtifact(outputPath);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var previousAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE");
        var previousRedraw = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_BATCH_NUMBER_REDRAW");
        var previousFormat = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_BATCH_NUMBER_FORMAT");
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            InsertNumberingHeading(application, document, level: 1, "Batch Chapter");
            var formulas = new List<NumberedFormulaCase>();
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "u_1=1",
                numbered: true));
            AppendBatchAcceptanceParagraph(application, document);
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "u_2=2",
                numbered: false));
            AppendBatchAcceptanceParagraph(application, document);
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "u_3=3",
                numbered: false));
            AppendBatchAcceptanceParagraph(application, document);
            var migratedOle = InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.NativeOleMode,
                "u_4=4",
                numbered: true);
            FlattenNumberedOleForBatchAcceptance(document, migratedOle.FormulaId);
            formulas.Add(migratedOle);
            AppendBatchAcceptanceParagraph(application, document);

            var service = new WordFormulaService(application);
            var before = service.GetBatchEquationNumberingStats();
            AssertEqual(4, before.TotalCount,
                "Batch numbering did not discover all OLE/OMML display formulas.");
            AssertEqual(1, before.NumberedCount,
                "Batch numbering did not distinguish the one complete number from the flattened legacy formula that needs repair.");

            var referenceBookmark = InsertBatchNumberingReference(
                application,
                document,
                formulas[0].FormulaId);
            AssertReferenceText(document, referenceBookmark, "(1)");

            var redrawResult = service.ApplyBatchEquationNumbering(
                EquationNumberFormat.Heading1DotId,
                redrawExisting: true);
            AssertEqual(4, redrawResult.NumberedCount,
                "Batch redraw did not retain all four numbered formulas.");
            AssertEqual(2, redrawResult.AddedCount,
                "Batch redraw did not add numbers to both still-unnumbered formulas.");
            AssertEqual(2, redrawResult.RedrawnCount,
                "Batch redraw did not rebuild the original number and the legacy OLE number repaired while resolving references.");
            AssertEquationNumberFormat(
                document,
                addIn,
                EquationNumberFormat.Heading1DotId,
                formulas,
                new[] { "1.1", "1.2", "1.3", "1.4" });
            AssertReferenceText(document, referenceBookmark, "(1.1)");
            AssertBatchFormulaMetadata(document, formulas, expectedNumbered: true);

            var afterRedraw = service.GetBatchEquationNumberingStats();
            AssertEqual(4, afterRedraw.TotalCount,
                "Batch numbering changed the logical formula count.");
            AssertEqual(4, afterRedraw.NumberedCount,
                "Batch numbering did not persist numbered metadata for every formula.");

            // Exercise the actual Ribbon callback for the non-redraw path. In
            // acceptance mode the environment variables stand in for the user
            // choosing “No, only fill missing numbers” and then selecting 1-1.
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_BATCH_NUMBER_REDRAW",
                "0");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_BATCH_NUMBER_FORMAT",
                EquationNumberFormat.Heading1DashId);
            addIn.OnBatchEquationNumbering(new object());
            WaitForBatchEquationNumbers(
                document,
                addIn,
                formulas,
                EquationNumberFormat.Heading1DashId,
                new[] { "1-1", "1-2", "1-3", "1-4" });
            AssertReferenceText(document, referenceBookmark, "(1-1)");
            AssertBatchFormulaMetadata(document, formulas, expectedNumbered: true);

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEquationNumberFormat(
                reopened,
                addIn,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "1-4" });
            AssertReferenceText(reopened, referenceBookmark, "(1-1)");
            AssertBatchFormulaMetadata(reopened, formulas, expectedNumbered: true);

            Console.WriteLine(
                "Word batch equation-numbering acceptance passed: mixed existing/missing OLE/OMML numbers were detected, a legacy standalone OLE display formula was migrated to the standard three-column numbered layout, redraw and fill-only paths kept formula ids and REF targets stable, the selected chapter format updated references, and numbered metadata survived save/reopen.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_BATCH_NUMBER_FORMAT",
                previousFormat);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_BATCH_NUMBER_REDRAW",
                previousRedraw);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AppendBatchAcceptanceParagraph(
        Word.Application application,
        Word.Document document)
    {
        Word.Selection? selection = null;
        Word.Range? content = null;
        try
        {
            selection = application.Selection;
            content = document.Content;
            var end = Math.Max(content.Start, content.End - 1);
            selection.SetRange(end, end);
            selection.TypeParagraph();
        }
        finally
        {
            Release(content);
            Release(selection);
        }
    }

    private static string InsertBatchNumberingReference(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Selection? selection = null;
        Word.Range? content = null;
        Word.Range? referenceRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            selection = application.Selection;
            var target = WordEquationNumbering.GetEquationReferenceTargets(document)
                .Single(item => string.Equals(
                    item.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase));
            EnsureBatchBodyParagraphAtDocumentEnd(document);
            content = document.Content;
            var insertionPosition = Math.Max(content.Start, content.End - 1);
            selection.SetRange(insertionPosition, insertionPosition);
            selection.TypeText("Reference: ");
            var start = selection.Start;
            WordEquationNumbering.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized);
            var end = selection.Start;
            selection.TypeParagraph();
            referenceRange = document.Range(start, end);
            bookmarks = document.Bookmarks;
            const string bookmarkName = "VTTestBatchEquationNumberReference";
            bookmark = bookmarks.Add(bookmarkName, referenceRange);
            return bookmarkName;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(referenceRange);
            Release(content);
            Release(selection);
        }
    }

    private static void EnsureBatchBodyParagraphAtDocumentEnd(
        Word.Document document)
    {
        Word.Range? content = null;
        Word.Range? probe = null;
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Range? tableRange = null;
        try
        {
            content = document.Content;
            var end = Math.Max(content.Start, content.End - 1);
            probe = document.Range(end, end);
            if (!(bool)probe.get_Information(Word.WdInformation.wdWithInTable))
                return;
            tables = document.Tables;
            if (tables.Count == 0)
                throw new InvalidDataException(
                    "Word reported a table insertion point without a document table.");
            table = tables[tables.Count];
            tableRange = table.Range;
            tableRange.InsertParagraphAfter();
        }
        finally
        {
            Release(tableRange);
            Release(table);
            Release(tables);
            Release(probe);
            Release(content);
        }
    }

    private static void FlattenNumberedOleForBatchAcceptance(
        Word.Document document,
        string formulaId)
    {
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Table? table = null;
        Word.Range? converted = null;
        try
        {
            shape = FindBatchOleShape(document, formulaId)
                ?? throw new InvalidDataException(
                    $"OLE formula {formulaId} is missing before batch-numbering acceptance.");
            shapeRange = shape.Range;
            if (!(bool)shapeRange.get_Information(Word.WdInformation.wdWithInTable)
                || shapeRange.Tables.Count == 0)
                throw new InvalidDataException(
                    "The numbered OLE acceptance formula was not inserted in the standard table layout.");

            WordEquationNumbering.RemoveFormulaNumberingArtifacts(document, formulaId);
            Release(table);
            table = shapeRange.Tables[1];
            object separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
            object nestedTables = false;
            converted = table.ConvertToText(ref separator, ref nestedTables);

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = FindBatchOleShape(document, formulaId)
                ?? throw new InvalidDataException(
                    "The OLE acceptance formula disappeared while flattening its table.");
            shapeRange = shape.Range;
            if ((bool)shapeRange.get_Information(Word.WdInformation.wdWithInTable))
                throw new InvalidDataException(
                    "The OLE acceptance formula remained in a table after flattening.");
        }
        finally
        {
            Release(converted);
            Release(table);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static Word.InlineShape? FindBatchOleShape(
        Word.Document document,
        string formulaId)
    {
        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is null
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var result = shape;
                    shape = null;
                    return result;
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static void WaitForBatchEquationNumbers(
        Word.Document document,
        VisualTeX.WordVsto.ThisAddIn addIn,
        IReadOnlyList<NumberedFormulaCase> formulas,
        string formatId,
        IReadOnlyList<string> expectedNumbers)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(30);
        Exception? lastError = null;
        var dumped = false;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(3));
                AssertEquationNumberFormat(
                    document,
                    addIn,
                    formatId,
                    formulas,
                    expectedNumbers);
                return;
            }
            catch (Exception error)
            {
                lastError = error;
                if (!dumped)
                {
                    dumped = true;
                    DumpBatchNumberingState(document, formulas);
                }
                Thread.Sleep(100);
            }
        }
        throw new InvalidOperationException(
            "The batch-equation-numbering Ribbon command did not complete in time.",
            lastError);
    }

    private static void DumpBatchNumberingState(
        Word.Document document,
        IReadOnlyList<NumberedFormulaCase> formulas)
    {
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formula in formulas)
            {
                var visibleName = WordEquationNumbering.EquationBookmarkName(formula.FormulaId);
                var numberName = WordEquationNumbering.NativeNumberBookmarkName(formula.FormulaId);
                var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formula.FormulaId);
                Console.WriteLine(
                    $"[batch-number diagnostic] {formula.FormulaId}: "
                    + $"visible={bookmarks.Exists(visibleName)}, "
                    + $"nativeNumber={bookmarks.Exists(numberName)}, "
                    + $"nativeCaption={bookmarks.Exists(captionName)}");
                Word.Bookmark? formulaBookmark = null;
                Word.Range? formulaRange = null;
                Word.Paragraphs? paragraphs = null;
                Word.Paragraph? paragraph = null;
                Word.Range? paragraphRange = null;
                try
                {
                    formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        formula.FormulaId);
                    if (formulaBookmark is not null)
                    {
                        formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
                        paragraphs = formulaRange.Paragraphs;
                        paragraph = paragraphs[1];
                        paragraphRange = paragraph.Range;
                        Console.WriteLine(
                            $"[batch-number diagnostic] formula={formulaRange.Start}-{formulaRange.End}, "
                            + $"paragraph={paragraphRange.Start}-{paragraphRange.End}");
                    }
                }
                finally
                {
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                    Release(formulaRange);
                    Release(formulaBookmark);
                }
                if (!bookmarks.Exists(visibleName)) continue;
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                Word.Fields? fields = null;
                try
                {
                    bookmark = bookmarks[visibleName];
                    range = bookmark.Range;
                    fields = range.Fields;
                    Console.WriteLine(
                        $"[batch-number diagnostic] visible={range.Start}-{range.End}, "
                        + $"visibleText={range.Text}, fields={fields.Count}");
                    for (var index = 1; index <= fields.Count; index++)
                    {
                        Word.Field? field = null;
                        Word.Range? code = null;
                        try
                        {
                            field = fields[index];
                            code = field.Code;
                            Console.WriteLine(
                                $"[batch-number diagnostic] fieldCode={code.Text}");
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
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static void AssertBatchFormulaMetadata(
        Word.Document document,
        IReadOnlyList<NumberedFormulaCase> formulas,
        bool expectedNumbered)
    {
        foreach (var formula in formulas)
        {
            FormulaMetadata metadata;
            Word.Range? formulaRange = null;
            Word.Bookmark? bookmark = null;
            Word.InlineShape? shape = null;
            try
            {
                if (string.Equals(
                        formula.ObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    metadata = WordOmmlFormulaStore.TryRead(document, formula.FormulaId)
                        ?? throw new InvalidDataException(
                            $"OMML metadata disappeared for formula {formula.FormulaId}.");
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        formula.FormulaId)
                        ?? throw new InvalidDataException(
                            $"OMML bookmark disappeared for formula {formula.FormulaId}.");
                    formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                }
                else
                {
                    shape = FindBatchOleShape(document, formula.FormulaId)
                        ?? throw new InvalidDataException(
                            $"OLE metadata disappeared for formula {formula.FormulaId}.");
                    metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException(
                            $"OLE metadata is unreadable for formula {formula.FormulaId}.");
                    formulaRange = shape.Range;
                }

                AssertEqual(formula.FormulaId, metadata.FormulaId,
                    "Batch numbering changed a formula id.");
                AssertEqual(expectedNumbered, metadata.Numbered,
                    "Batch numbering did not persist the numbered flag.");
                AssertEqual("block", metadata.DisplayMode,
                    "Batch numbering changed a display formula into inline mode.");
                AssertTrue(
                    (bool)formulaRange.get_Information(Word.WdInformation.wdWithInTable)
                    && formulaRange.Tables.Count > 0
                    && formulaRange.Tables[1].Columns.Count >= 3,
                    "Batch numbering did not preserve the standard three-column numbered layout.");
            }
            finally
            {
                Release(formulaRange);
                Release(shape);
                Release(bookmark);
            }
        }
    }
}
