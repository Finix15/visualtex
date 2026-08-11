using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string WordRangeReferencePrefix = "visualtex-word-vsto-range:";

    private static void RunWordOffice2019SequentialNumberedInsertion(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunSequentialNumberedInsertionScenario(
                application,
                artifactRoot,
                "drifted-live-selection",
                captureInsideNumberedTable: false);
            RunSequentialNumberedInsertionScenario(
                application,
                artifactRoot,
                "captured-inside-numbered-table",
                captureInsideNumberedTable: true);
            RunMiddleInsertionFromEarlierNumberedFormulaScenario(
                application,
                artifactRoot);
            Console.WriteLine(
                "Office 2019 sequential numbered insertion acceptance passed: "
                + "a captured create anchor survived live-selection drift, a caret "
                + "captured inside an existing numbered table was redirected after "
                + "that formula's native SEQ paragraph, and a middle insertion no "
                + "longer jumps to the document tail.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunSequentialNumberedInsertionScenario(
        Word.Application application,
        string artifactRoot,
        string scenarioName,
        bool captureInsideNumberedTable)
    {
        Word.Document? document = null;
        Word.Table? firstTable = null;
        Word.Range? firstTableRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            var service = new WordFormulaService(application);
            var firstSession = CreateNumberedOmmlSession(
                document,
                WordRangeReference(0, 0),
                "x=1");
            service.InsertOmml(
                firstSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>=</mo><mn>1</mn></math>");

            firstTable = document.Tables[1];
            firstTableRange = firstTable.Range;
            string capturedAnchor;
            if (captureInsideNumberedTable)
            {
                capturedAnchor = WordRangeReference(
                    firstTableRange.Start,
                    firstTableRange.Start);
            }
            else
            {
                capturedAnchor = service.ReadSelection().ObjectId
                    ?? throw new InvalidOperationException(
                        "Word did not expose the post-formula create anchor.");
            }

            // Simulate Office 2019 moving the live Selection back into the old
            // numbered table while the external VisualTeX editor owns focus.
            application.Selection.SetRange(
                firstTableRange.Start,
                firstTableRange.Start);

            var secondSession = CreateNumberedOmmlSession(
                document,
                capturedAnchor,
                "y=2");
            service.InsertOmml(
                secondSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>y</mi><mo>=</mo><mn>2</mn></math>");

            AssertSequentialNumberedInsertion(document, scenarioName);
            AssertNumberedSpacingCleanup(document, scenarioName);
            if (!captureInsideNumberedTable)
            {
                AssertCompactTailExpandsForTyping(
                    application,
                    service,
                    document,
                    secondSession.FormulaId,
                    scenarioName);
            }
            else
            {
                Word.Row? extraRow = null;
                var legacyRowPath = Path.Combine(
                    artifactRoot,
                    $"word-office2019-{scenarioName}-legacy-row.docx");
                try
                {
                    extraRow = firstTable.Rows.Add();
                    AssertEqual(
                        2,
                        firstTable.Rows.Count,
                        $"{scenarioName}: acceptance could not create the legacy empty numbered-table row.");

                    // A freshly added empty Word row can exist only in the COM model
                    // and be omitted from document.Content.WordOpenXML. The real bug
                    // reported by users is a persisted/reopened legacy document, where
                    // the extra row is serialized. Save and reopen so this acceptance
                    // exercises the same repair path without forcing every healthy
                    // 100-formula update to scan Rows.Count through COM.
                    document.SaveAs2(
                        legacyRowPath,
                        Word.WdSaveFormat.wdFormatXMLDocument);
                    Release(extraRow);
                    extraRow = null;
                    Release(firstTableRange);
                    firstTableRange = null;
                    Release(firstTable);
                    firstTable = null;
                    document.Close(Word.WdSaveOptions.wdSaveChanges);
                    Release(document);
                    document = application.Documents.Open(
                        legacyRowPath,
                        ReadOnly: false,
                        AddToRecentFiles: false,
                        Visible: false);
                    document.Activate();
                    firstTable = document.Tables[1];
                    firstTableRange = firstTable.Range;
                    AssertEqual(
                        2,
                        firstTable.Rows.Count,
                        $"{scenarioName}: persisted legacy row did not survive reopen.");

                    WordEquationNumbering.UpdateEquationNumbers(document);
                    AssertEqual(
                        1,
                        firstTable.Rows.Count,
                        $"{scenarioName}: update numbering did not remove the persisted legacy empty numbered-table row.");
                    AssertNumberedSpacingCleanup(document, scenarioName + "-repair");
                }
                finally { Release(extraRow); }
            }
            var artifactPath = Path.Combine(
                artifactRoot,
                $"word-office2019-{scenarioName}.docx");
            document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
        }
        finally
        {
            Release(firstTableRange);
            Release(firstTable);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void RunMiddleInsertionFromEarlierNumberedFormulaScenario(
        Word.Application application,
        string artifactRoot)
    {
        Word.Document? document = null;
        Word.Table? firstTable = null;
        Word.Range? firstTableRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? firstBookmark = null;
        Word.Bookmark? middleBookmark = null;
        Word.Bookmark? secondBookmark = null;
        Word.Range? firstBookmarkRange = null;
        Word.Range? middleBookmarkRange = null;
        Word.Range? secondBookmarkRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            var service = new WordFormulaService(application);

            var firstSession = CreateNumberedOmmlSession(
                document,
                WordRangeReference(0, 0),
                "x=1");
            service.InsertOmml(
                firstSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>=</mo><mn>1</mn></math>");

            var afterFirstAnchor = service.ReadSelection().ObjectId
                ?? throw new InvalidOperationException(
                    "Middle-insert acceptance could not capture the anchor after the first formula.");
            var secondSession = CreateNumberedOmmlSession(
                document,
                afterFirstAnchor,
                "y=2");
            service.InsertOmml(
                secondSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>y</mi><mo>=</mo><mn>2</mn></math>");

            firstTable = document.Tables[1];
            firstTableRange = firstTable.Range.Duplicate;
            var earlierFormulaAnchor = WordRangeReference(
                firstTableRange.Start,
                firstTableRange.Start);

            // Reproduce the problematic state seen after creating an empty line
            // between two numbered formulas: Word may report the create anchor
            // from the earlier formula's table while the live Selection later
            // drifts elsewhere as the external editor takes focus. The insertion
            // must stay between formula #1 and #2, never at the document tail.
            application.Selection.EndKey(Word.WdUnits.wdStory);
            var middleSession = CreateNumberedOmmlSession(
                document,
                earlierFormulaAnchor,
                "z=3");
            service.InsertOmml(
                middleSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>z</mi><mo>=</mo><mn>3</mn></math>");

            var artifactPath = Path.Combine(
                artifactRoot,
                "word-office2019-middle-insertion.docx");
            document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"    [diag] middle insertion: tables={document.Tables.Count}; "
                + $"omaths={document.OMaths.Count}; paragraphs={document.Paragraphs.Count}");
            for (var tableIndex = 1; tableIndex <= document.Tables.Count; tableIndex++)
            {
                Word.Table? diagnosticTable = null;
                Word.Range? diagnosticRange = null;
                try
                {
                    diagnosticTable = document.Tables[tableIndex];
                    diagnosticRange = diagnosticTable.Range;
                    Console.WriteLine(
                        $"    [diag] table#{tableIndex}: "
                        + $"rows={diagnosticTable.Rows.Count}; cols={diagnosticTable.Columns.Count}; "
                        + $"range={diagnosticRange.Start}:{diagnosticRange.End}");
                }
                finally
                {
                    Release(diagnosticRange);
                    Release(diagnosticTable);
                }
            }

            AssertEqual(
                3,
                document.Tables.Count,
                "Middle-insert acceptance did not create three independent numbered tables.");
            AssertEqual(
                3,
                document.OMaths.Count,
                "Middle-insert acceptance lost a native OMML formula.");

            bookmarks = document.Bookmarks;
            firstBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                firstSession.FormulaId)];
            middleBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                middleSession.FormulaId)];
            secondBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                secondSession.FormulaId)];
            firstBookmarkRange = firstBookmark.Range;
            middleBookmarkRange = middleBookmark.Range;
            secondBookmarkRange = secondBookmark.Range;
            AssertTrue(
                firstBookmarkRange.Start < middleBookmarkRange.Start,
                "Middle-insert acceptance placed the new formula before the first formula.");
            AssertTrue(
                middleBookmarkRange.Start < secondBookmarkRange.Start,
                "Middle-insert acceptance still moved the new formula to the document tail.");
            AssertNumberedSpacingCleanup(document, "middle-insertion");

        }
        finally
        {
            Release(secondBookmarkRange);
            Release(middleBookmarkRange);
            Release(firstBookmarkRange);
            Release(secondBookmark);
            Release(middleBookmark);
            Release(firstBookmark);
            Release(bookmarks);
            Release(firstTableRange);
            Release(firstTable);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateNumberedOmmlSession(
        Word.Document document,
        string sourceObjectId,
        string latex)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            SourceDocumentId = document.FullName,
            SourceObjectId = sourceObjectId,
            Title = "Office 2019 sequential numbered insertion acceptance",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = true,
            FontSizePt = 11,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
        };
    }

    private static string WordRangeReference(int start, int end) =>
        $"{WordRangeReferencePrefix}{start}:{end}";

    private static void AssertNumberedSpacingCleanup(
        Word.Document document,
        string scenarioName)
    {
        for (var index = 1; index <= document.Tables.Count; index++)
        {
            Word.Table? table = null;
            try
            {
                table = document.Tables[index];
                AssertEqual(
                    1,
                    table.Rows.Count,
                    $"{scenarioName}: numbered table {index} retained an empty structural row.");
            }
            finally { Release(table); }
        }

        var ordinaryParagraphCount = 0;
        Word.Range? lastOrdinaryParagraph = null;
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Frames? frames = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                if ((bool)range.get_Information(Word.WdInformation.wdWithInTable))
                    continue;
                frames = range.Frames;
                if (frames.Count > 0) continue;
                ordinaryParagraphCount++;
                Release(lastOrdinaryParagraph);
                lastOrdinaryParagraph = range.Duplicate;
            }
            finally
            {
                Release(frames);
                Release(range);
                Release(paragraph);
            }
        }
        try
        {
            AssertEqual(
                1,
                ordinaryParagraphCount,
                $"{scenarioName}: an ordinary blank paragraph remained between numbered formulas.");
            AssertTrue(
                lastOrdinaryParagraph is not null,
                $"{scenarioName}: Word lost its required terminal paragraph.");
            AssertEqual(
                document.Content.End,
                lastOrdinaryParagraph!.End,
                $"{scenarioName}: the only ordinary paragraph is not the terminal paragraph.");
            AssertTrue(
                lastOrdinaryParagraph.Font.Size <= 1.1f,
                $"{scenarioName}: the terminal structural paragraph still occupies normal line height.");
            AssertEqual(
                Word.WdLineSpacing.wdLineSpaceExactly,
                lastOrdinaryParagraph.ParagraphFormat.LineSpacingRule,
                $"{scenarioName}: the terminal structural paragraph is not compacted.");
            AssertTrue(
                lastOrdinaryParagraph.ParagraphFormat.LineSpacing <= 1.1f,
                $"{scenarioName}: the terminal structural paragraph line spacing is too large.");
        }
        finally { Release(lastOrdinaryParagraph); }
    }

    private static void AssertCompactTailExpandsForTyping(
        Word.Application application,
        WordFormulaService service,
        Word.Document document,
        string tailFormulaId,
        string scenarioName)
    {
        Word.Paragraph? lastParagraph = null;
        Word.Range? lastRange = null;
        Word.Selection? selection = null;
        try
        {
            lastParagraph = document.Paragraphs[document.Paragraphs.Count];
            lastRange = lastParagraph.Range;
            selection = application.Selection;
            selection.SetRange(lastRange.Start, lastRange.Start);
            service.NormalizeTypingCaretAfterInlineFormula(selection);
            Release(lastRange);
            lastRange = lastParagraph.Range;
            AssertTrue(
                lastRange.Font.Size > 1.1f,
                $"{scenarioName}: selecting the compact terminal paragraph did not restore normal typing size.");
            AssertEqual(
                Word.WdLineSpacing.wdLineSpaceSingle,
                lastRange.ParagraphFormat.LineSpacingRule,
                $"{scenarioName}: selecting the compact terminal paragraph did not restore normal line spacing.");

            // Restore the compact structural state so the saved acceptance artifact
            // matches the idle document state seen by users after insertion.
            WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                document,
                tailFormulaId);
            AssertNumberedSpacingCleanup(document, scenarioName + "-recompact");
        }
        finally
        {
            Release(selection);
            Release(lastRange);
            Release(lastParagraph);
        }
    }

    private static void AssertSequentialNumberedInsertion(
        Word.Document document,
        string scenarioName)
    {
        AssertEqual(
            2,
            document.Tables.Count,
            $"{scenarioName}: the second numbered formula did not create an independent table.");
        AssertEqual(
            2,
            document.OMaths.Count,
            $"{scenarioName}: one native formula was swallowed during sequential insertion.");

        var tableRanges = new List<(int Start, int End)>();
        for (var index = 1; index <= document.Tables.Count; index++)
        {
            Word.Table? table = null;
            Word.Cell? centerCell = null;
            Word.Range? centerRange = null;
            Word.Range? tableRange = null;
            try
            {
                table = document.Tables[index];
                centerCell = table.Cell(1, 2);
                centerRange = centerCell.Range;
                AssertEqual(
                    1,
                    centerRange.OMaths.Count,
                    $"{scenarioName}: numbered table {index} does not own exactly one center-cell OMath.");
                tableRange = table.Range;
                tableRanges.Add((tableRange.Start, tableRange.End));
            }
            finally
            {
                Release(tableRange);
                Release(centerRange);
                Release(centerCell);
                Release(table);
            }
        }

        var captionRanges = new List<(int Start, int End)>();
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!bookmark.Name.StartsWith("VTEqCap_", StringComparison.Ordinal))
                        continue;
                    range = bookmark.Range;
                    captionRanges.Add((range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        tableRanges.Sort((left, right) => left.Start.CompareTo(right.Start));
        captionRanges.Sort((left, right) => left.Start.CompareTo(right.Start));
        AssertEqual(
            2,
            captionRanges.Count,
            $"{scenarioName}: expected two independent native SEQ caption bookmarks.");
        AssertTrue(
            tableRanges[0].End <= captionRanges[0].Start,
            $"{scenarioName}: the first native SEQ caption was inserted inside its formula table.");
        AssertTrue(
            captionRanges[0].End <= tableRanges[1].Start,
            $"{scenarioName}: the second formula was inserted before the first formula's native SEQ caption.");
        AssertTrue(
            tableRanges[1].End <= captionRanges[1].Start,
            $"{scenarioName}: the second native SEQ caption was inserted inside its formula table.");

        var sequenceResults = new List<(int Position, string Text)>();
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!(code.Text ?? string.Empty).TrimStart().StartsWith(
                            "SEQ ",
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    result = field.Result;
                    sequenceResults.Add((result.Start, (result.Text ?? string.Empty).Trim()));
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

        var orderedSequenceResults = sequenceResults
            .OrderBy(item => item.Position)
            .Select(item => item.Text)
            .ToArray();
        AssertEqual(
            2,
            orderedSequenceResults.Length,
            $"{scenarioName}: expected two native Equation SEQ fields.");
        AssertEqual(
            "1",
            orderedSequenceResults[0],
            $"{scenarioName}: the first formula was renumbered out of document order.");
        AssertEqual(
            "2",
            orderedSequenceResults[1],
            $"{scenarioName}: the second formula did not receive sequence number 2.");
    }
}
