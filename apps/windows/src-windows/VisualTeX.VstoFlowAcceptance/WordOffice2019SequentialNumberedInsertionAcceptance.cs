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
            Console.WriteLine(
                "Office 2019 sequential numbered insertion acceptance passed: "
                + "a captured create anchor survived live-selection drift, and a "
                + "caret captured inside an existing numbered table was redirected "
                + "after that formula's native SEQ paragraph.");
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
