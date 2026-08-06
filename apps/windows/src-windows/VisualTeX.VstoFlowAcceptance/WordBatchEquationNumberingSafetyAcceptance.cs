using System.Drawing;
using System.Windows.Forms;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBatchEquationNumberingSafety(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        RunBatchEquationNumberDialogLayout(artifactRoot);
        RunBatchEquationNumberingRollback(client, artifactRoot);
    }

    private static void RunBatchEquationNumberDialogLayout(string artifactRoot)
    {
        var cases = new[]
        {
            new
            {
                Name = "fill",
                Stats = new WordBatchEquationNumberingStats
                {
                    TotalCount = 40,
                    NumberedCount = 2,
                    CurrentFormatId = EquationNumberFormat.ContinuousId,
                },
                Redraw = false,
            },
            new
            {
                Name = "redraw",
                Stats = new WordBatchEquationNumberingStats
                {
                    TotalCount = 40,
                    NumberedCount = 40,
                    CurrentFormatId = EquationNumberFormat.Heading2DashId,
                },
                Redraw = true,
            },
        };

        foreach (var item in cases)
        {
            using var dialog = new BatchEquationNumberDialog(item.Stats, item.Redraw)
            {
                ShowInTaskbar = false,
                Opacity = 0d,
            };
            dialog.Show();
            Application.DoEvents();
            dialog.PerformLayout();
            AssertControlTreeFits(dialog, dialog.ClientRectangle, dialog.Text);

            using var bitmap = new Bitmap(
                Math.Max(1, dialog.ClientSize.Width),
                Math.Max(1, dialog.ClientSize.Height));
            dialog.DrawToBitmap(bitmap, dialog.ClientRectangle);
            var path = Path.Combine(
                artifactRoot,
                $"batch-equation-number-dialog-{item.Name}.png");
            bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            dialog.Close();
            Console.WriteLine(
                $"Batch-number dialog layout passed ({item.Name}): "
                + $"{dialog.ClientSize.Width}x{dialog.ClientSize.Height}. Artifact: {path}");
        }
    }

    private static void AssertControlTreeFits(
        Control parent,
        Rectangle clientBounds,
        string context)
    {
        foreach (Control child in parent.Controls)
        {
            if (!child.Visible) continue;
            AssertTrue(child.Left >= clientBounds.Left,
                $"{context}: {child.GetType().Name} extends left of its parent.");
            AssertTrue(child.Top >= clientBounds.Top,
                $"{context}: {child.GetType().Name} extends above its parent.");
            AssertTrue(child.Right <= clientBounds.Right + 1,
                $"{context}: {child.GetType().Name} is clipped on the right.");
            AssertTrue(child.Bottom <= clientBounds.Bottom + 1,
                $"{context}: {child.GetType().Name} is clipped at the bottom.");
            AssertControlTreeFits(
                child,
                new Rectangle(0, 0, child.ClientSize.Width, child.ClientSize.Height),
                context + "/" + child.GetType().Name);
        }
    }

    private static void RunBatchEquationNumberingRollback(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            var first = InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "r_1=1",
                numbered: false);
            AppendBatchAcceptanceParagraph(application, document);
            var second = InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "r_2=2",
                numbered: false);

            MoveFormulaIntoNonstandardTable(document, second.FormulaId);
            var service = new WordFormulaService(application);
            var beforeStats = service.GetBatchEquationNumberingStats();
            AssertEqual(2, beforeStats.TotalCount,
                "Rollback fixture did not contain two display formulas.");
            AssertEqual(0, beforeStats.NumberedCount,
                "Rollback fixture unexpectedly started with a numbered formula.");
            var beforeTables = document.Tables.Count;
            var beforeMaths = document.OMaths.Count;
            var beforeText = document.Content.Text;

            Exception? expected = null;
            try
            {
                service.ApplyBatchEquationNumbering(
                    EquationNumberFormat.ContinuousId,
                    redrawExisting: false);
            }
            catch (Exception error)
            {
                expected = error;
                Console.WriteLine(
                    "Expected batch-numbering failure observed: " + error.Message);
            }
            AssertTrue(expected is not null,
                "The invalid formula paragraph did not trigger batch-numbering rollback.");

            AssertEqual(beforeTables, document.Tables.Count,
                "Batch-numbering rollback left a table behind.");
            AssertEqual(beforeMaths, document.OMaths.Count,
                "Batch-numbering rollback changed the OMath count.");
            AssertEqual(beforeText, document.Content.Text,
                "Batch-numbering rollback changed document text.");

            foreach (var formula in new[] { first, second })
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"Rollback lost OMML metadata for {formula.FormulaId}.");
                AssertTrue(!metadata.Numbered,
                    $"Rollback left {formula.FormulaId} marked numbered=true.");
                AssertTrue(
                    !WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                        document,
                        formula.FormulaId),
                    $"Rollback left complete numbering artifacts for {formula.FormulaId}.");
            }

            var outputPath = Path.Combine(
                artifactRoot,
                "word-batch-equation-numbering-rollback.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "Batch-numbering rollback passed: structural changes and metadata "
                + $"were restored. Artifact: {outputPath}");
        }
        finally
        {
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
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void MoveFormulaIntoNonstandardTable(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? formatted = null;
        Word.Range? content = null;
        Word.Range? anchor = null;
        Word.Table? table = null;
        Word.Cell? cell = null;
        Word.Range? cellRange = null;
        Word.Range? insertion = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? migratedRange = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Rollback fixture could not find {formulaId}.");
            var metadata = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidDataException(
                    $"Rollback fixture could not read {formulaId} metadata.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            formatted = formulaRange.Duplicate;

            content = document.Content;
            var position = Math.Max(content.Start, content.End - 1);
            anchor = document.Range(position, position);
            table = document.Tables.Add(anchor, 1, 1);
            cell = table.Cell(1, 1);
            cellRange = cell.Range;
            insertion = cellRange.Duplicate;
            insertion.End = Math.Max(insertion.Start, insertion.End - 1);
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.FormattedText = formatted.FormattedText;

            paragraphRange.Delete();

            Release(cellRange);
            cellRange = cell.Range;
            maths = cellRange.OMaths;
            math = maths[1];
            migratedRange = math.Range;
            WordOmmlFormulaStore.Wrap(
                document,
                migratedRange,
                metadata,
                replaceExisting: true);
        }
        finally
        {
            Release(migratedRange);
            Release(math);
            Release(maths);
            Release(insertion);
            Release(cellRange);
            Release(cell);
            Release(table);
            Release(anchor);
            Release(content);
            Release(formatted);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(bookmark);
        }
    }
}
