using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOleNumberingMigration(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(artifactRoot, "word-ole-numbering-migration-assets");
        Directory.CreateDirectory(assetRoot);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Range? formulaRange = null;
        Word.InlineShape? shape = null;
        Word.Table? table = null;
        Word.Cell? leftCell = null;
        Word.Cell? centerCell = null;
        Word.Cell? rightCell = null;
        Word.Range? centerRange = null;
        Word.Range? rightRange = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();

            const string latex = "x^2+y^2=1";
            var formulaId = Guid.NewGuid().ToString("D");
            var svg = PerformanceSvg(901);
            var (pngPath, _) = CreatePerformanceOleAssets(
                assetRoot,
                901,
                latex,
                svg);
            var metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Title = "OLE numbering migration acceptance",
                Latex = latex,
                CodeFormat = "latex",
                DisplayMode = "block",
                Numbered = true,
                RenderWidthPx = ExportWidth,
                RenderHeightPx = ExportHeight,
                Baseline = ExportBaseline,
                FontSizePt = 11,
                RenderFontSizePt = 11,
                CreatedWithVersion = "1.2.6",
                UpdatedWithVersion = "1.2.6",
                CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
                UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
                },
            };
            metadata.Validate();

            insertion = document.Range(0, 0);
            object link = false;
            object save = true;
            object insertionObject = insertion;
            shape = document.InlineShapes.AddPicture(
                pngPath,
                ref link,
                ref save,
                ref insertionObject);
            shape.Width = ExportWidth * 0.75f;
            shape.Height = ExportHeight * 0.75f;
            WordFormulaMetadataReader.Write(shape, metadata);
            formulaRange = shape.Range.Duplicate;

            if (document.Tables.Count != 0)
                throw new InvalidDataException(
                    "Structural OLE surrogate did not start as a standalone display object.");

            // OLE and cross-platform picture formulas share the same non-OMML
            // structural migration branch. Using a metadata-tagged picture here
            // isolates the exact ConvertToTable/Columns.Add behavior without
            // depending on the native OLE server's cross-process security policy.
            WordEquationNumbering.ReconcileFormula(
                document,
                formulaRange,
                shape.Height,
                metadata);

            if (document.Tables.Count != 1)
                throw new InvalidDataException(
                    $"Numbered display migration should create exactly one table, actual {document.Tables.Count}.");
            if (document.InlineShapes.Count != 1)
                throw new InvalidDataException(
                    $"Numbered display migration lost or duplicated the formula object, actual {document.InlineShapes.Count}.");

            table = document.Tables[1];
            if (table.Columns.Count != 3)
                throw new InvalidDataException(
                    $"Numbered display table should have three columns, actual {table.Columns.Count}.");
            leftCell = table.Cell(1, 1);
            centerCell = table.Cell(1, 2);
            rightCell = table.Cell(1, 3);
            centerRange = centerCell.Range;
            rightRange = rightCell.Range;
            if (centerRange.InlineShapes.Count != 1)
                throw new InvalidDataException(
                    "Migrated display formula is not in the center cell.");
            if ((rightRange.Text ?? string.Empty).IndexOf("(1)", StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "Migrated display formula did not render visible equation number (1).");

            section = document.Sections[1];
            pageSetup = section.PageSetup;
            var availableWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
            var actualWidth = leftCell.Width + centerCell.Width + rightCell.Width;
            if (Math.Abs(actualWidth - availableWidth) > 2f)
                throw new InvalidDataException(
                    $"Migrated numbered table width is {actualWidth:F2} pt; expected {availableWidth:F2} pt.");
            if (Math.Abs(leftCell.Width - availableWidth * 0.2f) > 2f
                || Math.Abs(centerCell.Width - availableWidth * 0.6f) > 2f
                || Math.Abs(rightCell.Width - availableWidth * 0.2f) > 2f)
                throw new InvalidDataException(
                    $"Migrated numbered column widths are {leftCell.Width:F2}/{centerCell.Width:F2}/{rightCell.Width:F2} pt instead of 20/60/20 percent.");

            var artifactPath = Path.Combine(
                artifactRoot,
                "word-ole-unnumbered-to-numbered-structural.docx");
            document.SaveAs2(artifactPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "Word OLE numbering structural migration acceptance passed: the standalone non-OMML formula stayed visible after numbering and the 3-column table remained within the text width.");
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(rightRange);
            Release(centerRange);
            Release(rightCell);
            Release(centerCell);
            Release(leftCell);
            Release(table);
            Release(shape);
            Release(formulaRange);
            Release(insertion);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
