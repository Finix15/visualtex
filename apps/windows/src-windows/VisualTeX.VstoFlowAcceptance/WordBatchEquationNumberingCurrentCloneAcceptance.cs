using Microsoft.Office.Core;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBatchEquationNumberingCurrentClone(string artifactRoot)
    {
        if (!AttachActiveWord)
            throw new InvalidOperationException(
                "This acceptance requires VISUALTEX_VSTO_ACCEPTANCE_ATTACH_WORD=1.");

        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(
            artifactRoot,
            "word-batch-equation-numbering-current-clone.docx");
        DeleteBulkPerformanceArtifact(outputPath);

        Word.Application? application = null;
        Word.Document? source = null;
        Word.Document? clone = null;
        Word.Range? sourceContent = null;
        Word.Range? cloneContent = null;
        Word.Window? cloneWindow = null;
        CustomXMLParts? sourceParts = null;
        CustomXMLParts? cloneParts = null;
        try
        {
            application = CreateWordApplication(visible: true);
            source = application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document is available.");
            sourceContent = source.Content;
            var sourceXml = sourceContent.WordOpenXML;
            var sourceName = source.Name;
            var sourceSaved = source.Saved;
            var sourceMathCount = source.OMaths.Count;
            var sourceBookmarkCount = source.Bookmarks.Count;

            var visualTeXXmlParts = new List<string>();
            sourceParts = source.CustomXMLParts;
            for (var index = 1; index <= sourceParts.Count; index++)
            {
                CustomXMLPart? part = null;
                try
                {
                    part = sourceParts[index];
                    var xml = part.XML;
                    if (xml.IndexOf(
                            "urn:visualtex:word-omml:1",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || xml.IndexOf(
                            "<formula ",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        visualTeXXmlParts.Add(xml);
                }
                finally { Release(part); }
            }

            clone = application.Documents.Add();
            cloneWindow = clone.Windows[1];
            cloneWindow.Visible = false;
            cloneContent = clone.Content;
            cloneContent.InsertXML(sourceXml);
            cloneParts = clone.CustomXMLParts;
            foreach (var xml in visualTeXXmlParts)
            {
                CustomXMLPart? added = null;
                try { added = cloneParts.Add(xml); }
                finally { Release(added); }
            }

            Console.WriteLine(
                $"Cloned active Word document '{sourceName}': "
                + $"source OMaths={sourceMathCount}, clone OMaths={clone.OMaths.Count}, "
                + $"source bookmarks={sourceBookmarkCount}, clone bookmarks={clone.Bookmarks.Count}, "
                + $"VisualTeX XML parts={visualTeXXmlParts.Count}.");
            AssertEqual(sourceMathCount, clone.OMaths.Count,
                "The current-document clone did not preserve every OMath.");

            clone.Activate();
            if (string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_VSTO_BATCH_NUMBER_CLONE_RESET"),
                    "1",
                    StringComparison.Ordinal))
            {
                ResetCloneNumberedMetadata(clone);
                Console.WriteLine(
                    "Reset cloned VisualTeX display metadata to numbered=false "
                    + "to reproduce a first-time batch-numbering run.");
            }

            var service = new WordFormulaService(application);
            var before = service.GetBatchEquationNumberingStats();
            Console.WriteLine(
                $"Current-document clone batch stats: total={before.TotalCount}, "
                + $"numbered={before.NumberedCount}, unnumbered={before.UnnumberedCount}.");
            AssertTrue(before.TotalCount > 0,
                "The current-document clone did not expose any VisualTeX display formulas.");

            var result = service.ApplyBatchEquationNumbering(
                EquationNumberFormat.ContinuousId,
                redrawExisting: false);
            AssertEqual(before.TotalCount, result.NumberedCount,
                "Batch numbering did not number every formula in the current-document clone.");

            const string reportedFormulaId =
                "1066c9a9-14bd-4e20-9b67-09f0abc519bb";
            Word.Bookmark? reportedBookmark = null;
            Word.Range? reportedRange = null;
            try
            {
                reportedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    clone,
                    reportedFormulaId);
                if (reportedBookmark is not null)
                {
                    reportedRange = WordOmmlFormulaStore.GetEquationRange(reportedBookmark);
                    AssertTrue(
                        (bool)reportedRange.get_Information(
                            Word.WdInformation.wdWithInTable),
                        "The reported multi-line OMML formula was not migrated into a numbered table.");
                    Console.WriteLine(
                        $"Reported formula {reportedFormulaId} migrated at "
                        + $"{reportedRange.Start}-{reportedRange.End}; "
                        + $"tables={reportedRange.Tables.Count}.");
                }
            }
            finally
            {
                Release(reportedRange);
                Release(reportedBookmark);
            }

            clone.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"Current-document clone batch numbering passed: "
                + $"{result.NumberedCount} formulas numbered. Artifact: {outputPath}");

            source.Activate();
            AssertEqual(sourceName, source.Name,
                "The source document identity changed during clone acceptance.");
            AssertEqual(sourceSaved, source.Saved,
                "The source document saved state changed during clone acceptance.");
        }
        finally
        {
            try { source?.Activate(); } catch { }
            try { clone?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(cloneParts);
            Release(sourceParts);
            Release(cloneWindow);
            Release(cloneContent);
            Release(sourceContent);
            Release(clone);
            Release(source);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ResetCloneNumberedMetadata(Word.Document document)
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
                    if (metadata?.DisplayMode != "block") continue;
                    metadata.Numbered = false;
                    WordFormulaMetadataReader.Write(shape, metadata);
                }
                finally { Release(shape); }
            }

            foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block") continue;
                metadata.Numbered = false;
                WordOmmlFormulaStore.Save(document, metadata);
            }
        }
        finally { Release(shapes); }
    }
}
