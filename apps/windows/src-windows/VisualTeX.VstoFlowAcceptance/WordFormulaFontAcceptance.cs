using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordFormulaFontAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMath? math = null;
        Word.Range? equationRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            application.Selection.SetRange(0, 0);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = document.FullName,
                SourceObjectId = WordRangeReference(0, 0),
                Title = "VisualTeX formula font acceptance",
                CodeFormat = "latex",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = false,
                FontSizePt = 14,
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Latex = @"x+\text{中文}",
                    },
                },
                ExportResult = new OfficeExportDocument
                {
                    FormulaLetterFont = "times",
                    FormulaChineseFont = "kaiti",
                },
            };

            service.InsertOmml(
                session,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>+</mo><mtext>中文</mtext></math>");

            AssertEqual(1, document.OMaths.Count, "Word did not create exactly one OMML formula.");
            math = document.OMaths[1];
            equationRange = math.Range;
            font = equationRange.Font;

            var asciiName = ReadWordFontName(() => font.NameAscii);
            var farEastName = ReadWordFontName(() => font.NameFarEast);
            var generalName = ReadWordFontName(() => font.Name);
            Console.WriteLine(
                $"Word OMML font acceptance: Name={generalName}; NameAscii={asciiName}; NameFarEast={farEastName}; Size={font.Size:0.##}");

            AssertTrue(
                string.Equals(asciiName, "Times New Roman", StringComparison.OrdinalIgnoreCase)
                || string.Equals(generalName, "Times New Roman", StringComparison.OrdinalIgnoreCase),
                $"Word OMML did not retain the requested western formula font. Name={generalName}; NameAscii={asciiName}.");
            AssertTrue(
                string.Equals(farEastName, "KaiTi", StringComparison.OrdinalIgnoreCase)
                || farEastName.IndexOf("楷", StringComparison.Ordinal) >= 0,
                $"Word OMML did not retain the requested East Asia formula font. NameFarEast={farEastName}.");

            var savedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidOperationException("Word OMML metadata was not persisted.");
            AssertEqual(
                "times",
                savedMetadata.FormulaLetterFont,
                "Word OMML metadata lost formulaLetterFont.");
            AssertEqual(
                "kaiti",
                savedMetadata.FormulaChineseFont,
                "Word OMML metadata lost formulaChineseFont.");

            var artifactPath = Path.Combine(artifactRoot, "word-formula-fonts.docx");
            document.SaveAs2(artifactPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("Word formula font acceptance passed.");
        }
        finally
        {
            Release(font);
            Release(equationRange);
            Release(math);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static string ReadWordFontName(Func<string> read)
    {
        try { return read() ?? string.Empty; }
        catch { return string.Empty; }
    }
}
