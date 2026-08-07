using System.Diagnostics;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointDirectOmmlStressAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        try
        {
            var wordProcessesBefore = Process.GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .ToHashSet();
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var service = new PowerPointFormulaService(application);
            var timings = new List<double>();

            for (var iteration = 1; iteration <= 100; iteration++)
            {
                var formulaId = Guid.NewGuid().ToString();
                var session = CreatePowerPointAcceptanceSession(
                    "create",
                    "wordOmml",
                    formulaId,
                    sourceObjectId: null,
                    originalMetadata: null);
                var stopwatch = Stopwatch.StartNew();
                var result = service.InsertOmml(session);
                stopwatch.Stop();
                timings.Add(stopwatch.Elapsed.TotalMilliseconds);

                shape = slide.Shapes[result.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape),
                    $"Direct PowerPoint MathML import did not create native Office Math at iteration {iteration}.");
                shape.Select(MsoTriState.msoTrue);
                var readback = service.ReadSelection();
                AssertEqual(formulaId, readback.FormulaId,
                    $"Direct PowerPoint MathML import changed FormulaId at iteration {iteration}.");
                AssertContains(readback.Metadata?.Latex, @"\frac",
                    $"Direct PowerPoint MathML fraction readback failed at iteration {iteration}.");
                AssertContains(readback.Metadata?.Latex, @"\sqrt",
                    $"Direct PowerPoint MathML radical readback failed at iteration {iteration}.");
                shape.Delete();
                Release(shape);
                shape = null;
            }

            var wordProcessesAfter = Process.GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .ToHashSet();
            var spawnedWordProcesses = wordProcessesAfter.Except(wordProcessesBefore).ToArray();
            AssertEqual(0, spawnedWordProcesses.Length,
                "PowerPoint direct OMML conversion unexpectedly started Microsoft Word.");

            var ordered = timings.OrderBy(value => value).ToArray();
            var p50 = ordered[ordered.Length / 2];
            var p95 = ordered[(int)Math.Floor((ordered.Length - 1) * 0.95)];
            var maximum = ordered[ordered.Length - 1];
            Console.WriteLine(
                $"PowerPoint direct OMML stress passed: 100/100 native MathML imports with LaTeX readback, " +
                $"no new WINWORD process, p50={p50:F1} ms, p95={p95:F1} ms, max={maximum:F1} ms.");
        }
        finally
        {
            Release(shape);
            Release(slide);
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
            ForceComCleanup();
        }
    }
}
