using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

[ComVisible(true)]
[Guid("D4A1A3CB-0ED7-4B2F-8A2B-5CB0B1E25421")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWordRibbonCallbacks
{
    [DispId(1)]
    void OnRibbonLoad(object ribbonUi);

    [DispId(2)]
    void OnInsertInline(object control);

    [DispId(3)]
    void OnInsertDisplay(object control);

    [DispId(4)]
    void OnEditSelected(object control);

    [DispId(5)]
    void OnConvertSelected(object control);

    [DispId(6)]
    void OnUpdateEquationNumbers(object control);

    [DispId(7)]
    void OnExportSelectedAsPicture(object control);

    [DispId(8)]
    void OnDeleteSelected(object control);

    [DispId(9)]
    void OnOpenDesktop(object control);

    [DispId(10)]
    void OnInsertEquationReference(object control);

    [DispId(11)]
    void OnInsertInlineOmml(object control);

    [DispId(12)]
    void OnInsertDisplayOmml(object control);

    [DispId(13)]
    void OnConvertSelectedToOmml(object control);

    [DispId(14)]
    object? GetRibbonImage(Office.IRibbonControl control);

    [DispId(15)]
    string GetFormulaFontSizeText(Office.IRibbonControl control);

    [DispId(16)]
    bool GetFormulaFontSizeEnabled(Office.IRibbonControl control);

    [DispId(17)]
    void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value);

    [DispId(18)]
    void OnDecreaseFormulaFontSize(object control);

    [DispId(19)]
    void OnIncreaseFormulaFontSize(object control);

    [DispId(20)]
    void OnBulkImport(object control);

    [DispId(21)]
    void OnRedrawSelectionToOmml(object control);

    [DispId(22)]
    void OnRedrawSelectionToOle(object control);

    [DispId(23)]
    void OnRedrawDocumentToOmml(object control);

    [DispId(24)]
    void OnRedrawDocumentToOle(object control);

    [DispId(25)]
    bool GetEquationNumberFormatPressed(Office.IRibbonControl control);

    [DispId(26)]
    void OnEquationNumberFormatChanged(Office.IRibbonControl control, bool pressed);
}

[ComVisible(true)]
[Guid("F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1")]
[ProgId("VisualTeX.WordVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IWordRibbonCallbacks))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IWordRibbonCallbacks
{
    static ThisAddIn() => VstoDependencyResolver.Install();

    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="OLE 行内公式" size="large" tag="oleInline" getImage="GetRibbonImage" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="OLE 行间公式" size="large" tag="oleDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="OMML 行内公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行内公式。" tag="ommlInline" getImage="GetRibbonImage" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="OMML 行间公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行间公式。" tag="ommlDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <button id="VisualTeX.WordVsto.ConvertSelected" label="转为原生 OLE" screentip="转为可嵌入编辑的原生 OLE" supertip="转换后对象随 Word 文档保存，并可通过 VisualTeX 双击重新编辑。" tag="convertToOle" getImage="GetRibbonImage" onAction="OnConvertSelected" />
          <button id="VisualTeX.WordVsto.ConvertSelectedToOmml" label="转为 Word OMML" screentip="转为 Word 原生公式" supertip="将所选 VisualTeX 公式转换为 Word 原生 OMML；可在 Word 中直接编辑，也可继续用 VisualTeX 编辑。" tag="convertToOmml" getImage="GetRibbonImage" onAction="OnConvertSelectedToOmml" />
          <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新公式编号" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" />
          <menu id="VisualTeX.WordVsto.NumberFormat" label="编号格式" screentip="设置当前文档的公式编号格式" supertip="选择后立即更新当前文档已有的 VisualTeX 公式编号，并应用于后续新插入的带编号公式。">
            <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
          </menu>
          <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="引用带编号公式" supertip="从当前文档的带编号公式中选择目标，并插入可自动更新的 Word REF 字段。" imageMso="HyperlinkInsert" onAction="OnInsertEquationReference" />
          <button id="VisualTeX.WordVsto.ExportPicture" label="导出所选为图片" imageMso="PictureInsertFromFile" onAction="OnExportSelectedAsPicture" />
          <button id="VisualTeX.WordVsto.Delete" label="删除所选公式" imageMso="Delete" onAction="OnDeleteSelected" />
          <button id="VisualTeX.WordVsto.BulkImport" label="批量导入" size="large" screentip="批量导入 LaTeX / Markdown" supertip="将 Markdown 或 LaTeX 文档解析为 Word 原生文字，以及可单独编辑和调整字号的行内/行间公式。" tag="batchImport" getImage="GetRibbonImage" onAction="OnBulkImport" />
          <button id="VisualTeX.WordVsto.OpenDesktop" label="打开 VisualTeX" imageMso="FileOpen" onAction="OnOpenDesktop" />
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX 重绘">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="重绘所选" size="large" screentip="重绘所选 LaTeX 代码" supertip="识别所选文字中的 $...$、$$...$$、\\(...\\)、\\[...\\] 和常见公式环境，并原位替换为可编辑公式。" tag="batchImport" getImage="GetRibbonImage">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="重绘为 Word OMML" screentip="原位替换为 Word 原生公式" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="重绘为 VisualTeX OLE" screentip="原位替换为可双击编辑的 OLE" onAction="OnRedrawSelectionToOle" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="重绘全文" size="large" screentip="重绘整个文档中的 LaTeX 代码" supertip="扫描当前文档并原位替换全部 LaTeX 公式；开始前会再次确认。" imageMso="RefreshAll">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="全文重绘为 Word OMML" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="全文重绘为 VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="公式字号">
          <button id="VisualTeX.WordVsto.FontSizeDecrease" label="减小" imageMso="FontSizeDecrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
          <comboBox id="VisualTeX.WordVsto.FontSize" label="字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
            <item id="VisualTeX.WordVsto.FontSizeChu" label="初号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoChu" label="小初" />
            <item id="VisualTeX.WordVsto.FontSizeYi" label="一号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoYi" label="小一" />
            <item id="VisualTeX.WordVsto.FontSizeEr" label="二号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoEr" label="小二" />
            <item id="VisualTeX.WordVsto.FontSizeSan" label="三号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSan" label="小三" />
            <item id="VisualTeX.WordVsto.FontSizeSi" label="四号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSi" label="小四" />
            <item id="VisualTeX.WordVsto.FontSizeWu" label="五号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoWu" label="小五" />
            <item id="VisualTeX.WordVsto.FontSizeLiu" label="六号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoLiu" label="小六" />
            <item id="VisualTeX.WordVsto.FontSizeQi" label="七号" />
            <item id="VisualTeX.WordVsto.FontSizeBa" label="八号" />
            <item id="VisualTeX.WordVsto.FontSize8" label="8" />
            <item id="VisualTeX.WordVsto.FontSize9" label="9" />
            <item id="VisualTeX.WordVsto.FontSize10" label="10" />
            <item id="VisualTeX.WordVsto.FontSize10_5" label="10.5" />
            <item id="VisualTeX.WordVsto.FontSize11" label="11" />
            <item id="VisualTeX.WordVsto.FontSize12" label="12" />
            <item id="VisualTeX.WordVsto.FontSize14" label="14" />
            <item id="VisualTeX.WordVsto.FontSize16" label="16" />
            <item id="VisualTeX.WordVsto.FontSize18" label="18" />
            <item id="VisualTeX.WordVsto.FontSize20" label="20" />
            <item id="VisualTeX.WordVsto.FontSize24" label="24" />
            <item id="VisualTeX.WordVsto.FontSize28" label="28" />
            <item id="VisualTeX.WordVsto.FontSize36" label="36" />
            <item id="VisualTeX.WordVsto.FontSize48" label="48" />
            <item id="VisualTeX.WordVsto.FontSize72" label="72" />
          </comboBox>
          <button id="VisualTeX.WordVsto.FontSizeIncrease" label="增大" imageMso="FontSizeIncrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    private Application? _application;
    private WordFormulaService? _formulaService;
    private OfficeUiDispatcher? _dispatcher;
    private VisualTeXSessionClient? _sessionClient;
    private WordDoubleClickHook? _doubleClickHook;
    private static readonly object BulkAcceptanceLogGate = new();
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private readonly object _nativeOleTargetGate = new();
    private CancellationTokenSource? _lifetime;
    private string _lastDoubleClickFormulaId = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;
    private string? _activeSessionId;
    private bool _nativeOleTargetActive;
    private int _nativeOleTargetLeft;
    private int _nativeOleTargetTop;
    private int _nativeOleTargetRight;
    private int _nativeOleTargetBottom;
    private int _formulaFontInvalidationPending;
    private int _normalizingTypingCaret;
    private int _typingCaretNormalizationPending;
    private object? _ribbonUi;
    private Office.COMAddIn? _comAddIn;

    public string GetCustomUI(string ribbonId) => RibbonXml;

    public void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInstance,
        ref Array custom)
    {
        // Real VSTO-flow acceptances host the current source assembly manually
        // while Word may also auto-load the installed COM add-in into the same
        // application. Keep that installed instance inert only in acceptance so
        // one physical double-click cannot be handled by two event subscribers.
        // Production Word startup never sets this environment variable.
        if (addInInstance is Office.COMAddIn
            && string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            WordDoubleClickHook.TraceMessage(
                "installed-addin-suppressed-for-manual-acceptance");
            return;
        }

        _application = (Application)application;
        _comAddIn = addInInstance as Office.COMAddIn;
        if (_comAddIn is not null)
        {
            try { _comAddIn.Object = this; } catch { }
        }
        _formulaService = new WordFormulaService(_application);
        _dispatcher = new OfficeUiDispatcher();
        _sessionClient = new VisualTeXSessionClient();
        _lifetime = new CancellationTokenSource();
        _ = PrewarmCompanionAsync(_sessionClient, _lifetime.Token);
        _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
        _application.WindowSelectionChange += OnWindowSelectionChange;
        string? doubleClickError = null;
        try
        {
            _doubleClickHook = new WordDoubleClickHook(
                ShouldInterceptNativeOleDoubleClick,
                OnNativeOleDoubleClick);
            _doubleClickHook.Start();
        }
        catch (Exception error)
        {
            try { _doubleClickHook?.Dispose(); } catch { }
            _doubleClickHook = null;
            doubleClickError = error.Message;
        }
        SetStatus(doubleClickError is null
            ? "VisualTeX Word VSTO 已就绪。"
            : $"VisualTeX 已就绪，但 OLE 双击监听不可用：{doubleClickError}");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi)
    {
        _ribbonUi = ribbonUi;
        InvalidateFormulaFontControls();
        InvalidateEquationNumberFormatControls();
    }
    public object? GetRibbonImage(Office.IRibbonControl control) =>
        RibbonIconProvider.GetImage(control?.Tag);
    public string GetFormulaFontSizeText(Office.IRibbonControl control)
    {
        try
        {
            var size = _formulaService?.GetSelectedFormulaFontSize();
            return size.HasValue
                ? FormulaFontSize.FormatDisplay(size.Value)
                : string.Empty;
        }
        catch { return string.Empty; }
    }
    public bool GetFormulaFontSizeEnabled(Office.IRibbonControl control)
    {
        try { return _formulaService?.GetSelectedFormulaFontSize().HasValue == true; }
        catch { return false; }
    }
    public void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value) =>
        ApplyFormulaFontSize(ParseFontSize(value));
    public void OnDecreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");
            ApplyFormulaFontSize(FormulaFontSize.PreviousPreset(current));
        }
        catch (Exception error) { SetStatus($"无法设置公式字号：{error.Message}"); }
    }
    public void OnIncreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");
            ApplyFormulaFontSize(FormulaFontSize.NextPreset(current));
        }
        catch (Exception error) { SetStatus($"无法设置公式字号：{error.Message}"); }
    }
    public void OnInsertInline(object control) =>
        BeginSession("create", "inline", FormulaOleContract.NativeOleMode);
    public void OnInsertDisplay(object control) =>
        BeginSession("create", "block", FormulaOleContract.NativeOleMode);
    public void OnInsertInlineOmml(object control) =>
        BeginSession("create", "inline", FormulaOleContract.WordOmmlMode);
    public void OnInsertDisplayOmml(object control) =>
        BeginSession("create", "block", FormulaOleContract.WordOmmlMode);
    public void OnEditSelected(object control) => BeginSession("edit", null, null);
    public void OnConvertSelected(object control) =>
        BeginSession(
            "edit",
            null,
            FormulaOleContract.NativeOleMode,
            conversionOnly: true);
    public void OnConvertSelectedToOmml(object control) =>
        BeginSession(
            "edit",
            null,
            FormulaOleContract.WordOmmlMode,
            conversionOnly: true);
    public void OnUpdateEquationNumbers(object control) => _ = UpdateEquationNumbersAsync();
    public bool GetEquationNumberFormatPressed(Office.IRibbonControl control)
    {
        try
        {
            var current = _formulaService?.GetEquationNumberFormatId()
                ?? EquationNumberFormat.ContinuousId;
            return string.Equals(current, control?.Tag, StringComparison.Ordinal);
        }
        catch { return false; }
    }
    public void OnEquationNumberFormatChanged(
        Office.IRibbonControl control,
        bool pressed)
    {
        if (!pressed)
        {
            InvalidateEquationNumberFormatControls();
            return;
        }
        _ = SetEquationNumberFormatAsync(control?.Tag);
    }
    public void OnBulkImport(object control) => _ = BulkImportAsync();
    public void OnRedrawSelectionToOmml(object control) =>
        _ = RedrawLatexAsync(wholeDocument: false, FormulaOleContract.WordOmmlMode);
    public void OnRedrawSelectionToOle(object control) =>
        _ = RedrawLatexAsync(wholeDocument: false, FormulaOleContract.NativeOleMode);
    public void OnRedrawDocumentToOmml(object control) =>
        _ = RedrawLatexAsync(wholeDocument: true, FormulaOleContract.WordOmmlMode);
    public void OnRedrawDocumentToOle(object control) =>
        _ = RedrawLatexAsync(wholeDocument: true, FormulaOleContract.NativeOleMode);
    public void OnExportSelectedAsPicture(object control) => _ = ExportSelectedAsPictureAsync();
    public void OnDeleteSelected(object control) => _ = DeleteSelectedAsync();
    public void OnInsertEquationReference(object control) => _ = InsertEquationReferenceAsync();
    public void OnOpenDesktop(object control)
    {
        try
        {
            (_sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable."))
                .OpenDesktop();
            SetStatus("VisualTeX 已打开。");
        }
        catch (Exception error)
        {
            SetStatus($"无法打开 VisualTeX：{error.Message}");
        }
    }

    private static double ParseFontSize(string value) => FormulaFontSize.Parse(value);

    private void ApplyFormulaFontSize(double value)
    {
        try
        {
            var applied = (_formulaService
                    ?? throw new InvalidOperationException("Word formula service is unavailable."))
                .SetSelectedFormulaFontSize(value);
            SetStatus($"公式字号已设置为 {FormulaFontSize.Describe(applied)}。");
        }
        catch (Exception error)
        {
            SetStatus($"无法设置公式字号：{error.Message}");
        }
        finally { InvalidateFormulaFontControls(); }
    }

    private void ScheduleFormulaFontControlsInvalidation()
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null
            || Interlocked.Exchange(ref _formulaFontInvalidationPending, 1) != 0)
            return;
        dispatcher.Post(() =>
        {
            Interlocked.Exchange(ref _formulaFontInvalidationPending, 0);
            InvalidateFormulaFontControls();
            InvalidateEquationNumberFormatControls();
        });
    }

    private void InvalidateFormulaFontControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            ui.InvalidateControl("VisualTeX.WordVsto.FontSize");
            ui.InvalidateControl("VisualTeX.WordVsto.FontSizeDecrease");
            ui.InvalidateControl("VisualTeX.WordVsto.FontSizeIncrease");
        }
        catch { }
    }

    private void InvalidateEquationNumberFormatControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormat");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatContinuous");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading1Dot");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading1Dash");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading2Dot");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading2Dash");
        }
        catch { }
    }

    private void ScheduleTypingCaretNormalization()
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null
            || Interlocked.Exchange(ref _typingCaretNormalizationPending, 1) != 0)
            return;
        dispatcher.Post(() =>
        {
            Interlocked.Exchange(ref _typingCaretNormalizationPending, 0);
            var service = _formulaService;
            var application = _application;
            if (service is null || application is null) return;
            Selection? currentSelection = null;
            try
            {
                currentSelection = application.Selection;
                if (Interlocked.CompareExchange(ref _normalizingTypingCaret, 1, 0) != 0)
                    return;
                try { service.NormalizeTypingCaretAfterInlineFormula(currentSelection); }
                finally { Interlocked.Exchange(ref _normalizingTypingCaret, 0); }
            }
            catch { }
            finally { ReleaseComObject(currentSelection); }
        });
    }

    private void OnWindowSelectionChange(Selection selection)
    {
        // Defer Ribbon callbacks until Word finishes entering/leaving a native
        // math zone. Synchronous OMML inspection here can disturb its caret.
        ScheduleFormulaFontControlsInvalidation();
        var service = _formulaService;
        var application = _application;
        if (service is null || application is null)
        {
            ClearNativeOleTarget();
            return;
        }

        Range? range = null;
        Window? window = null;
        try
        {
            if (Interlocked.CompareExchange(ref _normalizingTypingCaret, 1, 0) == 0)
            {
                try { service.NormalizeTypingCaretAfterInlineFormula(selection); }
                finally { Interlocked.Exchange(ref _normalizingTypingCaret, 0); }
            }
            // During a mouse click Word can raise SelectionChange while the OLE
            // object is still selected, then collapse the caret at the object
            // tail without raising a second event. Re-check on the Office UI
            // queue after Word finishes that transition.
            ScheduleTypingCaretNormalization();

            // Do not inspect OMML metadata here. Word fires SelectionChange while
            // entering its native equation editor, and touching the OMath at that
            // point can disturb the caret state. Only perform the heavier metadata
            // read after the fast OLE type check succeeds.
            if (!service.IsSelectedNativeOle())
            {
                ClearNativeOleTarget();
                return;
            }
            var selected = service.ReadSelection(selection);
            if (!WordDoubleClickRouting.ShouldOpenVisualTeX(selected)
                || !string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                ClearNativeOleTarget();
                return;
            }

            range = selection.Range;
            window = application.ActiveWindow;
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            if (width <= 0 || height <= 0)
            {
                ClearNativeOleTarget();
                return;
            }
            const int padding = 4;
            lock (_nativeOleTargetGate)
            {
                _nativeOleTargetLeft = left - padding;
                _nativeOleTargetTop = top - padding;
                _nativeOleTargetRight = left + width + padding;
                _nativeOleTargetBottom = top + height + padding;
                _nativeOleTargetActive = true;
            }
            WordDoubleClickHook.TraceMessage(
                $"cache-active formulaId={selected.FormulaId} rect={left - padding},{top - padding},{left + width + padding},{top + height + padding}");
        }
        catch
        {
            ClearNativeOleTarget();
        }
        finally
        {
            ReleaseComObject(window);
            ReleaseComObject(range);
        }
    }

    private void OnWindowBeforeDoubleClick(Selection selection, ref bool cancel)
    {
        try
        {
            var selected = _formulaService?.ReadSelection(selection);
            if (selected?.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
                return;

            var shouldOpenVisualTeX = WordDoubleClickRouting.ShouldOpenVisualTeX(selected);
            WordDoubleClickHook.TraceMessage(
                $"window-before-double-click formulaId={selected.FormulaId} "
                + $"objectMode={selected.ObjectMode ?? "<null>"} "
                + $"shouldOpenVisualTeX={shouldOpenVisualTeX}");
            if (!shouldOpenVisualTeX) return;

            cancel = true;
            ClearNativeOleTarget();
            TryBeginDoubleClickSession(selected);
        }
        catch (Exception error)
        {
            SetStatus($"VisualTeX 双击检测失败：{error.Message}");
        }
    }

    private bool ShouldInterceptNativeOleDoubleClick(int screenX, int screenY)
    {
        lock (_nativeOleTargetGate)
        {
            return _nativeOleTargetActive
                && screenX >= _nativeOleTargetLeft
                && screenX <= _nativeOleTargetRight
                && screenY >= _nativeOleTargetTop
                && screenY <= _nativeOleTargetBottom;
        }
    }

    private void OnNativeOleDoubleClick()
    {
        WordDoubleClickHook.TraceMessage("addin-callback-received");
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            try
            {
                var selected = service.ReadSelection();
                WordDoubleClickHook.TraceMessage(
                    $"addin-selection formulaId={selected.FormulaId ?? "<null>"} objectMode={selected.ObjectMode ?? "<null>"}");
                if (!string.Equals(
                        selected.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                    return false;
                var started = TryBeginDoubleClickSession(selected);
                WordDoubleClickHook.TraceMessage($"addin-session-started={started}");
                return started;
            }
            catch (Exception error)
            {
                SetStatus($"VisualTeX OLE 双击检测失败：{error.Message}");
                return false;
            }
        });
    }

    private bool TryBeginDoubleClickSession(OfficeSelection? selected)
    {
        var formulaId = selected?.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId)
            || !WordDoubleClickRouting.ShouldOpenVisualTeX(selected))
            return false;

        var now = DateTimeOffset.UtcNow;
        if (formulaId == _lastDoubleClickFormulaId
            && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
            return false;
        _lastDoubleClickFormulaId = formulaId!;
        _lastDoubleClickAt = now;
        BeginSession("edit", null, null, capturedSelection: selected);
        return true;
    }

    private void ClearNativeOleTarget()
    {
        lock (_nativeOleTargetGate)
        {
            _nativeOleTargetActive = false;
            _nativeOleTargetLeft = 0;
            _nativeOleTargetTop = 0;
            _nativeOleTargetRight = 0;
            _nativeOleTargetBottom = 0;
        }
    }

    private static async Task PrewarmCompanionAsync(
        VisualTeXSessionClient client,
        CancellationToken cancellationToken)
    {
        try
        {
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            await client.PrewarmConverterAsync(cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            // Startup must remain non-blocking. The first explicit Office action
            // retries the full diagnostic/startup path and reports any failure.
        }
    }

    private void BeginSession(
        string mode,
        string? displayMode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection = null,
        bool conversionOnly = false)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(
            mode,
            displayMode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedDisplayMode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly,
        CancellationToken cancellationToken)
    {
        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                cancellationToken).ConfigureAwait(false))
        {
            var activeSessionId = Volatile.Read(ref _activeSessionId);
            if (!string.IsNullOrWhiteSpace(activeSessionId) && _sessionClient is not null)
            {
                try
                {
                    await _sessionClient.OpenEditorAsync(activeSessionId!, cancellationToken)
                        .ConfigureAwait(false);
                    SetStatus("已有 VisualTeX 编辑任务，已将编辑窗口切换到前台。");
                }
                catch (Exception error)
                {
                    SetStatus($"已有编辑任务，但无法置前窗口：{error.Message}");
                }
            }
            else
            {
                SetStatus("VisualTeX 正在准备编辑窗口，请稍候再试。");
            }
            return;
        }

        var openPerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal)
            ? Stopwatch.StartNew()
            : null;
        long openPerformanceCheckpoint = 0;
        void TraceOpenPerformance(string stage)
        {
            if (openPerformance is null) return;
            var elapsed = openPerformance.ElapsedMilliseconds;
            Console.WriteLine(
                $"    [perf] OpenSession.{stage}: +{elapsed - openPerformanceCheckpoint}ms ({elapsed}ms total)");
            openPerformanceCheckpoint = elapsed;
        }

        string? sessionId = null;
        string? imagePath = null;
        string? svgPath = null;
        string? emfPath = null;
        string? mathMl = null;
        try
        {
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("Word dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("Word formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            TraceOpenPerformance("health");
            var selection = capturedSelection?.Metadata is not null
                ? capturedSelection
                : await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            TraceOpenPerformance("read-selection");
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

            // A create command may be invoked while the previous formula is
            // still selected. Only edit commands are allowed to seed the new
            // Session from that selection; every create Session starts blank.
            var metadata = mode == "edit"
                ? NormalizeEditableMetadata(selection.Metadata)
                : null;
            var targetObjectMode = requestedObjectMode
                ?? (mode == "create" ? FormulaOleContract.NativeOleMode : selection.ObjectMode)
                ?? FormulaOleContract.NativeOleMode;
            var requiresObjectModeChange = mode == "edit"
                && !string.Equals(
                    selection.ObjectMode,
                    targetObjectMode,
                    StringComparison.Ordinal);
            var lines = metadata?.Lines ?? new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = string.Empty },
            };
            var fontSizePt = metadata?.FontSizePt
                ?? await dispatcher.InvokeAsync(service.ReadCurrentTypingFontSize)
                    .ConfigureAwait(false);
            var request = new CreateVstoSessionRequest
            {
                Mode = mode,
                Host = "word",
                FormulaId = metadata?.FormulaId,
                SourceDocumentId = selection.DocumentId,
                SourceObjectId = mode == "edit" ? selection.ObjectId : null,
                Title = metadata?.Title ?? "Word Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = requestedDisplayMode ?? metadata?.DisplayMode ?? "inline",
                ObjectMode = targetObjectMode,
                Numbered = (requestedDisplayMode ?? metadata?.DisplayMode) == "block"
                    && (metadata?.Numbered ?? false),
                FontSizePt = FormulaFontSize.Normalize(fontSizePt),
                OriginalMetadata = metadata,
                AutoCommitOnClose = true,
            };
            TraceOpenPerformance("build-request");
            var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
            TraceOpenPerformance("create-session");
            sessionId = session.Id;
            Volatile.Write(ref _activeSessionId, session.Id);
            if (conversionOnly)
            {
                await client.OpenConverterAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                SetStatus("正在直接转换 Word 公式格式…");
            }
            else
            {
                await client.OpenEditorAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                TraceOpenPerformance("open-editor-window");
                SetStatus("VisualTeX 编辑器已打开。");
            }
            session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromMinutes(30),
                cancellationToken).ConfigureAwait(false);
            if (session.Status == "cancelled" || session.ExplicitCancel)
            {
                SetStatus("已取消，Word 文档未修改。");
                return;
            }
            if (session.Status == "failed")
                throw new InvalidOperationException(session.Error ?? "VisualTeX Session 失败。");
            if (session.Mode == "edit"
                && !session.Dirty
                && (!requiresObjectModeChange || session.ExportResult is null))
            {
                await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
                SetStatus(requiresObjectModeChange
                    ? "未执行对象格式转换。"
                    : "公式内容未变化。");
                return;
            }

            var export = session.ExportResult
                ?? throw new InvalidOperationException("VisualTeX Session has no export result.");
            if (string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                var requiredMathMl = export.MathMl;
                if (string.IsNullOrWhiteSpace(requiredMathMl)
                    || !requiredMathMl!.TrimStart().StartsWith("<math", StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "VisualTeX Session has no valid MathML result for Word OMML.");
                mathMl = requiredMathMl;
            }
            else
            {
                imagePath = client.MaterializePng(session);
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    svgPath = client.MaterializeSvg(session);
                    emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                        svgPath,
                        export.Width,
                        export.Height);
                }
            }
            await dispatcher.InvokeAsync(() =>
            {
                var activeDocumentId = service.ReadActiveDocumentId();
                if (!string.Equals(
                        activeDocumentId,
                        session.SourceDocumentId,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("活动 Word 文档已切换，未写入公式。");
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    if (mathMl is null)
                        throw new InvalidOperationException(
                            "VisualTeX Word OMML MathML payload is unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOmml(session, mathMl)
                        : service.InsertOmml(session, mathMl);
                }
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    if (emfPath is null || imagePath is null)
                        throw new InvalidOperationException(
                            "VisualTeX native OLE previews are unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOle(session, imagePath, emfPath)
                        : service.InsertOle(session, imagePath, emfPath);
                }
                if (imagePath is null)
                    throw new InvalidOperationException(
                        "VisualTeX picture preview is unavailable.");
                return session.Mode == "edit"
                    ? service.Replace(session, imagePath)
                    : service.Insert(session, imagePath);
            }).ConfigureAwait(false);
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            if (requiresObjectModeChange
                && string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                SetStatus("已转换为 Word 原生 OMML：可在 Word 中直接编辑，也可继续用 VisualTeX 编辑。");
            else if (requiresObjectModeChange
                && string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
                SetStatus("已转换为原生 OLE：可双击使用 VisualTeX 编辑，并随 Word 文档保存。");
            else
                SetStatus(session.Mode == "edit" ? "Word 公式已更新。" : "Word 公式已插入。");
        }
        catch (OperationCanceledException)
        {
            SetStatus("VisualTeX 操作已取消。");
        }
        catch (Exception error)
        {
            if (sessionId is not null && _sessionClient is not null)
            {
                try
                {
                    await _sessionClient.FailAsync(sessionId, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus($"VisualTeX Word 写入失败：{error.Message}");
        }
        finally
        {
            if (emfPath is not null)
            {
                try { File.Delete(emfPath); } catch { }
            }
            if (svgPath is not null)
            {
                try { File.Delete(svgPath); } catch { }
            }
            if (imagePath is not null)
            {
                try { File.Delete(imagePath); } catch { }
            }
            if (sessionId is not null
                && string.Equals(
                    Volatile.Read(ref _activeSessionId),
                    sessionId,
                    StringComparison.Ordinal))
                Volatile.Write(ref _activeSessionId, null);
            _operationGate.Release();
        }
    }

    private static FormulaMetadata? NormalizeEditableMetadata(FormulaMetadata? source)
    {
        if (source is null) return null;
        var metadata = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(source))
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
        if (metadata.Lines.Count == 0) return metadata;

        var last = metadata.Lines[metadata.Lines.Count - 1];
        var split = FormulaEquationTag.Extract(last.Latex);
        if (!string.Equals(last.Latex, split.Latex, StringComparison.Ordinal))
            last.Latex = split.Latex;
        metadata.EquationTag ??= split.EquationTag;
        metadata.Latex = string.Join("\n", metadata.Lines.Select(line => line.Latex));
        if (!string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            metadata.EquationTag = null;
        metadata.Validate();
        return metadata;
    }

    private async Task RedrawLatexAsync(bool wholeDocument, string objectMode)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || client is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        var maxRenderMilliseconds = 0L;
        var totalRenderMilliseconds = 0L;
        try
        {
            var plan = await dispatcher.InvokeAsync(
                    () => service.CaptureLatexRedrawPlan(wholeDocument))
                .ConfigureAwait(false);
            var modeLabel = string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
                ? "VisualTeX OLE"
                : "Word OMML";
            if (wholeDocument
                && !string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                var confirmed = await dispatcher.InvokeAsync(() =>
                    System.Windows.Forms.MessageBox.Show(
                        $"将在整个文档中原位重绘 {plan.Targets.Count} 个 LaTeX 公式为 {modeLabel}。\r\n\r\n"
                        + "该操作会保留正文并删除公式两侧的 LaTeX 定界符，可通过一次 Ctrl+Z 整体撤销。是否继续？",
                        "VisualTeX LaTeX 重绘",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes).ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus("已取消全文 LaTeX 重绘，Word 文档未修改。");
                    return;
                }
            }

            WriteRedrawAcceptanceLog(
                $"redraw-start scope={(wholeDocument ? "document" : "selection")} "
                + $"mode={objectMode} formulas={plan.Targets.Count}");
            SetStatus($"正在准备重绘 {plan.Targets.Count} 个 LaTeX 公式为 {modeLabel}…");
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);

            for (var index = 0; index < plan.Targets.Count; index++)
            {
                lifetime.Token.ThrowIfCancellationRequested();
                var target = plan.Targets[index];
                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
                var key = string.Join(
                    "\u001F",
                    objectMode,
                    target.DisplayMode,
                    target.FontSizePt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    target.Latex);
                if (!rendered.TryGetValue(key, out var template))
                {
                    SetStatus($"正在渲染公式 {index + 1}/{plan.Targets.Count}…");
                    var stopwatch = Stopwatch.StartNew();
                    template = await RenderBulkFormulaTemplateAsync(
                            client,
                            run,
                            objectMode,
                            plan.DocumentId,
                            target.FontSizePt,
                            lifetime.Token)
                        .ConfigureAwait(false);
                    stopwatch.Stop();
                    totalRenderMilliseconds += stopwatch.ElapsedMilliseconds;
                    maxRenderMilliseconds = Math.Max(
                        maxRenderMilliseconds,
                        stopwatch.ElapsedMilliseconds);
                    WriteRedrawAcceptanceLog(
                        $"render index={index + 1} elapsedMs={stopwatch.ElapsedMilliseconds} "
                        + $"fontSizePt={target.FontSizePt:0.##} display={target.DisplayMode} "
                        + $"latex={target.Latex}");
                    rendered.Add(key, template);
                    converterSessionIds.Add(template.Session.Id);
                }
                else
                {
                    WriteRedrawAcceptanceLog(
                        $"render-cache-hit index={index + 1} display={target.DisplayMode} "
                        + $"latex={target.Latex}");
                }

                var independentSession = CloneBulkFormulaSession(
                    template.Session,
                    run,
                    plan.DocumentId,
                    target.FontSizePt,
                    objectMode);
                prepared.Add(target.Id, new PreparedWordBulkFormula
                {
                    Run = run,
                    Session = independentSession,
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                });
            }

            SetStatus("公式渲染完成，正在原位写入 Word…");
            var result = await dispatcher.InvokeAsync(
                    () => service.ApplyLatexRedrawPlan(plan, prepared))
                .ConfigureAwait(false);
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.CompleteAsync(sessionId, lifetime.Token)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            var uniqueRenderCount = Math.Max(1, rendered.Count);
            var averageRenderMilliseconds = totalRenderMilliseconds / uniqueRenderCount;
            WriteRedrawAcceptanceLog(
                $"redraw-complete formulas={result.FormulaCount} unique={rendered.Count} "
                + $"renderAverageMs={averageRenderMilliseconds} renderMaxMs={maxRenderMilliseconds} "
                + $"insertTotalMs={result.TotalInsertMilliseconds} insertMaxMs={result.MaxInsertMilliseconds}");
            var performanceSuffix = maxRenderMilliseconds <= 250
                ? $"渲染最大 {maxRenderMilliseconds} ms/公式"
                : $"渲染最大 {maxRenderMilliseconds} ms/公式（本机超过 250 ms 目标）";
            SetStatus(
                $"LaTeX 重绘完成：{result.FormulaCount} 个公式已转换为 {modeLabel}；{performanceSuffix}。");
        }
        catch (OperationCanceledException error)
        {
            WriteRedrawAcceptanceLog("redraw-cancelled " + error);
            SetStatus("VisualTeX LaTeX 重绘已取消。");
        }
        catch (Exception error)
        {
            WriteRedrawAcceptanceLog("redraw-failed " + error);
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.FailAsync(sessionId, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus($"VisualTeX LaTeX 重绘失败：{error.Message}");
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            error.Message,
                            "VisualTeX LaTeX 重绘",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }
        finally
        {
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            _operationGate.Release();
        }
    }

    private async Task BulkImportAsync()
    {
        WriteBulkAcceptanceLog("bulk-import-start");
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var application = _application;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || client is null
            || application is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        var operationStopwatch = Stopwatch.StartNew();
        string? bulkImportSessionId = null;
        try
        {
            var selection = await dispatcher.InvokeAsync(service.ReadSelection)
                .ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            var sourceDocumentId = selection.DocumentId;
            var fontSizePt = FormulaFontSize.Normalize(
                await dispatcher.InvokeAsync(service.ReadCurrentTypingFontSize)
                    .ConfigureAwait(false));
            var resolvedImport = await ResolveBulkImportDocumentAsync(
                    client,
                    sourceDocumentId,
                    selection.ObjectId,
                    fontSizePt,
                    lifetime.Token)
                .ConfigureAwait(false);
            bulkImportSessionId = resolvedImport.SessionId;
            var document = resolvedImport.Document;
            if (document is null)
            {
                WriteBulkAcceptanceLog("bulk-import-cancelled-no-document");
                SetStatus("已取消批量导入，Word 文档未修改。");
                return;
            }

            WriteBulkAcceptanceLog(
                $"parsed blocks={document.Blocks.Count} formulas={document.FormulaCount} "
                + $"mode={document.FormulaObjectMode} fontSizePt={fontSizePt:0.##}");
            SetStatus(
                $"正在准备批量导入：{document.Blocks.Count} 个块，{document.FormulaCount} 个公式…");
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);
            var formulaRuns = document.Blocks
                .SelectMany(block => block.Runs)
                .Where(run => run.IsFormula)
                .ToList();
            var objectMode = document.FormulaObjectMode == WordBulkFormulaObjectMode.Ole
                ? FormulaOleContract.NativeOleMode
                : FormulaOleContract.WordOmmlMode;
            var pendingKeys = new HashSet<string>(StringComparer.Ordinal);
            var formulaKeys = new Dictionary<string, string>(StringComparer.Ordinal);
            var pendingTemplates = new List<(
                string Key,
                WordBulkRun Run,
                OfficeSessionDocument Session)>();

            for (var index = 0; index < formulaRuns.Count; index++)
            {
                lifetime.Token.ThrowIfCancellationRequested();
                var run = formulaRuns[index];
                var key = string.Join(
                    "\u001F",
                    objectMode,
                    run.DisplayMode,
                    fontSizePt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    run.Latex,
                    run.EquationTag ?? string.Empty);
                formulaKeys.Add(run.Id, key);
                if (rendered.ContainsKey(key) || !pendingKeys.Add(key))
                    continue;

                WriteBulkAcceptanceLog(
                    $"render-prepare index={index + 1}/{formulaRuns.Count} "
                    + $"display={run.DisplayMode} latex={run.Latex}");
                SetStatus($"正在准备公式 {index + 1}/{formulaRuns.Count}…");
                var conversionSession = await CreateBulkFormulaConversionSessionAsync(
                        client,
                        run,
                        objectMode,
                        sourceDocumentId,
                        fontSizePt,
                        lifetime.Token)
                    .ConfigureAwait(false);
                pendingTemplates.Add((key, run, conversionSession));
                converterSessionIds.Add(conversionSession.Id);
            }

            if (pendingTemplates.Count > 0)
            {
                WriteBulkAcceptanceLog(
                    $"render-batch-start unique={pendingTemplates.Count} total={formulaRuns.Count}");
                SetStatus($"正在批量渲染 {pendingTemplates.Count} 个独立公式…");
                var renderStopwatch = Stopwatch.StartNew();
                await client.OpenConverterBatchAsync(
                        pendingTemplates.Select(item => item.Session.Id).ToList(),
                        lifetime.Token)
                    .ConfigureAwait(false);
                foreach (var pending in pendingTemplates)
                {
                    var completedSession = await client.WaitForCommitAsync(
                            pending.Session.Id,
                            TimeSpan.FromMinutes(3),
                            lifetime.Token)
                        .ConfigureAwait(false);
                    var template = MaterializeBulkFormulaTemplate(
                        client,
                        pending.Run,
                        objectMode,
                        completedSession);
                    rendered.Add(pending.Key, template);
                    WriteBulkAcceptanceLog(
                        $"render-batch-item sessionId={completedSession.Id} "
                        + $"status={completedSession.Status} display={pending.Run.DisplayMode}");
                }
                renderStopwatch.Stop();
                WriteBulkAcceptanceLog(
                    $"render-batch-complete unique={pendingTemplates.Count} "
                    + $"elapsedMs={renderStopwatch.ElapsedMilliseconds}");
            }

            foreach (var run in formulaRuns)
            {
                var key = formulaKeys[run.Id];
                var template = rendered[key];
                var independentSession = CloneBulkFormulaSession(
                    template.Session,
                    run,
                    sourceDocumentId,
                    fontSizePt,
                    objectMode);
                prepared.Add(run.Id, new PreparedWordBulkFormula
                {
                    Run = run,
                    Session = independentSession,
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                });
            }

            SetStatus("公式渲染完成，正在写入 Word…");
            var insertStopwatch = Stopwatch.StartNew();
            var result = await dispatcher.InvokeAsync(() =>
                    service.InsertBulkDocument(
                        document,
                        prepared,
                        sourceDocumentId,
                        selection.ObjectId))
                .ConfigureAwait(false);
            insertStopwatch.Stop();
            WriteBulkAcceptanceLog(
                $"bulk-insert-complete blocks={result.BlockCount} "
                + $"formulas={result.FormulaCount} elapsedMs={insertStopwatch.ElapsedMilliseconds}");
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.CompleteAsync(sessionId, lifetime.Token)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            if (!string.IsNullOrWhiteSpace(bulkImportSessionId))
            {
                try
                {
                    await client.CompleteAsync(bulkImportSessionId!, lifetime.Token)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            operationStopwatch.Stop();
            WriteBulkAcceptanceLog(
                $"bulk-import-complete blocks={result.BlockCount} formulas={result.FormulaCount} "
                + $"elapsedMs={operationStopwatch.ElapsedMilliseconds}");
            SetStatus(
                $"批量导入完成：{result.BlockCount} 个内容块，{result.FormulaCount} 个独立公式；"
                + $"耗时 {operationStopwatch.Elapsed.TotalSeconds:0.0} 秒。");
        }
        catch (OperationCanceledException error)
        {
            WriteBulkAcceptanceLog("bulk-import-cancelled " + error);
            SetStatus("VisualTeX 批量导入已取消。");
        }
        catch (Exception error)
        {
            WriteBulkAcceptanceLog("bulk-import-failed " + error);
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.FailAsync(sessionId, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            if (!string.IsNullOrWhiteSpace(bulkImportSessionId))
            {
                try
                {
                    await client.FailAsync(bulkImportSessionId!, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus($"VisualTeX 批量导入失败：{error.Message}");
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            error.Message,
                            "VisualTeX 批量导入",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }
        finally
        {
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            _operationGate.Release();
        }
    }

    private async Task<(WordBulkImportDocument? Document, string? SessionId)>
        ResolveBulkImportDocumentAsync(
            VisualTeXSessionClient client,
            string? sourceDocumentId,
            string? sourceObjectId,
            double fontSizePt,
            CancellationToken cancellationToken)
    {
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            var sourcePath = Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_BULK_SOURCE_PATH");
            WriteBulkAcceptanceLog(
                $"resolve-acceptance-source path={sourcePath ?? "<null>"} "
                + $"format={Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT") ?? "<null>"} "
                + $"mode={Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE") ?? "<null>"}");
            var acceptanceSource = !string.IsNullOrWhiteSpace(sourcePath)
                ? File.ReadAllText(sourcePath!, Encoding.UTF8)
                : Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE")
                  ?? throw new InvalidOperationException(
                      "Acceptance bulk import requires VISUALTEX_VSTO_BULK_SOURCE_PATH or VISUALTEX_VSTO_BULK_SOURCE.");
            var acceptanceFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT")
                ?.Trim().ToLowerInvariant() switch
            {
                "markdown" => WordBulkSourceFormat.Markdown,
                "latex" => WordBulkSourceFormat.Latex,
                _ => WordBulkSourceFormat.Auto,
            };
            var acceptanceObjectMode = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE"),
                "ole",
                StringComparison.OrdinalIgnoreCase)
                ? WordBulkFormulaObjectMode.Ole
                : WordBulkFormulaObjectMode.Omml;
            return (
                WordBulkImportParser.Parse(
                    acceptanceSource,
                    acceptanceFormat,
                    acceptanceObjectMode),
                null);
        }

        await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = string.Empty,
        };
        var session = await client.CreateSessionAsync(
            new CreateVstoSessionRequest
            {
                Mode = "create",
                Host = "word",
                SourceDocumentId = sourceDocumentId,
                SourceObjectId = sourceObjectId,
                Title = "Word 文档批量导入",
                Lines = new List<FormulaLine> { line },
                ActiveLineId = line.Id,
                CodeFormat = "auto-document",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = false,
                FontSizePt = fontSizePt,
                AutoCommitOnClose = false,
            },
            cancellationToken).ConfigureAwait(false);
        WriteBulkAcceptanceLog($"bulk-import-ui-created sessionId={session.Id}");
        await client.OpenBulkImportAsync(session.Id, cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog($"bulk-import-ui-opened sessionId={session.Id}");
        session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromHours(1),
                cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog(
            $"bulk-import-ui-finished sessionId={session.Id} status={session.Status} "
            + $"error={session.Error ?? "<null>"}");
        if (session.Status == "cancelled" || session.ExplicitCancel)
            return (null, session.Id);
        if (session.Status == "failed")
            throw new InvalidOperationException(
                session.Error ?? "VisualTeX 文档导入窗口返回失败状态。");
        if (session.Status is not ("committing" or "completed"))
            throw new InvalidOperationException(
                $"VisualTeX 文档导入窗口返回了意外状态：{session.Status}。");

        var source = string.Join("\n", session.Lines.Select(item => item.Latex));
        var objectMode = string.Equals(
            session.ObjectMode,
            FormulaOleContract.NativeOleMode,
            StringComparison.Ordinal)
            ? WordBulkFormulaObjectMode.Ole
            : WordBulkFormulaObjectMode.Omml;
        if (string.Equals(
                session.CodeFormat,
                "visualtex-document-json",
                StringComparison.OrdinalIgnoreCase))
        {
            return (
                WordBulkImportParser.ParseSerialized(source, objectMode),
                session.Id);
        }
        var format = session.CodeFormat.Trim().ToLowerInvariant() switch
        {
            "markdown-document" => WordBulkSourceFormat.Markdown,
            "latex-document" => WordBulkSourceFormat.Latex,
            _ => WordBulkSourceFormat.Auto,
        };
        return (
            WordBulkImportParser.Parse(source, format, objectMode),
            session.Id);
    }

    private static async Task<RenderedWordBulkFormulaTemplate> RenderBulkFormulaTemplateAsync(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        string? sourceDocumentId,
        double fontSizePt,
        CancellationToken cancellationToken)
    {
        var session = await CreateBulkFormulaConversionSessionAsync(
                client,
                run,
                objectMode,
                sourceDocumentId,
                fontSizePt,
                cancellationToken)
            .ConfigureAwait(false);
        await client.OpenConverterAsync(session.Id, cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog($"converter-opened sessionId={session.Id}");
        session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromMinutes(3),
                cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog(
            $"converter-finished sessionId={session.Id} status={session.Status} "
            + $"error={session.Error ?? "<null>"}");
        return MaterializeBulkFormulaTemplate(client, run, objectMode, session);
    }

    private static async Task<OfficeSessionDocument> CreateBulkFormulaConversionSessionAsync(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        string? sourceDocumentId,
        double fontSizePt,
        CancellationToken cancellationToken)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = run.Latex,
        };
        var originalMetadata = CreateBulkFormulaMetadata(
            formulaId,
            line,
            run,
            fontSizePt);
        WriteBulkAcceptanceLog(
            $"converter-create display={run.DisplayMode} mode={objectMode} latex={run.Latex}");
        var session = await client.CreateSessionAsync(
            new CreateVstoSessionRequest
            {
                Mode = "create",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = sourceDocumentId,
                Title = "Bulk imported Word formula",
                Lines = new List<FormulaLine> { line },
                ActiveLineId = line.Id,
                CodeFormat = "latex",
                DisplayMode = run.DisplayMode,
                ObjectMode = objectMode,
                Numbered = false,
                FontSizePt = fontSizePt,
                OriginalMetadata = originalMetadata,
                AutoCommitOnClose = false,
            },
            cancellationToken).ConfigureAwait(false);
        WriteBulkAcceptanceLog($"converter-created sessionId={session.Id}");
        return session;
    }

    private static RenderedWordBulkFormulaTemplate MaterializeBulkFormulaTemplate(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        OfficeSessionDocument session)
    {
        if (session.Status == "failed")
        {
            var detail = string.IsNullOrWhiteSpace(session.Error)
                || string.Equals(session.Error, "[object Object]", StringComparison.Ordinal)
                ? "MathJax 无法解析该公式，转换窗口没有返回有效错误文本。"
                : session.Error!.Trim();
            var formula = run.Latex.Length <= 500
                ? run.Latex
                : run.Latex.Substring(0, 500) + "…";
            throw new InvalidOperationException(
                $"公式渲染失败：{formula}\r\n原因：{detail}");
        }
        if (session.Status == "cancelled" || session.ExplicitCancel)
            throw new OperationCanceledException("批量公式渲染已取消。" );
        var export = session.ExportResult
            ?? throw new InvalidOperationException(
                $"公式 {run.Latex} 没有生成导出结果。" );
        if (string.IsNullOrWhiteSpace(export.MathMl))
            throw new InvalidDataException(
                $"公式 {run.Latex} 没有生成 MathML。" );

        var template = new RenderedWordBulkFormulaTemplate
        {
            Session = session,
            MathMl = export.MathMl,
        };
        if (string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal))
        {
            template.PngPath = client.MaterializePng(session);
            template.SvgPath = client.MaterializeSvg(session);
            template.EmfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                template.SvgPath,
                export.Width,
                export.Height);
        }
        return template;
    }

    private static FormulaMetadata CreateBulkFormulaMetadata(
        string formulaId,
        FormulaLine line,
        WordBulkRun run,
        double fontSizePt)
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        return new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Bulk imported Word formula",
            Latex = line.Latex,
            Lines = new List<FormulaLine> { line },
            CodeFormat = "latex",
            DisplayMode = run.DisplayMode,
            Numbered = false,
            EquationTag = run.DisplayMode == "block" ? run.EquationTag : null,
            FontSizePt = FormulaFontSize.Normalize(fontSizePt),
            RenderFontSizePt = FormulaFontSize.Normalize(fontSizePt),
            CreatedWithVersion = "1.2.4",
            UpdatedWithVersion = "1.2.4",
            CreatedAt = now,
            UpdatedAt = now,
        };
    }

    private static OfficeSessionDocument CloneBulkFormulaSession(
        OfficeSessionDocument template,
        WordBulkRun run,
        string? sourceDocumentId,
        double fontSizePt,
        string objectMode)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = run.Latex,
        };
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = sourceDocumentId,
            Title = "Bulk imported Word formula",
            Lines = new List<FormulaLine> { line },
            CodeFormat = "latex",
            DisplayMode = run.DisplayMode,
            ObjectMode = objectMode,
            Numbered = false,
            FontSizePt = fontSizePt,
            Status = "committing",
            Dirty = true,
            OriginalMetadata = CreateBulkFormulaMetadata(
                formulaId,
                line,
                run,
                fontSizePt),
            ExportResult = template.ExportResult,
        };
    }

    private static void WriteRedrawAcceptanceLog(string message)
    {
        var path = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG");
        if (string.IsNullOrWhiteSpace(path)) return;
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            lock (BulkAcceptanceLogGate)
            {
                File.AppendAllText(
                    path!,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
    }

    private static void WriteBulkAcceptanceLog(string message)
    {
        var path = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");
        if (string.IsNullOrWhiteSpace(path)) return;
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            lock (BulkAcceptanceLogGate)
            {
                File.AppendAllText(
                    path!,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        try { File.Delete(path!); } catch { }
    }

    private async Task UpdateEquationNumbersAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            var count = await dispatcher.InvokeAsync(service.UpdateEquationNumbers)
                .ConfigureAwait(false);
            SetStatus($"已更新 {count} 个 Word 公式编号。");
        }
        catch (Exception error)
        {
            SetStatus($"更新 Word 公式编号失败：{error.Message}");
        }
    }

    private async Task SetEquationNumberFormatAsync(string? requestedFormatId)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        var format = EquationNumberFormat.Resolve(requestedFormatId);
        try
        {
            var current = await dispatcher.InvokeAsync(service.GetEquationNumberFormatId)
                .ConfigureAwait(false);
            if (string.Equals(current, format.Id, StringComparison.Ordinal))
            {
                SetStatus($"当前公式编号格式已是“{format.DisplayName}”。");
                return;
            }

            var count = await dispatcher.InvokeAsync(
                    () => service.SetEquationNumberFormat(format.Id))
                .ConfigureAwait(false);
            SetStatus($"公式编号格式已设置为“{format.DisplayName}”，并更新了 {count} 个带编号公式。");
        }
        catch (Exception error)
        {
            SetStatus($"设置公式编号格式失败：{error.Message}");
        }
        finally { InvalidateEquationNumberFormatControls(); }
    }

    private async Task InsertEquationReferenceAsync()
    {
        var dispatcher = _dispatcher;
        var application = _application;
        if (dispatcher is null || application is null) return;
        try
        {
            var inserted = await dispatcher.InvokeAsync(() =>
            {
                Document? document = null;
                Selection? selection = null;
                Window? window = null;
                try
                {
                    document = application.ActiveDocument;
                    selection = application.Selection;
                    if (document.ReadOnly)
                        throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
                    var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
                    if (targets.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "当前文档没有带编号的 VisualTeX 行间公式。请先插入行间公式并勾选“添加公式编号”。",
                            "VisualTeX",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return string.Empty;
                    }

                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        var requestedIndex = 0;
                        _ = int.TryParse(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX"),
                            out requestedIndex);
                        requestedIndex = Math.Max(0, Math.Min(targets.Count - 1, requestedIndex));
                        var target = targets[requestedIndex];
                        WordEquationNumbering.InsertEquationReference(
                            document,
                            selection,
                            target,
                            EquationReferenceStyle.Parenthesized);
                        return target.NumberText;
                    }

                    using var dialog = new EquationReferenceDialog(targets);
                    System.Windows.Forms.DialogResult result;
                    try
                    {
                        window = application.ActiveWindow;
                        result = dialog.ShowDialog(new NativeWindowOwner(new IntPtr(window.Hwnd)));
                    }
                    catch
                    {
                        result = dialog.ShowDialog();
                    }
                    if (result != System.Windows.Forms.DialogResult.OK
                        || dialog.SelectedTarget is null)
                        return string.Empty;
                    WordEquationNumbering.InsertEquationReference(
                        document,
                        selection,
                        dialog.SelectedTarget,
                        dialog.SelectedStyle);
                    return dialog.SelectedTarget.NumberText;
                }
                finally
                {
                    ReleaseComObject(window);
                    ReleaseComObject(selection);
                    ReleaseComObject(document);
                }
            }).ConfigureAwait(false);
            if (!string.IsNullOrWhiteSpace(inserted))
                SetStatus($"已插入公式 ({inserted}) 的交叉引用；更新编号时引用会同步刷新。");
        }
        catch (Exception error)
        {
            SetStatus($"插入公式引用失败：{error.Message}");
        }
    }

    private async Task ExportSelectedAsPictureAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            await dispatcher.InvokeAsync(service.ExportSelectedOleAsPicture)
                .ConfigureAwait(false);
            SetStatus("Word OLE 公式已导出为跨平台图片。");
        }
        catch (Exception error)
        {
            SetStatus($"导出 Word OLE 公式失败：{error.Message}");
        }
    }

    private async Task DeleteSelectedAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            await dispatcher.InvokeAsync(service.DeleteSelectedFormula).ConfigureAwait(false);
            SetStatus("Word 公式已删除。");
        }
        catch (Exception error)
        {
            SetStatus($"删除 Word 公式失败：{error.Message}");
        }
    }

    private void SetStatus(string message)
    {
        var dispatcher = _dispatcher;
        var application = _application;
        if (dispatcher is null || application is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            try { application.StatusBar = message; } catch { }
            return true;
        });
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    private void Dispose()
    {
        _lifetime?.Cancel();
        if (_application is not null)
        {
            try { _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick; } catch { }
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
        }
        try { _doubleClickHook?.Dispose(); } catch { }
        _doubleClickHook = null;
        ClearNativeOleTarget();
        _sessionClient?.Dispose();
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
        _lifetime = null;
        Volatile.Write(ref _activeSessionId, null);
        _ribbonUi = null;
        if (_comAddIn is not null)
        {
            try { _comAddIn.Object = null; } catch { }
            ReleaseComObject(_comAddIn);
            _comAddIn = null;
        }
        _application = null;
    }
}
