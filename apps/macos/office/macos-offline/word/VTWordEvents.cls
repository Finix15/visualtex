VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VTWordEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As Word.Selection)
    On Error Resume Next
    If Sel Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    VTCleanupEmptyVisualTeXImageMacroButtons Sel.Range
    VisualTeX_StabilizeImageEquationNumberSelection Sel.Range
    VisualTeX_SynchronizeSelectedImageFormulaSize Sel
    VTInvalidateWordImageFontSizeControl
End Sub

Private Sub App_DocumentOpen(ByVal Doc As Word.Document)
    On Error Resume Next
    VTEnsureImageMacroMigrationScheduled
    On Error GoTo 0
End Sub

Private Sub App_NewDocument(ByVal Doc As Word.Document)
    On Error Resume Next
    VTEnsureImageMacroMigrationScheduled
    On Error GoTo 0
End Sub

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    Dim selectedShape As InlineShape
    Dim editErrorNumber As Long
    Dim editErrorDescription As String

    ' Keep the direct-picture hot path free of tracing, metadata inflation and
    ' document searches. Word can invoke FormatPicture after this event, so a
    ' selected shape carrying either the compact reference or durable VisualTeX
    ' metadata prefix must be cancelled before dispatching the edit.
    On Error GoTo IgnoreDoubleClick
    If Sel Is Nothing Then Exit Sub
    If VTTryGetDirectVisualTeXImageCandidate(Sel, selectedShape) Then
        Cancel = True
        On Error GoTo VisualTeXEditFailed
        If Not VTDispatchDirectVisualTeXImageEdit(selectedShape) Then
            ' A forged prefix or damaged metadata is not a VisualTeX formula.
            ' Release the event so Word can run its native picture command.
            Cancel = False
        End If
        Exit Sub
    End If
    If Sel.InlineShapes.Count > 0 Then Exit Sub

    VTBeginWordDoubleClickTraceContext
    On Error GoTo IgnoreTracedDoubleClick
    VTTraceWordDoubleClick "window-before-double-click-enter", Sel, ""
    Cancel = VTHandleWordBeforeDoubleClick(Sel)
    VTTraceWordDoubleClick _
        "window-before-double-click-result", Sel, _
        "cancel=" & CStr(Cancel)
    GoTo DoubleClickFinished

VisualTeXEditFailed:
    editErrorNumber = Err.Number
    editErrorDescription = Err.Description
    Cancel = True
    On Error Resume Next
    VTShowError "Word image formula edit", _
        editErrorNumber, editErrorDescription
    On Error GoTo 0
    Exit Sub

IgnoreDoubleClick:
    ' Ordinary Word double-click behavior remains untouched before the trace
    ' context is entered, including ordinary inline and floating pictures.
    Exit Sub

IgnoreTracedDoubleClick:
    VTTraceWordDoubleClick _
        "window-before-double-click-error", Sel, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    ' Ordinary Word double-click behavior must remain untouched when the
    ' clicked object is not a valid VisualTeX formula.
DoubleClickFinished:
    VTEndWordDoubleClickTraceContext
End Sub
