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
    VisualTeX_StabilizeImageEquationNumberSelection Sel.Range
    VisualTeX_SynchronizeSelectedImageFormulaSize Sel
    VTInvalidateWordImageFontSizeControl
End Sub

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    On Error GoTo IgnoreDoubleClick
    If Sel Is Nothing Then Exit Sub
    Cancel = VTHandleWordBeforeDoubleClick(Sel)
    Exit Sub

IgnoreDoubleClick:
    ' Ordinary Word double-click behavior must remain untouched when the
    ' clicked object is not a valid VisualTeX formula.
End Sub
