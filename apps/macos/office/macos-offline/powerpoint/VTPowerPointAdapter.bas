Attribute VB_Name = "VTPowerPointAdapter"
Option Explicit

Private Const VT_POWERPOINT_HOST As String = "powerpoint"
Private Const VT_POWERPOINT_STATUS_FILE As String = "/OfficePluginStatus/powerpoint.json"
Private Const VT_POWERPOINT_SOURCE_REVISION As String = _
    "powerpoint-office-performance-20260801-r4"
Private Const VT_SHAPE_PREFIX As String = "VisualTeX_"
Private Const VT_DEFAULT_PLACEHOLDER_WIDTH As Single = 180!
Private Const VT_DEFAULT_PLACEHOLDER_HEIGHT As Single = 42!
Private Const VT_POWERPOINT_REFERENCE_FONT_SIZE_PT As Double = 14#
Private Const VT_POWERPOINT_DEFAULT_FONT_SIZE_PT As Double = 18#
Private Const VT_POWERPOINT_MIN_FONT_SIZE_PT As Double = 1#
Private Const VT_POWERPOINT_MAX_FONT_SIZE_PT As Double = 512#
Private Const VT_POWERPOINT_FONT_CONTROL_ID As String = _
    "VisualTeX.Mac.PowerPoint.FormulaFontSize"
Private Const VT_POWERPOINT_FONT_TAG As String = "VisualTeXFontSizePt"
Private Const VT_POWERPOINT_REFERENCE_WIDTH_TAG As String = _
    "VisualTeXReferenceWidthPt"
Private Const VT_POWERPOINT_REFERENCE_HEIGHT_TAG As String = _
    "VisualTeXReferenceHeightPt"
Private VT_POWERPOINT_EVENT_SINK As VTPowerPointEvents
Private VT_POWERPOINT_RIBBON As IRibbonUI

Public Sub Auto_Open()
    On Error Resume Next
    VTInitializePowerPointEvents
    VTPrewarmApplication VT_POWERPOINT_HOST
    VTWritePowerPointHealth
    On Error GoTo 0
End Sub

Public Sub VTInitializePowerPointEvents()
    Set VT_POWERPOINT_EVENT_SINK = New VTPowerPointEvents
    Set VT_POWERPOINT_EVENT_SINK.App = PowerPoint.Application
End Sub

Public Sub VTPowerPointRibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set VT_POWERPOINT_RIBBON = ribbon
    VTInitializePowerPointEvents
    VTInvalidatePowerPointFormulaFontSizeControl
    VTWritePowerPointHealth
End Sub

Public Sub VisualTeX_NewFormula()
    On Error GoTo Failed

    Dim sessionId As String
    Dim formulaId As String
    Dim pendingMarker As String
    Dim currentSlide As Slide
    Dim placeholder As Shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim requestJson As String
    Dim powerPointJson As String
    Dim requestedFontSizePt As Double
    Dim failureStage As String

    failureStage = "validate presentation"
    VTRequireWritablePowerPointPresentation
    Set currentSlide = ActiveWindow.View.Slide

    failureStage = "resolve formula font size"
    requestedFontSizePt = VTPreferredPowerPointFormulaFontSize()

    failureStage = "create identifiers"
    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight

    failureStage = "create placeholder shape"
    Set placeholder = currentSlide.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        (slideWidth - VT_DEFAULT_PLACEHOLDER_WIDTH) / 2!, _
        (slideHeight - VT_DEFAULT_PLACEHOLDER_HEIGHT) / 2!, _
        VT_DEFAULT_PLACEHOLDER_WIDTH, _
        VT_DEFAULT_PLACEHOLDER_HEIGHT)

    failureStage = "format placeholder shape"
    placeholder.Name = VT_SHAPE_PREFIX & formulaId
    placeholder.Fill.Visible = msoFalse
    placeholder.Line.Visible = msoTrue
    placeholder.Line.Weight = 1!
    placeholder.Line.ForeColor.RGB = RGB(128, 128, 128)
    placeholder.TextFrame.TextRange.Text = "VisualTeX"
    placeholder.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

    failureStage = "attach placeholder metadata"
    VTSetShapeTag placeholder, "VisualTeXFormulaId", formulaId
    VTSetShapeTag placeholder, "VisualTeXSessionId", sessionId
    VTSetShapeTag placeholder, "VisualTeXPending", "1"
    VTSetShapeTag placeholder, VT_POWERPOINT_FONT_TAG, _
        VTJsonNumber(requestedFontSizePt)
    placeholder.AlternativeText = pendingMarker

    failureStage = "build request"
    powerPointJson = VTPowerPointGeometryJson( _
        currentSlide, placeholder, requestedFontSizePt)
    requestJson = VTRequestJson( _
        sessionId, _
        VT_POWERPOINT_HOST, _
        "create", _
        formulaId, _
        "block", _
        False, _
        VTPresentationIdentity(), _
        placeholder.Name, _
        "", _
        pendingMarker, _
        powerPointJson)

    failureStage = "write request and open VisualTeX editor"
    Call VTWriteAndLaunchSession( _
        VT_POWERPOINT_HOST, sessionId, requestJson)
    Exit Sub

Failed:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    If Len(failureStage) > 0 Then
        errorDescription = "Stage: " & failureStage & ". " & errorDescription
    End If
    On Error Resume Next
    If Not placeholder Is Nothing Then placeholder.Delete
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError "PowerPoint formula creation", errorNumber, errorDescription
End Sub

Public Sub VisualTeX_EditSelected()
    On Error GoTo Failed

    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape VTSelectedSingleShape()
    Exit Sub

Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_DoubleClickEditSelected()
    ' Invoked by the native macOS double-click monitor. Let VBA errors return
    ' to the monitor instead of displaying a modal message for ordinary shapes.
    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape VTSelectedSingleShape()
End Sub

Public Sub VisualTeX_EditShape(ByVal selectedShape As Shape)
    On Error GoTo Failed
    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape selectedShape
    Exit Sub
Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
End Sub

Public Function VTIsVisualTeXPowerPointShape(ByVal selectedShape As Shape) As Boolean
    Dim formulaId As String
    Dim encodedMetadata As String
    Dim parsedFormulaId As String
    Dim displayMode As String
    Dim numbered As Boolean

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then Exit Function
    formulaId = VTShapeFormulaId(selectedShape)
    encodedMetadata = VTShapeMetadata(selectedShape)
    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function
    If Not VTTryParseFormulaReference(selectedShape.Title, parsedFormulaId, displayMode, numbered) Then Exit Function
    If parsedFormulaId <> formulaId Then Exit Function
    VTIsVisualTeXPowerPointShape = True
    Exit Function
InvalidShape:
    VTIsVisualTeXPowerPointShape = False
End Function

Private Sub VTPowerPointEditShape(ByVal selectedShape As Shape)
    Dim formulaId As String
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim sessionId As String
    Dim requestJson As String
    Dim powerPointJson As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    If selectedShape Is Nothing Then
        Err.Raise vbObjectError + 7500, "VisualTeX", "Select one VisualTeX formula shape."
    End If
    formulaId = VTShapeFormulaId(selectedShape)
    encodedMetadata = VTShapeMetadata(selectedShape)
    formulaReference = selectedShape.Title
    VTValidateEditEnvelope encodedMetadata, formulaReference, formulaId, displayMode, numbered
    If Len(formulaId) = 0 Then formulaId = VTShapeFormulaId(selectedShape)
    If Len(formulaId) = 0 Then
        Err.Raise vbObjectError + 7500, "VisualTeX", "The selected PowerPoint shape has no VisualTeX formula id."
    End If
    VTEnsurePowerPointFormulaScaleState _
        selectedShape, fontSizePt, referenceWidthPt, referenceHeightPt

    sessionId = VTNewUuidV4()
    powerPointJson = VTPowerPointGeometryJson( _
        ActiveWindow.View.Slide, selectedShape, fontSizePt, _
        referenceWidthPt, referenceHeightPt)
    requestJson = VTRequestJson( _
        sessionId, _
        VT_POWERPOINT_HOST, _
        "edit", _
        formulaId, _
        "block", _
        False, _
        VTPresentationIdentity(), _
        selectedShape.Name, _
        encodedMetadata, _
        "", _
        powerPointJson)
    Call VTWriteAndLaunchSession( _
        VT_POWERPOINT_HOST, sessionId, requestJson)
End Sub

Public Sub VisualTeX_DeleteSelected()
    On Error GoTo Failed
    Dim selectedShape As Shape
    VTRequireWritablePowerPointPresentation
    Set selectedShape = VTSelectedSingleShape()
    If Len(VTShapeFormulaId(selectedShape)) = 0 Then
        Err.Raise vbObjectError + 7501, "VisualTeX", "Select one VisualTeX formula shape."
    End If
    selectedShape.Delete
    Exit Sub
Failed:
    VTShowError "PowerPoint delete", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenApplication()
    On Error GoTo Failed
    VTOpenApplication VT_POWERPOINT_HOST
    Exit Sub
Failed:
    VTShowError "application launch", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_ApplyPendingResult()
    On Error GoTo Failed

    Dim sessionId As String
    Dim dispatch As Object
    Dim actionName As String
    Dim hostName As String

    sessionId = VTReadActiveSessionId(VT_POWERPOINT_HOST)
    Set dispatch = VTReadDispatch(sessionId)
    actionName = CStr(dispatch("action"))
    hostName = CStr(dispatch("host"))
    If hostName <> VT_POWERPOINT_HOST Then
        Err.Raise vbObjectError + 7502, "VisualTeX", "The active VisualTeX dispatch is not for PowerPoint."
    End If

    Select Case actionName
        Case "commit": VTFinalizePowerPointDispatch sessionId, dispatch
        Case "cancel": VTCancelPowerPointDispatch sessionId, dispatch
        Case Else
            Err.Raise vbObjectError + 7503, "VisualTeX", "The VisualTeX PowerPoint dispatch action is invalid."
    End Select
    Exit Sub

Failed:
    Err.Raise Err.Number, "VisualTeX PowerPoint callback", Err.Description
End Sub

Private Sub VTFinalizePowerPointDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim formulaId As String
    Dim metadata As String
    Dim shapeName As String
    Dim sourceShapeName As String
    Dim imagePath As String
    Dim fallbackImagePath As String
    Dim vectorInsertErrorNumber As Long
    Dim vectorInsertErrorDescription As String
    Dim expectedPresentation As String
    Dim slideIndex As Long
    Dim slideId As Long
    Dim targetZOrder As Long
    Dim targetLeft As Double
    Dim targetTop As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim targetRotation As Double
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim currentSlide As Slide
    Dim committed As Shape
    Dim original As Shape
    Dim candidate As Shape
    Dim originalTemporaryName As String
    Dim candidateTemporaryName As String
    Dim originalRenamed As Boolean
    Dim formulaReference As String

    VTRequireWritablePowerPointPresentation
    VTRequireDispatchValue dispatch, "formulaId"
    VTRequireDispatchValue dispatch, "metadata"
    VTRequireDispatchValue dispatch, "shapeName"
    VTRequireDispatchValue dispatch, "sourceShapeName"
    VTRequireDispatchValue dispatch, "imagePath"
    VTRequireDispatchValue dispatch, "presentationIdentity"
    VTRequireDispatchValue dispatch, "slideIndex"
    VTRequireDispatchValue dispatch, "slideId"
    VTRequireDispatchValue dispatch, "targetLeft"
    VTRequireDispatchValue dispatch, "targetTop"
    VTRequireDispatchValue dispatch, "targetWidth"
    VTRequireDispatchValue dispatch, "targetHeight"
    VTRequireDispatchValue dispatch, "fontSizePt"
    VTRequireDispatchValue dispatch, "referenceWidthPt"
    VTRequireDispatchValue dispatch, "referenceHeightPt"
    VTRequireDispatchValue dispatch, "rotation"
    VTRequireDispatchValue dispatch, "zOrder"

    formulaId = CStr(dispatch("formulaId"))
    metadata = CStr(dispatch("metadata"))
    shapeName = CStr(dispatch("shapeName"))
    sourceShapeName = CStr(dispatch("sourceShapeName"))
    imagePath = CStr(dispatch("imagePath"))
    fallbackImagePath = VTDispatchOptionalPpt(dispatch, "fallbackImagePath")
    expectedPresentation = CStr(dispatch("presentationIdentity"))
    slideIndex = CLng(dispatch("slideIndex"))
    slideId = CLng(dispatch("slideId"))
    targetLeft = VTDispatchDoublePpt(dispatch, "targetLeft")
    targetTop = VTDispatchDoublePpt(dispatch, "targetTop")
    targetWidth = VTDispatchPositiveDoublePpt(dispatch, "targetWidth")
    targetHeight = VTDispatchPositiveDoublePpt(dispatch, "targetHeight")
    fontSizePt = VTDispatchPositiveDoublePpt(dispatch, "fontSizePt")
    referenceWidthPt = VTDispatchPositiveDoublePpt( _
        dispatch, "referenceWidthPt")
    referenceHeightPt = VTDispatchPositiveDoublePpt( _
        dispatch, "referenceHeightPt")
    targetRotation = VTDispatchDoublePpt(dispatch, "rotation")
    targetZOrder = CLng(dispatch("zOrder"))

    If Not VTIsCanonicalUuid(formulaId) Or _
       shapeName <> VT_SHAPE_PREFIX & formulaId Or _
       Not VTIsEncodedMetadata(metadata) Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "VisualTeX PowerPoint result metadata is invalid."
    End If
    formulaReference = VTFormulaReference(formulaId, "block", False)
    If expectedPresentation <> VTPresentationIdentity() Then
        Err.Raise vbObjectError + 7515, "VisualTeX", "The active PowerPoint presentation changed while VisualTeX was open."
    End If
    If slideIndex <= 0 Or slideIndex > ActivePresentation.Slides.Count Or slideId <= 0 Or targetZOrder <= 0 Then
        Err.Raise vbObjectError + 7516, "VisualTeX", "VisualTeX PowerPoint target reference is invalid."
    End If
    VTValidateAbsoluteVisualTeXPath imagePath
    If Not VTPathFileExists(imagePath) Then
        Err.Raise vbObjectError + 7517, "VisualTeX", "VisualTeX PowerPoint SVG result is missing."
    End If
    If Len(fallbackImagePath) > 0 Then
        VTValidateAbsoluteVisualTeXPath fallbackImagePath
        If Not VTPathFileExists(fallbackImagePath) Then fallbackImagePath = ""
    End If

    Set currentSlide = ActivePresentation.Slides(slideIndex)
    If currentSlide.SlideID <> slideId Then
        Err.Raise vbObjectError + 7518, "VisualTeX", "The original PowerPoint slide no longer exists."
    End If
    On Error Resume Next
    Set committed = currentSlide.Shapes(shapeName)
    On Error GoTo TransactionFailed
    If Not committed Is Nothing Then
        If VTIsCommittedPowerPointShape( _
            committed, shapeName, formulaReference, metadata, formulaId, sessionId, _
            targetLeft, targetTop, targetWidth, targetHeight, targetRotation, _
            targetZOrder, fontSizePt, referenceWidthPt, referenceHeightPt) Then
            Exit Sub
        End If
    End If
    Set committed = Nothing

    On Error Resume Next
    Set original = currentSlide.Shapes(sourceShapeName)
    On Error GoTo TransactionFailed
    If original Is Nothing Then
        Err.Raise vbObjectError + 7519, "VisualTeX", "The original VisualTeX PowerPoint shape no longer exists."
    End If

    candidateTemporaryName = "VisualTeXPendingResult_" & Replace$(Left$(sessionId, 13), "-", "")
    originalTemporaryName = "VisualTeXOriginal_" & Replace$(Left$(sessionId, 13), "-", "")
    ' Modern PowerPoint for Mac preserves an imported SVG as vector artwork.
    ' Prefer SVG so formulas remain sharp at arbitrary zoom. Keep the PNG only
    ' as a compatibility fallback for Office builds that reject SVG AddPicture.
    On Error Resume Next
    Set candidate = currentSlide.Shapes.AddPicture( _
        FileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=CSng(targetLeft), _
        Top:=CSng(targetTop), _
        Width:=CSng(targetWidth), _
        Height:=CSng(targetHeight))
    vectorInsertErrorNumber = Err.Number
    vectorInsertErrorDescription = Err.Description
    Err.Clear
    On Error GoTo TransactionFailed
    If candidate Is Nothing Then
        If Len(fallbackImagePath) = 0 Then
            Err.Raise vbObjectError + 7521, "VisualTeX", _
                "PowerPoint could not insert the VisualTeX SVG: " & _
                CStr(vectorInsertErrorNumber) & " " & vectorInsertErrorDescription
        End If
        Set candidate = currentSlide.Shapes.AddPicture( _
            FileName:=fallbackImagePath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=CSng(targetLeft), _
            Top:=CSng(targetTop), _
            Width:=CSng(targetWidth), _
            Height:=CSng(targetHeight))
    End If
    candidate.Name = candidateTemporaryName
    candidate.LockAspectRatio = msoFalse
    candidate.Left = CSng(targetLeft)
    candidate.Top = CSng(targetTop)
    candidate.Width = CSng(targetWidth)
    candidate.Height = CSng(targetHeight)
    candidate.LockAspectRatio = msoTrue
    candidate.Rotation = CSng(targetRotation)
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    VTSetShapeTag candidate, "VisualTeXFormulaId", formulaId
    VTSetShapeTag candidate, "VisualTeXSessionId", sessionId
    VTSetShapeTag candidate, "VisualTeXPending", "0"
    VTSetPowerPointFormulaScaleState _
        candidate, fontSizePt, referenceWidthPt, referenceHeightPt
    On Error Resume Next
    VTSetShapeTag candidate, "VisualTeXMetadata", metadata
    Err.Clear
    On Error GoTo TransactionFailed

    ' The original still occupies targetZOrder. Put the candidate immediately
    ' above it; deleting the original as the final mutation shifts the candidate
    ' into the exact original z-order without any fallible operation afterwards.
    VTRestoreZOrder candidate, targetZOrder + 1
    original.Name = originalTemporaryName
    originalRenamed = True
    candidate.Name = shapeName

    If candidate.Name <> shapeName Or _
       Abs(candidate.Left - targetLeft) > 0.1 Or Abs(candidate.Top - targetTop) > 0.1 Or _
       Abs(candidate.Width - targetWidth) > 0.1 Or Abs(candidate.Height - targetHeight) > 0.1 Or _
       Abs(candidate.Rotation - targetRotation) > 0.1 Or candidate.ZOrderPosition <> targetZOrder + 1 Or _
       candidate.AlternativeText <> metadata Or candidate.Title <> formulaReference Or _
       candidate.Tags("VisualTeXFormulaId") <> formulaId Or _
       candidate.Tags("VisualTeXSessionId") <> sessionId Or _
       candidate.Tags("VisualTeXPending") <> "0" Or _
       candidate.Tags(VT_POWERPOINT_FONT_TAG) <> VTJsonNumber(fontSizePt) Or _
       candidate.Tags(VT_POWERPOINT_REFERENCE_WIDTH_TAG) <> _
            VTJsonNumber(referenceWidthPt) Or _
       candidate.Tags(VT_POWERPOINT_REFERENCE_HEIGHT_TAG) <> _
            VTJsonNumber(referenceHeightPt) Then
        Err.Raise vbObjectError + 7520, "VisualTeX", "PowerPoint did not persist the VisualTeX formula properties."
    End If

    original.Delete
    Exit Sub

TransactionFailed:
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    On Error Resume Next
    If Not candidate Is Nothing Then candidate.Delete
    If originalRenamed And Not original Is Nothing Then original.Name = sourceShapeName
    On Error GoTo 0
    Err.Raise transactionErrorNumber, "VisualTeX PowerPoint transaction", transactionErrorDescription
End Sub

Private Function VTIsCommittedPowerPointShape( _
    ByVal target As Shape, _
    ByVal expectedName As String, _
    ByVal formulaReference As String, _
    ByVal metadata As String, _
    ByVal formulaId As String, _
    ByVal sessionId As String, _
    ByVal targetLeft As Double, _
    ByVal targetTop As Double, _
    ByVal targetWidth As Double, _
    ByVal targetHeight As Double, _
    ByVal targetRotation As Double, _
    ByVal targetZOrder As Long, _
    ByVal fontSizePt As Double, _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double) As Boolean

    On Error GoTo NotCommitted
    VTIsCommittedPowerPointShape = _
        target.Name = expectedName And _
        Abs(target.Left - targetLeft) <= 0.1 And _
        Abs(target.Top - targetTop) <= 0.1 And _
        Abs(target.Width - targetWidth) <= 0.1 And _
        Abs(target.Height - targetHeight) <= 0.1 And _
        Abs(target.Rotation - targetRotation) <= 0.1 And _
        target.ZOrderPosition = targetZOrder And _
        target.AlternativeText = metadata And _
        target.Title = formulaReference And _
        target.Tags("VisualTeXFormulaId") = formulaId And _
        target.Tags("VisualTeXSessionId") = sessionId And _
        target.Tags("VisualTeXPending") = "0" And _
        target.Tags(VT_POWERPOINT_FONT_TAG) = VTJsonNumber(fontSizePt) And _
        target.Tags(VT_POWERPOINT_REFERENCE_WIDTH_TAG) = _
            VTJsonNumber(referenceWidthPt) And _
        target.Tags(VT_POWERPOINT_REFERENCE_HEIGHT_TAG) = _
            VTJsonNumber(referenceHeightPt)
    Exit Function

NotCommitted:
    Err.Clear
    VTIsCommittedPowerPointShape = False
End Function

Private Sub VTCancelPowerPointDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim currentSlide As Slide
    Dim shapeItem As Shape

    pendingMarker = VTDispatchOptionalPpt(dispatch, "pendingMarker")
    If Len(pendingMarker) = 0 Or Presentations.Count = 0 Then Exit Sub
    On Error Resume Next
    For Each currentSlide In ActivePresentation.Slides
        For Each shapeItem In currentSlide.Shapes
            If shapeItem.AlternativeText = pendingMarker And _
               shapeItem.Tags("VisualTeXSessionId") = sessionId And _
               shapeItem.Tags("VisualTeXPending") = "1" Then
                shapeItem.Delete
                Exit Sub
            End If
        Next shapeItem
    Next currentSlide
    On Error GoTo 0
End Sub

Private Function VTSelectedSingleShape() As Shape
    If ActiveWindow Is Nothing Then
        Err.Raise vbObjectError + 7505, "VisualTeX", "PowerPoint has no active window."
    End If
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Err.Raise vbObjectError + 7506, "VisualTeX", "Select exactly one VisualTeX formula shape."
    End If
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        Err.Raise vbObjectError + 7507, "VisualTeX", "Select exactly one VisualTeX formula shape."
    End If
    Set VTSelectedSingleShape = ActiveWindow.Selection.ShapeRange(1)
End Function

Private Function VTShapeFormulaId(ByVal target As Shape) As String
    Dim value As String
    On Error Resume Next
    value = target.Tags("VisualTeXFormulaId")
    On Error GoTo 0
    If Not VTIsCanonicalUuid(value) And Left$(target.Name, Len(VT_SHAPE_PREFIX)) = VT_SHAPE_PREFIX Then
        value = Mid$(target.Name, Len(VT_SHAPE_PREFIX) + 1)
    End If
    If VTIsCanonicalUuid(value) Then VTShapeFormulaId = value
End Function

Private Function VTShapeMetadata(ByVal target As Shape) As String
    Dim value As String
    On Error Resume Next
    value = target.Tags("VisualTeXMetadata")
    On Error GoTo 0
    If Not VTIsEncodedMetadata(value) Then value = target.AlternativeText
    If VTIsEncodedMetadata(value) Then VTShapeMetadata = value
End Function

Private Function VTFindUniqueFormulaShape(ByVal shapeName As String) As Shape
    Dim slideItem As Slide
    Dim candidate As Shape
    Dim match As Shape
    Dim count As Long

    If Len(shapeName) = 0 Or Len(shapeName) > 128 Or InStr(shapeName, vbCr) > 0 Or InStr(shapeName, vbLf) > 0 Then
        Err.Raise vbObjectError + 7508, "VisualTeX", "VisualTeX PowerPoint shape name is invalid."
    End If
    For Each slideItem In ActivePresentation.Slides
        Set candidate = Nothing
        On Error Resume Next
        Set candidate = slideItem.Shapes(shapeName)
        Err.Clear
        On Error GoTo 0
        If Not candidate Is Nothing Then
            count = count + 1
            Set match = candidate
        End If
    Next slideItem
    If count <> 1 Or match Is Nothing Then
        Err.Raise vbObjectError + 7509, "VisualTeX", "PowerPoint must contain exactly one matching VisualTeX formula shape."
    End If
    Set VTFindUniqueFormulaShape = match
End Function

Private Function VTValidPowerPointFormulaFontSize( _
    ByVal value As Double) As Boolean

    VTValidPowerPointFormulaFontSize = _
        value >= VT_POWERPOINT_MIN_FONT_SIZE_PT And _
        value <= VT_POWERPOINT_MAX_FONT_SIZE_PT
End Function

Private Function VTTryReadPowerPointShapeTagDouble( _
    ByVal target As Shape, _
    ByVal tagName As String, _
    ByRef value As Double) As Boolean

    Dim storedValue As String

    value = 0#
    If target Is Nothing Or Len(tagName) = 0 Then Exit Function
    On Error GoTo InvalidValue
    storedValue = target.Tags(tagName)
    If Len(storedValue) = 0 Then Exit Function
    value = VTParseInvariantDouble(storedValue)
    If Abs(value) > 10000000# Then Exit Function
    VTTryReadPowerPointShapeTagDouble = True
    Exit Function

InvalidValue:
    value = 0#
    Err.Clear
End Function

Private Sub VTSetPowerPointFormulaScaleState( _
    ByVal target As Shape, _
    ByVal fontSizePt As Double, _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double)

    If target Is Nothing Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceWidthPt > 10000# Or referenceHeightPt > 10000# Then
        Err.Raise vbObjectError + 7523, "VisualTeX", _
            "The PowerPoint formula point-size state is invalid."
    End If
    VTSetShapeTag target, VT_POWERPOINT_FONT_TAG, VTJsonNumber(fontSizePt)
    VTSetShapeTag target, VT_POWERPOINT_REFERENCE_WIDTH_TAG, _
        VTJsonNumber(referenceWidthPt)
    VTSetShapeTag target, VT_POWERPOINT_REFERENCE_HEIGHT_TAG, _
        VTJsonNumber(referenceHeightPt)
End Sub

Private Sub VTEnsurePowerPointFormulaScaleState( _
    ByVal target As Shape, _
    ByRef fontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double)

    Dim hasFontSize As Boolean
    Dim hasReferenceWidth As Boolean
    Dim hasReferenceHeight As Boolean
    Dim observedFontSizePt As Double
    Dim scaleStateNeedsWrite As Boolean

    If target Is Nothing Or Not VTIsVisualTeXPowerPointShape(target) Then
        Err.Raise vbObjectError + 7524, "VisualTeX", _
            "Select one VisualTeX PowerPoint SVG formula."
    End If
    hasFontSize = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_FONT_TAG, fontSizePt)
    hasReferenceWidth = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_REFERENCE_WIDTH_TAG, referenceWidthPt)
    hasReferenceHeight = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_REFERENCE_HEIGHT_TAG, referenceHeightPt)

    If Not hasFontSize Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Then
        fontSizePt = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
        scaleStateNeedsWrite = True
    End If
    If Not hasReferenceWidth Or Not hasReferenceHeight Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Then
        referenceWidthPt = target.Width * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT / fontSizePt
        referenceHeightPt = target.Height * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT / fontSizePt
        scaleStateNeedsWrite = True
    Else
        observedFontSizePt = target.Height / referenceHeightPt * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
        If VTValidPowerPointFormulaFontSize(observedFontSizePt) Then
            If Abs(observedFontSizePt - fontSizePt) > 0.05 Then
                fontSizePt = observedFontSizePt
                scaleStateNeedsWrite = True
            End If
        End If
    End If
    If scaleStateNeedsWrite Then
        VTSetPowerPointFormulaScaleState _
            target, fontSizePt, referenceWidthPt, referenceHeightPt
    End If
End Sub

Private Function VTPreferredPowerPointFormulaFontSize() As Double
    Dim selectedShape As Shape
    Dim selectedSize As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    selectedSize = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
    On Error GoTo UseDefault
    If ActiveWindow Is Nothing Then GoTo UseDefault
    If ActiveWindow.Selection.Type = ppSelectionText Then
        selectedSize = ActiveWindow.Selection.TextRange.Font.Size
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes And _
           ActiveWindow.Selection.ShapeRange.Count = 1 Then
        Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
        If VTIsVisualTeXPowerPointShape(selectedShape) Then
            VTEnsurePowerPointFormulaScaleState _
                selectedShape, selectedSize, referenceWidthPt, referenceHeightPt
        ElseIf selectedShape.HasTextFrame = msoTrue Then
            If selectedShape.TextFrame.HasText = msoTrue Then
                selectedSize = selectedShape.TextFrame.TextRange.Font.Size
            End If
        End If
    End If

UseDefault:
    If Not VTValidPowerPointFormulaFontSize(selectedSize) Then
        selectedSize = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
    End If
    VTPreferredPowerPointFormulaFontSize = selectedSize
End Function

Private Sub VTApplyPowerPointFormulaFontSize( _
    ByVal target As Shape, _
    ByVal requestedFontSizePt As Double)

    Dim storedFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim originalLeft As Single
    Dim originalTop As Single
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim resizeErrorNumber As Long
    Dim resizeErrorDescription As String

    If target Is Nothing Or _
       Not VTValidPowerPointFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7525, "VisualTeX", _
            "Enter a VisualTeX PowerPoint formula font size from 1 to 512 pt."
    End If
    VTEnsurePowerPointFormulaScaleState _
        target, storedFontSizePt, referenceWidthPt, referenceHeightPt
    targetWidth = referenceWidthPt * requestedFontSizePt / _
        VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
    targetHeight = referenceHeightPt * requestedFontSizePt / _
        VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
    If targetWidth <= 0# Or targetHeight <= 0# Or _
       targetWidth > 10000# Or targetHeight > 10000# Then
        Err.Raise vbObjectError + 7526, "VisualTeX", _
            "The requested PowerPoint formula font size produces unsupported dimensions."
    End If

    originalLeft = target.Left
    originalTop = target.Top
    originalWidth = target.Width
    originalHeight = target.Height
    centerX = CDbl(originalLeft) + CDbl(originalWidth) / 2#
    centerY = CDbl(originalTop) + CDbl(originalHeight) / 2#
    On Error GoTo ResizeFailed
    target.LockAspectRatio = msoFalse
    target.Width = CSng(targetWidth)
    target.Height = CSng(targetHeight)
    target.Left = CSng(centerX - targetWidth / 2#)
    target.Top = CSng(centerY - targetHeight / 2#)
    target.LockAspectRatio = msoTrue
    VTSetPowerPointFormulaScaleState _
        target, requestedFontSizePt, referenceWidthPt, referenceHeightPt
    If Abs(target.Width - targetWidth) > 0.1 Or _
       Abs(target.Height - targetHeight) > 0.1 Or _
       Abs((target.Left + target.Width / 2!) - centerX) > 0.1 Or _
       Abs((target.Top + target.Height / 2!) - centerY) > 0.1 Then
        Err.Raise vbObjectError + 7527, "VisualTeX", _
            "PowerPoint did not persist the requested formula point size."
    End If
    VTInvalidatePowerPointFormulaFontSizeControl
    Exit Sub

ResizeFailed:
    resizeErrorNumber = Err.Number
    resizeErrorDescription = Err.Description
    On Error Resume Next
    target.LockAspectRatio = msoFalse
    target.Left = originalLeft
    target.Top = originalTop
    target.Width = originalWidth
    target.Height = originalHeight
    target.LockAspectRatio = msoTrue
    VTSetPowerPointFormulaScaleState _
        target, storedFontSizePt, referenceWidthPt, referenceHeightPt
    On Error GoTo 0
    Err.Raise resizeErrorNumber, "VisualTeX PowerPoint formula font size", _
        resizeErrorDescription
End Sub

Public Sub VisualTeX_SynchronizeSelectedFormulaSize( _
    ByVal selected As PowerPoint.Selection)

    Dim selectedShape As Shape
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    On Error GoTo SynchronizeFinished
    If selected Is Nothing Then Exit Sub
    If selected.Type <> ppSelectionShapes Or _
       selected.ShapeRange.Count <> 1 Then GoTo SynchronizeFinished
    Set selectedShape = selected.ShapeRange(1)
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then GoTo SynchronizeFinished
    VTEnsurePowerPointFormulaScaleState _
        selectedShape, fontSizePt, referenceWidthPt, referenceHeightPt

SynchronizeFinished:
    VTInvalidatePowerPointFormulaFontSizeControl
End Sub

Public Sub VisualTeX_SetSelectedFormulaFontSize( _
    ByVal enteredText As String)

    Dim selectedShape As Shape
    Dim requestedFontSizePt As Double

    If Not VTTryParsePowerPointFormulaFontSize( _
       enteredText, requestedFontSizePt) Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "Enter a formula font size such as 18, 24, 28 or 32."
    End If
    Set selectedShape = VTSelectedSingleShape()
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one VisualTeX PowerPoint SVG formula first."
    End If
    VTApplyPowerPointFormulaFontSize selectedShape, requestedFontSizePt
End Sub

Private Function VTTryParsePowerPointFormulaFontSize( _
    ByVal enteredText As String, _
    ByRef fontSizePt As Double) As Boolean

    Dim normalized As String

    fontSizePt = 0#
    normalized = Trim$(enteredText)
    If Len(normalized) = 0 Or Len(normalized) > 32 Then Exit Function
    On Error GoTo InvalidSize
    fontSizePt = VTParseInvariantDouble(normalized)
    VTTryParsePowerPointFormulaFontSize = _
        VTValidPowerPointFormulaFontSize(fontSizePt)
    Exit Function

InvalidSize:
    fontSizePt = 0#
    Err.Clear
End Function

Public Sub VTInvalidatePowerPointFormulaFontSizeControl()
    On Error Resume Next
    If Not VT_POWERPOINT_RIBBON Is Nothing Then
        VT_POWERPOINT_RIBBON.InvalidateControl VT_POWERPOINT_FONT_CONTROL_ID
    End If
    On Error GoTo 0
End Sub

Private Function VTPowerPointSelectedFormulaFontState( _
    ByRef formulaCount As Long, _
    ByRef mixedSizes As Boolean, _
    ByRef fontSizePt As Double) As Boolean

    Dim candidate As Shape
    Dim candidateSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    On Error GoTo StateFailed
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function

    For Each candidate In ActiveWindow.Selection.ShapeRange
        If VTIsVisualTeXPowerPointShape(candidate) Then
            VTEnsurePowerPointFormulaScaleState _
                candidate, candidateSizePt, referenceWidthPt, _
                referenceHeightPt
            formulaCount = formulaCount + 1
            If formulaCount = 1 Then
                fontSizePt = candidateSizePt
            ElseIf Abs(candidateSizePt - fontSizePt) > 0.05 Then
                mixedSizes = True
            End If
        End If
    Next candidate

    VTPowerPointSelectedFormulaFontState = (formulaCount > 0)
    Exit Function

StateFailed:
    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    Err.Clear
End Function

Private Sub VTApplyPowerPointFontSizePresetToSelection( _
    ByVal requestedFontSizePt As Double)

    Dim candidate As Shape
    Dim appliedCount As Long

    If Not VTValidPowerPointFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "The selected SVG formula font-size preset is invalid."
    End If
    If ActiveWindow Is Nothing Or _
       ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one or more VisualTeX PowerPoint SVG formulas first."
    End If

    For Each candidate In ActiveWindow.Selection.ShapeRange
        If VTIsVisualTeXPowerPointShape(candidate) Then
            VTApplyPowerPointFormulaFontSize _
                candidate, requestedFontSizePt
            appliedCount = appliedCount + 1
        End If
    Next candidate
    If appliedCount = 0 Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one or more VisualTeX PowerPoint SVG formulas first."
    End If
    VTInvalidatePowerPointFormulaFontSizeControl
End Sub

Public Sub VTPowerPointRibbonGetFormulaFontSizeItemCount( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = VTFormulaFontPresetCount() + 1
End Sub

Public Sub VTPowerPointRibbonGetFormulaFontSizeItemLabel( _
    ByVal control As IRibbonControl, _
    ByVal itemIndex As Integer, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double

    On Error GoTo LabelFailed
    If itemIndex = 0 Then
        If Not VTPowerPointSelectedFormulaFontState( _
           formulaCount, mixedSizes, fontSizePt) Then
            returnedValue = _
                VTUnicodeText(35831, 36873, 25321) & " SVG " & _
                VTUnicodeText(20844, 24335)
        ElseIf mixedSizes Then
            returnedValue = _
                VTUnicodeText(24403, 21069) & ": " & _
                VTUnicodeText(28151, 21512, 23383, 21495)
        Else
            returnedValue = _
                VTUnicodeText(24403, 21069) & ": " & _
                VTFormulaFontSizeDisplayLabel(fontSizePt)
        End If
    Else
        returnedValue = VTFormulaFontPresetLabel(itemIndex - 1)
    End If
    Exit Sub

LabelFailed:
    returnedValue = VTUnicodeText(23383, 21495)
    Err.Clear
End Sub

Public Sub VTPowerPointRibbonGetFormulaFontSizeSelectedIndex( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double
    Dim presetIndex As Long

    returnedValue = 0
    On Error GoTo SelectedFinished
    If Not VTPowerPointSelectedFormulaFontState( _
       formulaCount, mixedSizes, fontSizePt) Then Exit Sub
    If mixedSizes Then Exit Sub
    presetIndex = VTFormulaFontPresetIndex(fontSizePt)
    If presetIndex >= 0 Then returnedValue = presetIndex + 1

SelectedFinished:
End Sub

Public Sub VTPowerPointRibbonApplyFormulaFontSizePreset( _
    ByVal control As IRibbonControl, _
    ByVal selectedId As String, _
    ByVal selectedIndex As Integer)

    On Error GoTo Failed
    If selectedIndex <= 0 Then Exit Sub
    VTApplyPowerPointFontSizePresetToSelection _
        VTFormulaFontPresetSize(selectedIndex - 1)
    Exit Sub

Failed:
    VTShowError "PowerPoint SVG formula font size", _
        Err.Number, Err.Description
End Sub

Private Function VTPowerPointGeometryJson( _
    ByVal currentSlide As Slide, _
    ByVal target As Shape, _
    Optional ByVal fontSizePt As Double = 0#, _
    Optional ByVal referenceWidthPt As Double = 0#, _
    Optional ByVal referenceHeightPt As Double = 0#) As String

    VTPowerPointGeometryJson = "{" & _
        """presentationIdentity"":" & VTJsonString(VTPresentationIdentity()) & "," & _
        """slideIndex"":" & CStr(currentSlide.SlideIndex) & "," & _
        """slideId"":" & CStr(currentSlide.SlideID) & "," & _
        """shapeName"":" & VTJsonString(target.Name) & "," & _
        """left"":" & VTJsonNumber(target.Left) & "," & _
        """top"":" & VTJsonNumber(target.Top) & "," & _
        """width"":" & VTJsonNumber(target.Width) & "," & _
        """height"":" & VTJsonNumber(target.Height) & "," & _
        """rotation"":" & VTJsonNumber(target.Rotation) & "," & _
        """zOrder"":" & CStr(target.ZOrderPosition) & "," & _
        """fontSizePt"":" & _
            IIf(fontSizePt > 0#, VTJsonNumber(fontSizePt), "null") & "," & _
        """referenceWidthPt"":" & _
            IIf(referenceWidthPt > 0#, VTJsonNumber(referenceWidthPt), "null") & "," & _
        """referenceHeightPt"":" & _
            IIf(referenceHeightPt > 0#, VTJsonNumber(referenceHeightPt), "null") & _
        "}"
End Function

Private Function VTPresentationIdentity() As String
    On Error Resume Next
    VTPresentationIdentity = ActivePresentation.FullName
    If Err.Number <> 0 Or Len(VTPresentationIdentity) = 0 Then
        Err.Clear
        VTPresentationIdentity = ActivePresentation.Name
    End If
    On Error GoTo 0
    VTPresentationIdentity = VTBoundedIdentity(VTPresentationIdentity)
End Function

Private Sub VTRequireWritablePowerPointPresentation()
    If Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7510, "VisualTeX", "Open a PowerPoint presentation first."
    End If
    If ActivePresentation.ReadOnly = msoTrue Then
        Err.Raise vbObjectError + 7511, "VisualTeX", "The active PowerPoint presentation is read-only."
    End If
    If ActiveWindow Is Nothing Then
        Err.Raise vbObjectError + 7512, "VisualTeX", "Switch PowerPoint to a normal editing view."
    End If
End Sub

Private Sub VTSetShapeTag(ByVal target As Shape, ByVal key As String, ByVal value As String)
    On Error Resume Next
    target.Tags.Delete key
    Err.Clear
    On Error GoTo Failed
    target.Tags.Add key, value
    Exit Sub
Failed:
    Err.Raise vbObjectError + 7513, "VisualTeX", "PowerPoint could not persist VisualTeX tag " & key & "."
End Sub

Private Sub VTRestoreZOrder(ByVal target As Shape, ByVal expectedPosition As Long)
    Dim guard As Long
    If expectedPosition <= 0 Then Exit Sub
    Do While target.ZOrderPosition > expectedPosition And guard < 4096
        target.ZOrder msoSendBackward
        guard = guard + 1
    Loop
    Do While target.ZOrderPosition < expectedPosition And guard < 8192
        target.ZOrder msoBringForward
        guard = guard + 1
    Loop
    If target.ZOrderPosition <> expectedPosition Then
        Err.Raise vbObjectError + 7514, "VisualTeX", "PowerPoint could not restore the formula z-order."
    End If
End Sub

Private Function VTDispatchOptionalPpt(ByVal dispatch As Object, ByVal key As String) As String
    If VTCollectionHasKey(dispatch, key) Then VTDispatchOptionalPpt = CStr(dispatch(key))
End Function

Private Function VTDispatchDoublePpt(ByVal dispatch As Object, ByVal key As String) As Double
    VTRequireDispatchValue dispatch, key
    VTDispatchDoublePpt = VTParseInvariantDouble(CStr(dispatch(key)))
    If Abs(VTDispatchDoublePpt) > 10000000# Then
        Err.Raise vbObjectError + 7521, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Private Function VTDispatchPositiveDoublePpt(ByVal dispatch As Object, ByVal key As String) As Double
    VTDispatchPositiveDoublePpt = VTDispatchDoublePpt(dispatch, key)
    If VTDispatchPositiveDoublePpt <= 0# Then
        Err.Raise vbObjectError + 7522, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Private Sub VTWritePowerPointHealth()
    Dim statusPath As String
    Dim payload As String
    statusPath = VTApplicationSupportRoot() & VT_POWERPOINT_STATUS_FILE
    payload = "{" & _
        """loaded"":true," & _
        """pluginVersion"":" & VTJsonString(VT_PLUGIN_VERSION) & "," & _
        """sourceRevision"":" & _
            VTJsonString(VT_POWERPOINT_SOURCE_REVISION) & "," & _
        """host"":""powerpoint""," & _
        """timestamp"":" & VTJsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & _
        "}"
    VTWriteTextAtomic statusPath, payload
End Sub
