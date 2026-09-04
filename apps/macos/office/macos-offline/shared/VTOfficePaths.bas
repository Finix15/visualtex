Attribute VB_Name = "VTOfficePaths"
Option Explicit

Private Const VT_WORD_APPLICATION_SCRIPTS_SUFFIX As String = "/Library/Application Scripts/com.microsoft.Word"
Private Const VT_POWERPOINT_APPLICATION_SCRIPTS_SUFFIX As String = "/Library/Application Scripts/com.microsoft.Powerpoint"
Private Const VT_RUNTIME_DIRECTORY_NAME As String = "/VisualTeXRuntime"

Private Function VTUserHomePath() As String
    Dim homePath As String
    Dim sandboxMarker As Long

    homePath = Environ$("HOME")
    sandboxMarker = InStr(1, homePath, "/Library/Containers/", vbTextCompare)
    If sandboxMarker > 1 Then
        homePath = Left$(homePath, sandboxMarker - 1)
    End If
    If Len(homePath) = 0 Or Left$(homePath, 1) <> "/" Then
        Err.Raise vbObjectError + 7100, "VisualTeX", "Unable to resolve the macOS home directory."
    End If
    VTUserHomePath = homePath
End Function

Public Function VTApplicationSupportRoot() As String
    Dim hostName As String
    Dim hostSuffix As String

    hostName = LCase$(Application.Name)
    If InStr(1, hostName, "powerpoint", vbTextCompare) > 0 Then
        hostSuffix = VT_POWERPOINT_APPLICATION_SCRIPTS_SUFFIX
    ElseIf InStr(1, hostName, "word", vbTextCompare) > 0 Then
        hostSuffix = VT_WORD_APPLICATION_SCRIPTS_SUFFIX
    Else
        Err.Raise vbObjectError + 7100, "VisualTeX", "Unable to identify the current Microsoft Office host."
    End If

    ' Each Office host writes its own VisualTeX runtime beneath Application
    ' Scripts. The desktop app can read both host directories as the same user.
    VTApplicationSupportRoot = VTUserHomePath() & hostSuffix & VT_RUNTIME_DIRECTORY_NAME
End Function

Public Function VTWordApplicationScriptsRoot() As String
    VTWordApplicationScriptsRoot = VTUserHomePath() & VT_WORD_APPLICATION_SCRIPTS_SUFFIX
End Function

Public Function VTOfficeGroupContainerRoot() As String
    VTOfficeGroupContainerRoot = VTUserHomePath() & _
        "/Library/Group Containers/UBF8T346G9.Office/VisualTeX"
End Function

Public Function VTOfficeUiLanguage() As String
    Dim fileNumber As Integer
    Dim payload As String
    Dim preferencePath As String

    On Error GoTo UseEnglish
    preferencePath = VTApplicationSupportRoot() & "/ui-language.json"
    If Len(Dir$(preferencePath)) = 0 Then GoTo UseEnglish

    fileNumber = FreeFile
    Open preferencePath For Binary Access Read As #fileNumber
    If LOF(fileNumber) > 4096 Then GoTo CloseAndUseEnglish
    payload = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    fileNumber = 0

    payload = Replace$(payload, " ", vbNullString)
    payload = Replace$(payload, vbCr, vbNullString)
    payload = Replace$(payload, vbLf, vbNullString)
    payload = Replace$(payload, vbTab, vbNullString)
    If InStr(1, payload, """schemaVersion"":1", vbBinaryCompare) = 0 Then GoTo UseEnglish
    If InStr(1, payload, """language"":""vi""", vbBinaryCompare) > 0 Then
        VTOfficeUiLanguage = "vi"
        Exit Function
    End If

UseEnglish:
    VTOfficeUiLanguage = "en"
    Exit Function

CloseAndUseEnglish:
    Close #fileNumber
    fileNumber = 0
    GoTo UseEnglish
End Function
