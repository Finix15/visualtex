Attribute VB_Name = "VTLauncher"
Option Explicit

Public Sub VTLaunchSession(ByVal hostName As String, ByVal sessionId As String)
    Dim scriptName As String
    Dim response As String

    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7300, "VisualTeX", "Invalid VisualTeX Session id."
    End If
    scriptName = VTOfficeLauncherScriptName(hostName)

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXSession", sessionId)
#Else
    Err.Raise vbObjectError + 7302, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7303, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Public Function VTWriteAndLaunchSession( _
    ByVal hostName As String, _
    ByVal sessionId As String, _
    ByVal requestJson As String) As String
    Dim normalizedHost As String
    Dim scriptName As String
    Dim encodedRequest As String
    Dim response As String
    Dim timingDetail As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    operationStage = "normalize-host"
    normalizedHost = VTCanonicalOfficeHost(hostName)
    operationStage = "resolve-script"
    scriptName = VTOfficeLauncherScriptName(normalizedHost)
    operationStage = "validate-request"
    VTValidateRequestPayload sessionId, requestJson
    operationStage = "encode-request"
    encodedRequest = VTBase64UrlEncodeUtf8(requestJson)

#If Mac Then
    operationStage = "applescript-task"
    response = AppleScriptTask( _
        scriptName, _
        "WriteAndOpenVisualTeXSession", _
        normalizedHost & "|" & sessionId & "|" & encodedRequest)
#Else
    Err.Raise vbObjectError + 7307, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    operationStage = "validate-response"
    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7308, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
    timingDetail = Mid$(response, 4)
    If InStr(1, timingDetail, "totalMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "writeMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "launchMs=", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7309, "VisualTeX", _
            "VisualTeX received an invalid write-and-launch timing response."
    End If
    VTWriteAndLaunchSession = timingDetail
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "VisualTeX write and launch", _
        operationStage & ": " & errorDescription
End Function

Public Sub VTPrewarmApplication(ByVal hostName As String)
    Dim normalizedHost As String
    Dim scriptName As String
    Dim response As String

    normalizedHost = VTCanonicalOfficeHost(hostName)
    scriptName = VTOfficeLauncherScriptName(normalizedHost)

#If Mac Then
    response = AppleScriptTask( _
        scriptName, "PrewarmVisualTeXApplication", normalizedHost)
#Else
    Err.Raise vbObjectError + 7310, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7311, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Public Sub VTOpenApplication(ByVal hostName As String)
    Dim scriptName As String
    Dim response As String

    scriptName = VTOfficeLauncherScriptName(hostName)

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXApplication", "")
#Else
    Err.Raise vbObjectError + 7305, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7306, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Private Function VTCanonicalOfficeHost(ByVal hostName As String) As String
    Select Case LCase$(Trim$(hostName))
        Case "word", "powerpoint"
            VTCanonicalOfficeHost = LCase$(Trim$(hostName))
        Case Else
            Err.Raise vbObjectError + 7301, "VisualTeX", "Invalid VisualTeX Office host."
    End Select
End Function

Private Function VTOfficeLauncherScriptName(ByVal hostName As String) As String
    Select Case VTCanonicalOfficeHost(hostName)
        Case "word": VTOfficeLauncherScriptName = "VisualTeXWord.scpt"
        Case "powerpoint": VTOfficeLauncherScriptName = "VisualTeXPowerPoint.scpt"
    End Select
End Function

Private Function VTAppleScriptErrorMessage(ByVal response As String) As String
    Dim fields() As String
    If Left$(response, 6) = "error|" Then
        fields = Split(response, "|")
        If UBound(fields) >= 2 Then
            VTAppleScriptErrorMessage = fields(2)
            Exit Function
        End If
    End If
    VTAppleScriptErrorMessage = "VisualTeX could not be opened."
End Function
