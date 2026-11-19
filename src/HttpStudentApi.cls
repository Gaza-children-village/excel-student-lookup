VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpStudentApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStudentApi

' Talks to the Civil Registry API served by the GCV application.

Private Const RESOLVE_TIMEOUT_MS As Long = 5000
Private Const CONNECT_TIMEOUT_MS As Long = 5000
Private Const SEND_TIMEOUT_MS As Long = 10000
Private Const RECEIVE_TIMEOUT_MS As Long = 10000
Private Const HTTP_NOT_FOUND As Long = 404

Private Const LOG_SOURCE As String = "student-lookup"

Private baseUrl As String

Public Sub Init(ByVal apiBaseUrl As String)
    If Len(Trim$(apiBaseUrl)) = 0 Then
        Err.Raise vbObjectError + 620, "HttpStudentApi.Init", "The API base URL is empty. Check the Settings sheet."
    End If
    baseUrl = Trim$(apiBaseUrl)
    If Right$(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
End Sub

Private Function IStudentApi_FetchStudent(ByVal studentId As String) As StudentLookupResult
    Dim result As StudentLookupResult
    Dim http As Object
    Dim requestUrl As String
    Dim started As Single

    If Len(baseUrl) = 0 Then
        Err.Raise vbObjectError + 621, "HttpStudentApi.FetchStudent", "Init must be called before FetchStudent."
    End If

    Set result = New StudentLookupResult
    requestUrl = baseUrl & studentId

    ' Network failures (service not running, timeout) are an expected outcome,
    ' returned to the caller as loConnectionFailed and logged below.
    On Error GoTo ConnectionError
    started = Timer
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts RESOLVE_TIMEOUT_MS, CONNECT_TIMEOUT_MS, SEND_TIMEOUT_MS, RECEIVE_TIMEOUT_MS
    http.Open "GET", requestUrl, False
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "X-GCV-Excel-Module-Version", MODULE_VERSION
    http.Send
    On Error GoTo 0

    result.HttpStatus = http.Status
    ' The URL contains the student ID, so it is never logged.
    LogDebug LOG_SOURCE, "Registry request: HTTP " & http.Status & " in " & ElapsedMs(started) & " ms."

    If http.Status = HTTP_NOT_FOUND Then
        result.Outcome = loNotFound
    ElseIf http.Status >= 200 And http.Status < 300 Then
        result.Outcome = loFound
        result.Body = Utf8ResponseText(http)
    Else
        result.Outcome = loHttpError
    End If

    Set IStudentApi_FetchStudent = result
    Exit Function

ConnectionError:
    result.Outcome = loConnectionFailed
    result.ErrorDescription = Err.Description
    LogWarn LOG_SOURCE, "Registry request failed after " & ElapsedMs(started) & " ms: " & Err.Description
    Set IStudentApi_FetchStudent = result
End Function

Private Function Utf8ResponseText(ByVal http As Object) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write http.ResponseBody
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    Utf8ResponseText = stream.ReadText
    stream.Close
End Function

Private Function ElapsedMs(ByVal started As Single) As Long
    Dim seconds As Single
    seconds = Timer - started
    If seconds < 0 Then seconds = seconds + 86400!     ' past midnight
    ElapsedMs = CLng(seconds * 1000)
End Function
