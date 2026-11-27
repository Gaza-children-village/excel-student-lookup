Attribute VB_Name = "GcvLocalApi"
Option Explicit

' ============================================================================
' Requests to the GCV desktop app on this PC (127.0.0.1:5555). The same file
' is in every GCV add-in that calls the app (student lookup, English names).
'
'   X-GCV-Excel-Addin  <add-in id>/<version>, on every request
'   X-GCV-Api-Key      the key the app writes at every start to
'                      HKCU\Software\GCV Tools\LocalApiKey; sent when present
'                      (the app requires it only once all PCs have add-ins
'                      that send it; then a missing or wrong key gives 401)
'   GET .../api/health only after a request failed, for a clearer message:
'                      {"registry": "ready"|"unavailable", "names": ..., "version": "x.y.z"}
'
' The key is never logged or shown.
' ============================================================================

Public Const KEY_REGISTRY_VALUE As String = "HKCU\Software\GCV Tools\LocalApiKey"
Public Const KEY_HEADER As String = "X-GCV-Api-Key"
Public Const ADDIN_HEADER As String = "X-GCV-Excel-Addin"
Public Const HTTP_UNAUTHORIZED As Long = 401
Public Const HTTP_UNAVAILABLE As Long = 503

Private Const HEALTH_CONNECT_MS As Long = 1500
Private Const HEALTH_ANSWER_MS As Long = 2000

Private mKeyForTests As Variant              ' Empty: read the registry

' The app's key, or "" when the app has not written one (yet).
Public Function LocalApiKey() As String
    Dim registryValue As Variant
    If Not IsEmpty(mKeyForTests) Then
        LocalApiKey = CStr(mKeyForTests)
        Exit Function
    End If
    On Error Resume Next                     ' no value yet: the key is optional for now
    registryValue = CreateObject("WScript.Shell").RegRead(KEY_REGISTRY_VALUE)
    If Err.Number = 0 Then LocalApiKey = Trim$(CStr(registryValue))
    On Error GoTo 0
End Function

' Tests only: a fixed key ("" = none); Empty reads the registry again.
Public Sub SetKeyForTests(ByVal keyValue As Variant)
    mKeyForTests = keyValue
End Sub

Public Function AddinHeaderValue() As String
    AddinHeaderValue = GcvAddin.ADDIN_ID & "/" & GcvAddin.ADDIN_VERSION
End Function

' Adds the add-in header and (when present) the key to an opened request.
Public Sub AddHeaders(ByVal http As Object)
    Dim key As String
    http.SetRequestHeader ADDIN_HEADER, AddinHeaderValue()
    key = LocalApiKey()
    If Len(key) > 0 Then http.SetRequestHeader KEY_HEADER, key
End Sub

' ".../api/student/" or ".../api/" -> ".../api/health".
Public Function HealthUrl(ByVal baseUrl As String) As String
    Dim position As Long
    position = InStr(1, baseUrl, "/api/", vbTextCompare)
    If position > 0 Then
        HealthUrl = Left$(baseUrl, position + 4) & "health"
    Else
        If Right$(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
        HealthUrl = baseUrl & "health"
    End If
End Function

' The answer of GET /api/health, or "" when the app does not answer.
Public Function ReadHealth(ByVal url As String) As String
    Dim http As Object
    On Error GoTo NoAnswer
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts HEALTH_CONNECT_MS, HEALTH_CONNECT_MS, HEALTH_ANSWER_MS, HEALTH_ANSWER_MS
    http.Open "GET", url, False
    http.SetRequestHeader "Accept", "application/json"
    AddHeaders http
    http.Send
    If http.Status = 200 Then ReadHealth = http.ResponseText
NoAnswer:
End Function
