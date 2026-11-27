Attribute VB_Name = "TestGcvLocalApi"
Option Explicit

' Tests for GcvLocalApi (same file in every add-in that calls the GCV app).
' Run RunLocalApiTests from the VBA editor (F5) in the DEV copy. No network.
' Do NOT ship this module or FakeHeaderSink.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunLocalApiTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Local API tests"
    On Error GoTo Crash

    Test_HealthUrl_FromEachBase
    Test_Headers_SendAddinAndKeyWhenPresent
    Test_Headers_LeaveOutAMissingKey

Finish:
    GcvLocalApi.SetKeyForTests Empty
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Local API tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

Private Sub Test_HealthUrl_FromEachBase()
    StartTest "Health URL from the API base"
    AssertEqual "http://127.0.0.1:5555/api/health", HealthUrl("http://127.0.0.1:5555/api/student/"), "student base"
    AssertEqual "http://127.0.0.1:5555/api/health", HealthUrl("http://127.0.0.1:5555/api/"), "names base"
    AssertEqual "http://host:1/x/health", HealthUrl("http://host:1/x"), "other base"
End Sub

Private Sub Test_Headers_SendAddinAndKeyWhenPresent()
    Dim sink As FakeHeaderSink
    StartTest "Requests carry the add-in and the key"
    Set sink = New FakeHeaderSink
    GcvLocalApi.SetKeyForTests "k123"
    GcvLocalApi.AddHeaders sink
    AssertEqual GcvAddin.ADDIN_ID & "/" & GcvAddin.ADDIN_VERSION, sink.HeaderValue(ADDIN_HEADER), "add-in"
    AssertEqual "k123", sink.HeaderValue(KEY_HEADER), "key"
End Sub

Private Sub Test_Headers_LeaveOutAMissingKey()
    Dim sink As FakeHeaderSink
    StartTest "No key yet: no key header"
    Set sink = New FakeHeaderSink
    GcvLocalApi.SetKeyForTests ""
    GcvLocalApi.AddHeaders sink
    AssertEqual False, sink.HasHeader(KEY_HEADER), "no key header"
    AssertEqual 1, sink.HeaderCount, "only the add-in header"
End Sub

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    If CStr(expected) = CStr(actual) Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub
