VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeStudentApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStudentApi

' Test double for the Civil Registry API. Unknown IDs return "not found".

Public CallCount As Long
Public FailAllConnections As Boolean
Public HealthBody As String          ' answer of Health(); "" = the app does not answer
Public HealthCalls As Long

Private responses As Object

Private Sub Class_Initialize()
    Set responses = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RespondFound(ByVal studentId As String, ByVal body As String)
    AddResponse studentId, loFound, 200, body
End Sub

' Same mapping as HttpStudentApi: 401 and 503 have their own outcomes.
Public Sub RespondHttpError(ByVal studentId As String, ByVal httpStatus As Long)
    Select Case httpStatus
        Case 401: AddResponse studentId, loUnauthorized, httpStatus, vbNullString
        Case 503: AddResponse studentId, loUnavailable, httpStatus, vbNullString
        Case Else: AddResponse studentId, loHttpError, httpStatus, vbNullString
    End Select
End Sub

Public Sub RespondConnectionFailed(ByVal studentId As String)
    AddResponse studentId, loConnectionFailed, 0, vbNullString
End Sub

Private Sub AddResponse(ByVal studentId As String, ByVal outcome As LookupOutcome, _
                        ByVal httpStatus As Long, ByVal body As String)
    Dim result As StudentLookupResult
    Set result = New StudentLookupResult
    result.Outcome = outcome
    result.HttpStatus = httpStatus
    result.Body = body
    If outcome = loConnectionFailed Then result.ErrorDescription = "fake: service unreachable"
    Set responses(studentId) = result
End Sub

Private Function IStudentApi_Health() As String
    HealthCalls = HealthCalls + 1
    IStudentApi_Health = HealthBody
End Function

Private Function IStudentApi_FetchStudent(ByVal studentId As String) As StudentLookupResult
    Dim result As StudentLookupResult

    CallCount = CallCount + 1

    If FailAllConnections Then
        Set result = New StudentLookupResult
        result.Outcome = loConnectionFailed
        result.ErrorDescription = "fake: service unreachable"
    ElseIf responses.Exists(studentId) Then
        Set result = responses(studentId)
    Else
        Set result = New StudentLookupResult
        result.Outcome = loNotFound
        result.HttpStatus = 404
    End If

    Set IStudentApi_FetchStudent = result
End Function
