VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStudentApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Boundary to the Civil Registry service. HttpStudentApi is the real
' implementation; FakeStudentApi is used by the tests.

Public Function FetchStudent(ByVal studentId As String) As StudentLookupResult
End Function

' GET /api/health of the GCV application, "" when it does not answer.
' Only called after a failed request.
Public Function Health() As String
End Function
