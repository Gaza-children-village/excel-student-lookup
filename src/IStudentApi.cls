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
