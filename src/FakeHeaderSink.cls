VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeHeaderSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Test double for a WinHttpRequest: keeps the headers set on it.

Private mHeaders As Object

Private Sub Class_Initialize()
    Set mHeaders = CreateObject("Scripting.Dictionary")
End Sub

Public Sub SetRequestHeader(ByVal headerName As String, ByVal headerValue As String)
    mHeaders(headerName) = headerValue
End Sub

Public Function HeaderValue(ByVal headerName As String) As String
    If mHeaders.Exists(headerName) Then HeaderValue = mHeaders(headerName)
End Function

Public Function HasHeader(ByVal headerName As String) As Boolean
    HasHeader = mHeaders.Exists(headerName)
End Function

Public Property Get HeaderCount() As Long
    HeaderCount = mHeaders.Count
End Property
