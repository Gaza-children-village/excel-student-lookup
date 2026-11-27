VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Counts what happened to each row of one lookup batch.

Public Total As Long
Public Cleared As Long
Public Found As Long
Public FoundKeptManual As Long
Public Invalid As Long
Public NotFound As Long
Public FailedKeptData As Long
Public Pending As Long

Public ConnectionDown As Boolean
Public Unauthorized As Boolean       ' 401 from the GCV application
Public ServiceUnavailable As Boolean ' 503: registry not ready
Public AppVersion As String          ' from /api/health after a failure
Public RegistryState As String       ' "ready" / "unavailable" / "" (no answer)
Public LastResult As RowResult
Public LastHttpStatus As Long
Public LastError As String

Public Sub Record(ByVal result As RowResult)
    Total = Total + 1
    LastResult = result
    Select Case result
        Case rrCleared: Cleared = Cleared + 1
        Case rrFound: Found = Found + 1
        Case rrFoundKeptManual: FoundKeptManual = FoundKeptManual + 1
        Case rrInvalid: Invalid = Invalid + 1
        Case rrNotFound: NotFound = NotFound + 1
        Case rrFailedKeptData: FailedKeptData = FailedKeptData + 1
        Case rrPending: Pending = Pending + 1
        Case Else
            Err.Raise vbObjectError + 610, "BatchSummary.Record", "Unknown row result: " & result
    End Select
End Sub

' No more requests in this batch: they would fail the same way.
Public Property Get StopRequests() As Boolean
    StopRequests = ConnectionDown Or Unauthorized Or ServiceUnavailable
End Property

Public Property Get ProblemCount() As Long
    ProblemCount = Invalid + NotFound + FailedKeptData + Pending
End Property
