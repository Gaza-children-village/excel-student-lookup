Attribute VB_Name = "LookupUi"
Option Explicit

' ============================================================================
' GCV Student Lookup add-in: start-up, ribbon group "students" and messages.
' The buttons are enabled only on the student sheet of a GCV workbook; Excel
' sheet/workbook switches refresh them (GcvAppEvents). Nothing reacts to typing.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Private Const LOG_SOURCE As String = "student-lookup"
Private Const MB_ICONERROR As Long = &H10
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000

Private mRibbon As Object                ' IRibbonUI
Private mEvents As GcvAppEvents

' ----------------------------------------------------------------------------
' Add-in life cycle
' ----------------------------------------------------------------------------

Public Sub AddinOpened()
    On Error Resume Next                 ' start-up must never show a VBA error
    LogInfo LOG_SOURCE, GcvAddin.ADDIN_NAME & " " & GcvAddin.ADDIN_VERSION & " opened."
    EnsureEvents
End Sub

Public Sub AddinClosing()
    If Not mEvents Is Nothing Then mEvents.Detach
    Set mEvents = Nothing
End Sub

Public Sub AddinSelfCheck()
    StudentRegistry.StatusHeader
    GcvWorkbook.IsStudentSheet Nothing
    T_SelectRows
End Sub

' Re-attaches the sheet-switch events after a VBA reset (they only refresh the
' ribbon buttons; the lookup itself never depends on them).
Private Sub EnsureEvents()
    If mEvents Is Nothing Then
        Set mEvents = New GcvAppEvents
        mEvents.Attach
    End If
End Sub

' ----------------------------------------------------------------------------
' Ribbon callbacks (customUI14.xml)
' ----------------------------------------------------------------------------

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub RefreshRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Public Sub Ribbon_StudentSheetEnabled(control As IRibbonControl, ByRef returnedVal)
    EnsureEvents
    returnedVal = Not GcvWorkbook.ActiveStudentSheet() Is Nothing
End Sub

Public Sub Ribbon_FetchSelected(control As IRibbonControl)
    EnsureEvents
    StudentRegistry.FetchSelectedRows
End Sub

Public Sub Ribbon_FetchPending(control As IRibbonControl)
    EnsureEvents
    StudentRegistry.FetchPendingRows
End Sub

' ----------------------------------------------------------------------------
' Messages (Unicode, right-to-left)
' ----------------------------------------------------------------------------

Public Sub ShowInfo(ByVal text As String)
    Show text, MB_ICONINFO
End Sub

Public Sub ShowWarning(ByVal text As String)
    Show text, MB_ICONWARNING
End Sub

Public Sub ShowError(ByVal text As String)
    Show text, MB_ICONERROR
End Sub

Private Sub Show(ByVal text As String, ByVal icon As Long)
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(FromCodePoints("062C 0644 0628 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0637 0644 0627 0628")), icon Or MB_RIGHT Or MB_RTLREADING
End Sub

Public Function T_OpenStudentSheet() As String
    T_OpenStudentSheet = FromCodePoints("0627 0641 062A 062D 0020 0648 0631 0642 0629 0020 0627 0644 0637 0644 0627 0628 0020 0641 064A 0020 0645 0644 0641 0020 0625 062F 0627 0631 0629 0020 0627 0644 0645 062F 0631 0633 0629 0020 0623 0648 0644 0627 064B 002E")
End Function

Public Function T_SelectRows() As String
    T_SelectRows = FromCodePoints("062D 062F 062F 0020 0635 0641 0648 0641 0020 0627 0644 0637 0644 0627 0628 0020 0627 0644 062A 064A 0020 062A 0631 064A 062F 0020 062C 0644 0628 0020 0628 064A 0627 0646 0627 062A 0647 0627 0020 0028 064A 0643 0641 064A 0020 062A 062D 062F 064A 062F 0020 062E 0644 064A 0629 0020 0641 064A 0020 0643 0644 0020 0635 0641 0029 060C 0020 062B 0645 0020 0627 0636 063A 0637 0020 0627 0644 0632 0631 002E")
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
