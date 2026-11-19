Attribute VB_Name = "GcvWorkbook"
Option Explicit

' ============================================================================
' Recognising GCV workbooks. The add-in reacts only to workbooks that carry
' these hidden defined names (the template has them; SetupStudentSheet adds
' them to an older copy):
'
'   GCV_TEMPLATE        ="GCV Administration 1.1"   template name and version
'   GCV_STUDENT_SHEET   ='<student sheet>'!$A$1     which sheet holds students
'
' The sheet reference follows renames and moves, so secretaries may rename
' the sheet. Every other workbook is left alone.
' ============================================================================

Public Const TEMPLATE_NAME As String = "GCV_TEMPLATE"
Public Const STUDENT_SHEET_NAME As String = "GCV_STUDENT_SHEET"
Public Const TEMPLATE_TEXT As String = "GCV Administration 1.1"

Public Function IsGcvWorkbook(ByVal wb As Workbook) As Boolean
    IsGcvWorkbook = Not StudentSheetOf(wb) Is Nothing
End Function

' The student sheet of a GCV workbook, or Nothing.
Public Function StudentSheetOf(ByVal wb As Workbook) As Worksheet
    Dim definedName As Name

    If wb Is Nothing Then Exit Function
    ' Names(...) raises when the name is missing; RefersToRange raises when the
    ' reference is broken (#REF!). Both mean "not a GCV student sheet".
    On Error Resume Next
    Set definedName = wb.Names(STUDENT_SHEET_NAME)
    If Not definedName Is Nothing Then Set StudentSheetOf = definedName.RefersToRange.Worksheet
    Err.Clear
    On Error GoTo 0
End Function

Public Function IsStudentSheet(ByVal sh As Object) As Boolean
    Dim ws As Worksheet
    If sh Is Nothing Then Exit Function
    If TypeName(sh) <> "Worksheet" Then Exit Function
    Set ws = StudentSheetOf(sh.Parent)
    If ws Is Nothing Then Exit Function
    ' Same workbook (names are unique in Excel) and same sheet name.
    IsStudentSheet = (ws.Name = sh.Name And ws.Parent.Name = sh.Parent.Name)
End Function

' The active sheet when it is a GCV student sheet, otherwise Nothing.
Public Function ActiveStudentSheet() As Worksheet
    Dim sh As Object
    On Error Resume Next                        ' no workbook open: ActiveSheet raises
    Set sh = ActiveSheet
    On Error GoTo 0
    If IsStudentSheet(sh) Then Set ActiveStudentSheet = sh
End Function

' Marks `ws` as the student sheet of its workbook (developer setup).
Public Sub MarkStudentSheet(ByVal ws As Worksheet)
    Dim wb As Workbook
    Set wb = ws.Parent
    On Error Resume Next
    wb.Names(STUDENT_SHEET_NAME).Delete
    wb.Names(TEMPLATE_NAME).Delete
    On Error GoTo 0
    wb.Names.Add Name:=STUDENT_SHEET_NAME, _
                 RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!$A$1", Visible:=False
    wb.Names.Add Name:=TEMPLATE_NAME, RefersTo:="=""" & TEMPLATE_TEXT & """", Visible:=False
End Sub
