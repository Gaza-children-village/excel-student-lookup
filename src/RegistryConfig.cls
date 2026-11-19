VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Settings read from the hidden "Settings" sheet (named cells cfg_*).
' Missing settings fall back to the defaults below, so a workbook without the
' Settings sheet still works.

Public ApiBaseUrl As String
Public IdLength As Long
Public ValidateCheckDigit As Boolean
Public StatusColumn As String
Public LoadedIdColumn As String
Public FirstDataRow As Long

Private Const SETTINGS_SHEET_NAME As String = "Settings"
Private Const NAME_PREFIX As String = "cfg_"

Private Const DEFAULT_API_BASE_URL As String = "http://127.0.0.1:5555/api/student/"
Private Const DEFAULT_ID_LENGTH As Long = 9
Private Const DEFAULT_VALIDATE_CHECK_DIGIT As Boolean = True
Private Const DEFAULT_STATUS_COLUMN As String = "AE"
Private Const DEFAULT_LOADED_ID_COLUMN As String = "AF"
Private Const DEFAULT_FIRST_DATA_ROW As Long = 4

Private Sub Class_Initialize()
    ApplyDefaults
End Sub

Public Sub ApplyDefaults()
    ApiBaseUrl = DEFAULT_API_BASE_URL
    IdLength = DEFAULT_ID_LENGTH
    ValidateCheckDigit = DEFAULT_VALIDATE_CHECK_DIGIT
    StatusColumn = DEFAULT_STATUS_COLUMN
    LoadedIdColumn = DEFAULT_LOADED_ID_COLUMN
    FirstDataRow = DEFAULT_FIRST_DATA_ROW
End Sub

Public Sub LoadFromWorkbook(ByVal wb As Workbook)
    Dim value As Variant

    ApplyDefaults

    value = NamedValue(wb, "ApiBaseUrl"): If HasValue(value) Then ApiBaseUrl = CStr(value)
    value = NamedValue(wb, "IdLength"): If HasValue(value) Then IdLength = CLng(value)
    value = NamedValue(wb, "ValidateCheckDigit"): If HasValue(value) Then ValidateCheckDigit = CBool(value)
    value = NamedValue(wb, "StatusColumn"): If HasValue(value) Then StatusColumn = UCase$(Trim$(CStr(value)))
    value = NamedValue(wb, "LoadedIdColumn"): If HasValue(value) Then LoadedIdColumn = UCase$(Trim$(CStr(value)))
    value = NamedValue(wb, "FirstDataRow"): If HasValue(value) Then FirstDataRow = CLng(value)

    Validate
    LogDebug "student-lookup", "Settings loaded: API " & ApiBaseUrl & ", ID length " & IdLength & _
                               ", check digit " & IIf(ValidateCheckDigit, "on", "off") & _
                               ", status column " & StatusColumn & ", checked-ID column " & LoadedIdColumn & _
                               ", first row " & FirstDataRow & "."
End Sub

Public Sub Validate()
    If IdLength < 1 Then RaiseInvalid "IdLength must be at least 1."
    If FirstDataRow < 2 Then RaiseInvalid "FirstDataRow must be 2 or more."
    ' Template columns A:AD hold student data (V:AC health, AD notes).
    If Not StatusColumn Like "[A-Z]*" Then RaiseInvalid "StatusColumn must be a column letter."
    If Not LoadedIdColumn Like "[A-Z]*" Then RaiseInvalid "LoadedIdColumn must be a column letter."
    If StatusColumn = LoadedIdColumn Then RaiseInvalid "StatusColumn and LoadedIdColumn must differ."
End Sub

' Creates the Settings sheet with the default values if it does not exist.
Public Sub EnsureSettingsSheet(ByVal wb As Workbook)
    Dim ws As Worksheet

    If SheetExists(wb, SETTINGS_SHEET_NAME) Then Exit Sub

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = SETTINGS_SHEET_NAME
    ws.Range("A1:C1").Value = Array("Setting", "Value", "Notes")

    AddSetting wb, ws, 2, "ApiBaseUrl", DEFAULT_API_BASE_URL, "Civil Registry API served by the GCV application"
    AddSetting wb, ws, 3, "IdLength", DEFAULT_ID_LENGTH, "Number of digits in an ID"
    AddSetting wb, ws, 4, "ValidateCheckDigit", DEFAULT_VALIDATE_CHECK_DIGIT, "TRUE = reject IDs with a wrong check digit"
    AddSetting wb, ws, 5, "StatusColumn", DEFAULT_STATUS_COLUMN, "Column that shows the lookup status"
    AddSetting wb, ws, 6, "LoadedIdColumn", DEFAULT_LOADED_ID_COLUMN, "Hidden column: ID whose data is in the row"
    AddSetting wb, ws, 7, "FirstDataRow", DEFAULT_FIRST_DATA_ROW, "First row with student data"

    ws.Columns("A:C").AutoFit
    ' Very hidden: secretaries cannot unhide it from Excel; edit it from the VBA editor.
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub AddSetting(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal rowNumber As Long, _
                       ByVal key As String, ByVal value As Variant, ByVal note As String)
    ws.Cells(rowNumber, 1).Value = key
    ws.Cells(rowNumber, 2).Value = value
    ws.Cells(rowNumber, 3).Value = note
    wb.Names.Add Name:=NAME_PREFIX & key, RefersTo:="='" & ws.Name & "'!$B$" & rowNumber
End Sub

Private Function NamedValue(ByVal wb As Workbook, ByVal key As String) As Variant
    Dim definedName As Name

    ' Workbook.Names raises an error when the name does not exist; a missing
    ' setting is expected and means "use the default".
    On Error Resume Next
    Set definedName = wb.Names(NAME_PREFIX & key)
    On Error GoTo 0

    If definedName Is Nothing Then
        NamedValue = Empty
    Else
        NamedValue = definedName.RefersToRange.Value2
    End If
End Function

Private Function HasValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsError(value) Then Exit Function
    HasValue = (Len(Trim$(CStr(value))) > 0)
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Sub RaiseInvalid(ByVal message As String)
    LogError "student-lookup", "Invalid setting: " & message
    Err.Raise vbObjectError + 630, "RegistryConfig", "Invalid setting: " & message
End Sub
