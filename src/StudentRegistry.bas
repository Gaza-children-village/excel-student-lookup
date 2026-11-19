Attribute VB_Name = "StudentRegistry"
Option Explicit

' ============================================================================
' GCV student lookup: fills student rows from the local Civil Registry API.
' GCV Student Lookup add-in. ON DEMAND only: nothing runs while the secretary
' types (so Ctrl+Z keeps working and saving or editing never sends requests).
' Ribbon buttons: FetchSelectedRows and FetchPendingRows, on the student sheet
' of a GCV workbook (see GcvWorkbook); the workbook itself has no code.
'
' Arabic text is built from Unicode code points (FromCodePoints) because the
' classic VBA editor does not preserve Arabic characters in source code.
' Version history: CHANGELOG.txt
' ============================================================================

Public Const MODULE_VERSION As String = "3.0.0"
Private Const LOG_SOURCE As String = "student-lookup"   ' gcv/logger source name

Public Enum LookupOutcome
    loFound = 1
    loNotFound = 2
    loHttpError = 3
    loConnectionFailed = 4
End Enum

Public Enum IdValidation
    ivValid = 0
    ivEmpty = 1
    ivNotDigits = 2
    ivWrongLength = 3
    ivBadCheckDigit = 4
End Enum

Public Enum RowResult
    rrCleared = 1
    rrFound = 2
    rrFoundKeptManual = 3
    rrInvalid = 4
    rrNotFound = 5
    rrFailedKeptData = 6
    rrPending = 7
End Enum

Private Enum FieldKind
    fkText = 1
    fkNumber = 2
    fkIsoDate = 3
    fkGender = 4
    fkIdText = 5
End Enum

Private Type FieldSpec
    ColumnLetter As String
    JsonKey As String
    Kind As FieldKind
End Type

Public Const ID_COLUMN As String = "B"
Private Const DATE_FORMAT As String = "dd/mm/yyyy"   ' same display as the GCV Administration template
Private Const TEXT_FORMAT As String = "@"
Private Const ISO_DATE_LENGTH As Long = 10

' ----------------------------------------------------------------------------
' Entry points
' ----------------------------------------------------------------------------

' Ribbon "fetch selected rows": looks up the rows of the current selection on
' the student sheet, also rows already filled (refresh; manual edits are
' kept). A row whose ID was deleted has its automatic data cleared.
Public Sub FetchSelectedRows()
    Dim ws As Worksheet
    Dim config As RegistryConfig
    Dim idCells As Range

    Set ws = GcvWorkbook.ActiveStudentSheet()
    If ws Is Nothing Then
        LookupUi.ShowInfo LookupUi.T_OpenStudentSheet()
        Exit Sub
    End If
    Set config = New RegistryConfig
    config.LoadFromWorkbook ws.Parent

    If TypeName(Selection) = "Range" Then Set idCells = SelectedIdCells(ws, Selection, config)
    If idCells Is Nothing Then
        LogInfo LOG_SOURCE, "Fetch selected rows: no student row selected."
        LookupUi.ShowInfo LookupUi.T_SelectRows()
        Exit Sub
    End If
    LogInfo LOG_SOURCE, "Fetch selected rows: " & CellCount(idCells) & " row(s)."
    RunLookup ws, idCells, config
End Sub

' ID cells of the selected rows within the student data (rows from
' FirstDataRow to the last row that holds an ID or lookup data).
Public Function SelectedIdCells(ByVal ws As Worksheet, ByVal selected As Range, _
                                ByVal config As RegistryConfig) As Range
    If Not selected.Worksheet Is ws Then
        If selected.Worksheet.Name <> ws.Name Then Exit Function
    End If
    Set SelectedIdCells = ChangedIdCells(ws, Intersect(selected.EntireRow, ws.Columns(ID_COLUMN)), config)
End Function

' Ribbon "fetch new rows": every row whose data is missing, belongs to a
' changed ID, or waits after a failed connection.
Public Sub FetchPendingRows()
    Dim ws As Worksheet
    Dim config As RegistryConfig
    Dim idCells As Range

    Set ws = GcvWorkbook.ActiveStudentSheet()
    If ws Is Nothing Then
        LookupUi.ShowInfo LookupUi.T_OpenStudentSheet()
        Exit Sub
    End If
    Set config = New RegistryConfig
    config.LoadFromWorkbook ws.Parent

    Set idCells = PendingIdCells(ws, config)
    If idCells Is Nothing Then
        LogInfo LOG_SOURCE, "Fetch new rows: no rows are waiting."
        LookupUi.ShowInfo FromCodePoints("0644 0627 0020 062A 0648 062C 062F 0020 0635 0641 0648 0641 0020 062A 062D 062A 0627 062C 0020 0625 0644 0649 0020 062C 0644 0628 0020 0627 0644 0628 064A 0627 0646 0627 062A 002E")
        Exit Sub
    End If

    LogInfo LOG_SOURCE, "Fetch new rows: " & CellCount(idCells) & " row(s) waiting."
    RunLookup ws, idCells, config
End Sub

' One-time preparation of the student sheet (run by the developer, not by
' secretaries, with the student sheet active): marks it as the GCV student
' sheet (GcvWorkbook), creates the Settings sheet, formats the ID column as text,
' adds the status header, hides the loaded-ID column and records which rows
' already hold data so existing rows are not treated as new students.
Public Sub SetupStudentSheet()
    Dim ws As Worksheet
    Dim config As RegistryConfig

    Set ws = ActiveSheet
    Set config = New RegistryConfig
    config.EnsureSettingsSheet ws.Parent
    config.LoadFromWorkbook ws.Parent

    ' Rewriting numeric IDs must not trigger lookups during setup.
    On Error GoTo Failed
    Application.EnableEvents = False
    PrepareStudentSheet ws, config
    GcvWorkbook.MarkStudentSheet ws
    Application.EnableEvents = True
    ws.Activate
    LogInfo LOG_SOURCE, "Student sheet prepared (module " & MODULE_VERSION & ")."

    MsgBox "Student sheet prepared." & vbCrLf & _
           "Status column: " & config.StatusColumn & vbCrLf & _
           "Hidden loaded-ID column: " & config.LoadedIdColumn & vbCrLf & _
           "Module version: " & MODULE_VERSION, vbInformation, "Student lookup setup"
    Exit Sub

Failed:
    Application.EnableEvents = True
    LogError LOG_SOURCE, "Student sheet setup failed"
    Err.Raise Err.Number, "StudentRegistry.SetupStudentSheet", Err.Description
End Sub

Public Sub PrepareStudentSheet(ByVal ws As Worksheet, ByVal config As RegistryConfig)
    Dim headerCell As Range
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim studentId As String

    lastRow = LastIdRow(ws, config)

    ' Record existing IDs before reformatting, so numeric IDs are normalised.
    For rowNumber = config.FirstDataRow To lastRow
        studentId = NormalizeIdCell(ws.Cells(rowNumber, ID_COLUMN), config.IdLength)
        If Len(studentId) > 0 And HasAutomaticData(ws, rowNumber) _
           And Len(LoadedIdOf(ws, rowNumber, config)) = 0 Then
            WriteLoadedId ws, rowNumber, config, studentId
        End If
    Next rowNumber

    ws.Range(ws.Cells(config.FirstDataRow, ID_COLUMN), _
             ws.Cells(ws.Rows.Count, ID_COLUMN)).NumberFormat = TEXT_FORMAT
    ws.Range(ws.Cells(config.FirstDataRow, config.LoadedIdColumn), _
             ws.Cells(ws.Rows.Count, config.LoadedIdColumn)).NumberFormat = TEXT_FORMAT

    If config.FirstDataRow > 1 Then
        Set headerCell = ws.Cells(config.FirstDataRow - 1, config.StatusColumn)
        If IsEmpty(headerCell.Value2) Then headerCell.Value = StatusHeader()
    End If

    ws.Columns(config.LoadedIdColumn).Hidden = True
End Sub

' ----------------------------------------------------------------------------
' Batch processing (testable: the API is injected)
' ----------------------------------------------------------------------------

Public Function ProcessIdCells(ByVal ws As Worksheet, ByVal idCells As Range, _
                               ByVal api As IStudentApi, _
                               ByVal config As RegistryConfig) As BatchSummary
    Dim summary As BatchSummary
    Dim cell As Range
    Dim position As Long
    Dim total As Long

    Set summary = New BatchSummary
    total = CellCount(idCells)

    For Each cell In idCells.Cells
        position = position + 1
        If total > 1 Then Application.StatusBar = "Student lookup: " & position & " / " & total
        ProcessRow ws, cell.Row, api, config, summary
        ' Row number and result only - never the ID or any student data.
        LogDebug LOG_SOURCE, "Row " & cell.Row & ": " & RowResultName(summary.LastResult)
    Next cell

    If total > 1 Then Application.StatusBar = False
    Set ProcessIdCells = summary
End Function

' Rows that need a lookup: an ID is present and either it was never checked
' (or changed since), or the last attempt could not reach the API.
' Rows already answered "invalid" or "not found" for the same ID are skipped.
Public Function PendingIdCells(ByVal ws As Worksheet, ByVal config As RegistryConfig) As Range
    Dim result As Range
    Dim rowNumber As Long
    Dim studentId As String
    Dim statusText As String

    For rowNumber = config.FirstDataRow To LastIdRow(ws, config)
        studentId = NormalizeIdValue(ws.Cells(rowNumber, ID_COLUMN).Value2, config.IdLength)
        statusText = CStr(ws.Cells(rowNumber, config.StatusColumn).Value2)

        If Len(studentId) > 0 Then
            If LoadedIdOf(ws, rowNumber, config) <> studentId _
               Or statusText = StatusFailedKeptData() Then
                If result Is Nothing Then
                    Set result = ws.Cells(rowNumber, ID_COLUMN)
                Else
                    Set result = Union(result, ws.Cells(rowNumber, ID_COLUMN))
                End If
            End If
        End If
    Next rowNumber

    Set PendingIdCells = result
End Function

Private Sub RunLookup(ByVal ws As Worksheet, ByVal idCells As Range, ByVal config As RegistryConfig)
    Dim httpApi As HttpStudentApi
    Dim api As IStudentApi
    Dim summary As BatchSummary

    Set httpApi = New HttpStudentApi
    httpApi.Init config.ApiBaseUrl
    Set api = httpApi

    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    LogInfo LOG_SOURCE, "Lookup started: " & CellCount(idCells) & " row(s)."
    Set summary = ProcessIdCells(ws, idCells, api, config)
    LogSummary summary

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not summary Is Nothing Then ShowSummary summary, config
    Exit Sub

Failed:
    LogError LOG_SOURCE, "Lookup stopped by an unexpected error"
    LookupUi.ShowError FromCodePoints("062A 0648 0642 0641 0020 062C 0644 0628 0020 0627 0644 0628 064A 0627 0646 0627 062A 0020 0628 0633 0628 0628 0020 062E 0637 0623 0020 063A 064A 0631 0020 0645 062A 0648 0642 0639 003A") & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                       FromCodePoints("064A 0645 0643 0646 0020 062C 0644 0628 0020 0627 0644 0635 0641 0648 0641 0020 063A 064A 0631 0020 0627 0644 0645 0643 062A 0645 0644 0629 0020 0644 0627 062D 0642 0627 064B 0020 0628 0632 0631 0020 00AB 062C 0644 0628 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0635 0641 0648 0641 0020 0627 0644 062C 062F 064A 062F 0629 00BB 002E")
    Resume Cleanup
End Sub

Private Sub LogSummary(ByVal summary As BatchSummary)
    LogInfo LOG_SOURCE, "Lookup finished: " & summary.Found & " found, " & summary.FoundKeptManual & _
                        " found with manual values kept, " & summary.NotFound & " not found, " & _
                        summary.Invalid & " invalid ID, " & (summary.Pending + summary.FailedKeptData) & _
                        " waiting for retry, " & summary.Cleared & " cleared."
    If summary.ConnectionDown Then
        LogWarn LOG_SOURCE, "Registry service unreachable: " & summary.LastError
    ElseIf summary.LastHttpStatus <> 0 Then
        LogWarn LOG_SOURCE, "Registry service returned HTTP " & summary.LastHttpStatus & "."
    End If
End Sub

Public Function RowResultName(ByVal result As RowResult) As String
    Select Case result
        Case rrCleared: RowResultName = "cleared"
        Case rrFound: RowResultName = "found"
        Case rrFoundKeptManual: RowResultName = "found, manual values kept"
        Case rrInvalid: RowResultName = "invalid ID"
        Case rrNotFound: RowResultName = "not found"
        Case rrFailedKeptData: RowResultName = "no connection, previous data kept"
        Case rrPending: RowResultName = "waiting for retry"
        Case Else: RowResultName = "unknown (" & result & ")"
    End Select
End Function

Private Sub ProcessRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal api As IStudentApi, _
                       ByVal config As RegistryConfig, ByVal summary As BatchSummary)
    Dim studentId As String
    Dim sameStudent As Boolean
    Dim result As StudentLookupResult
    Dim keptCount As Long

    studentId = NormalizeIdCell(ws.Cells(rowNumber, ID_COLUMN), config.IdLength)

    Select Case ValidateStudentId(studentId, config.IdLength, config.ValidateCheckDigit)
        Case ivEmpty
            ClearRowState ws, rowNumber, config
            summary.Record rrCleared
            Exit Sub
        Case ivValid
            ' continue
        Case Else
            ClearAutomaticFields ws, rowNumber
            SetRowState ws, rowNumber, config, StatusInvalidId(), studentId
            summary.Record rrInvalid
            Exit Sub
    End Select

    sameStudent = (LoadedIdOf(ws, rowNumber, config) = studentId)

    ' After a connection failure, do not wait for a timeout on every row.
    If summary.ConnectionDown Then
        MarkUnreachable ws, rowNumber, config, sameStudent, summary, StatusPending()
        Exit Sub
    End If

    Set result = api.FetchStudent(studentId)

    Select Case result.Outcome
        Case loFound
            ' A different student: the previous student's details must go.
            If Not sameStudent Then ClearAutomaticFields ws, rowNumber
            keptCount = WriteStudent(ws, rowNumber, result.Body, config.IdLength)
            If keptCount > 0 Then
                SetRowState ws, rowNumber, config, StatusFoundKeptManual(keptCount), studentId
                summary.Record rrFoundKeptManual
            Else
                SetRowState ws, rowNumber, config, StatusFound(), studentId
                summary.Record rrFound
            End If

        Case loNotFound
            ClearAutomaticFields ws, rowNumber
            SetRowState ws, rowNumber, config, StatusNotFound(), studentId
            summary.Record rrNotFound

        Case loHttpError
            summary.LastHttpStatus = result.HttpStatus
            MarkUnreachable ws, rowNumber, config, sameStudent, summary, _
                            StatusPending() & " (HTTP " & result.HttpStatus & ")"

        Case loConnectionFailed
            summary.ConnectionDown = True
            summary.LastError = result.ErrorDescription
            MarkUnreachable ws, rowNumber, config, sameStudent, summary, StatusPending()

        Case Else
            Err.Raise vbObjectError + 600, "StudentRegistry.ProcessRow", _
                      "Unknown lookup outcome: " & result.Outcome
    End Select
End Sub

' Data that belongs to the same ID is kept when the API cannot be reached.
' Data that belongs to a different (previous) ID is cleared, because keeping it
' next to the new ID would show the wrong student's details.
Private Sub MarkUnreachable(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                            ByVal config As RegistryConfig, ByVal sameStudent As Boolean, _
                            ByVal summary As BatchSummary, ByVal pendingStatus As String)
    If sameStudent And HasAutomaticData(ws, rowNumber) Then
        SetRowState ws, rowNumber, config, StatusFailedKeptData(), LoadedIdOf(ws, rowNumber, config)
        summary.Record rrFailedKeptData
    Else
        ClearAutomaticFields ws, rowNumber
        SetRowState ws, rowNumber, config, pendingStatus, vbNullString
        summary.Record rrPending
    End If
End Sub

Private Sub ShowSummary(ByVal summary As BatchSummary, ByVal config As RegistryConfig)
    Dim message As String

    If summary.ProblemCount = 0 Then Exit Sub

    If summary.Total = 1 Then
        If summary.Invalid = 1 Then
            message = FromCodePoints("0631 0642 0645 0020 0627 0644 0647 0648 064A 0629 0020 063A 064A 0631 0020 0635 062D 064A 062D 002E") & vbCrLf & _
                      FromCodePoints("064A 062C 0628 0020 0623 0646 0020 064A 062A 0643 0648 0646 0020 0645 0646 0020") & config.IdLength & FromCodePoints("0020 0623 0631 0642 0627 0645") & _
                      IIf(config.ValidateCheckDigit, FromCodePoints("0020 0648 0623 0646 0020 064A 0643 0648 0646 0020 0631 0642 0645 0020 0627 0644 062A 062D 0642 0642 0020 0028 0622 062E 0631 0020 0631 0642 0645 0029 0020 0635 062D 064A 062D 0627 064B 002E"), ".")
            LookupUi.ShowWarning message
            Exit Sub
        End If
        If summary.NotFound = 1 Then
            LookupUi.ShowInfo FromCodePoints("0644 0645 0020 064A 064F 0639 062B 0631 0020 0639 0644 0649 0020 0637 0627 0644 0628 0020 0628 0647 0630 0627 0020 0627 0644 0631 0642 0645 0020 0641 064A 0020 0627 0644 0633 062C 0644 0020 0627 0644 0645 062F 0646 064A 002E")
            Exit Sub
        End If
    End If

    message = FromCodePoints("0627 0644 0635 0641 0648 0641 0020 0627 0644 062A 064A 0020 062A 0645 062A 0020 0645 0639 0627 0644 062C 062A 0647 0627 003A 0020") & summary.Total & vbCrLf & _
              FromCodePoints("062A 0645 0020 062C 0644 0628 0020 0628 064A 0627 0646 0627 062A 0647 0627 003A 0020") & (summary.Found + summary.FoundKeptManual) & vbCrLf & _
              FromCodePoints("063A 064A 0631 0020 0645 0648 062C 0648 062F 0629 0020 0641 064A 0020 0627 0644 0633 062C 0644 0020 0627 0644 0645 062F 0646 064A 003A 0020") & summary.NotFound & vbCrLf & _
              FromCodePoints("0631 0642 0645 0020 0647 0648 064A 0629 0020 063A 064A 0631 0020 0635 062D 064A 062D 003A 0020") & summary.Invalid & vbCrLf & _
              FromCodePoints("0628 0627 0646 062A 0638 0627 0631 0020 0625 0639 0627 062F 0629 0020 0627 0644 0645 062D 0627 0648 0644 0629 003A 0020") & (summary.Pending + summary.FailedKeptData)

    If summary.ConnectionDown Then
        message = message & vbCrLf & vbCrLf & _
                  FromCodePoints("062A 0639 0630 0631 0020 0627 0644 0627 062A 0635 0627 0644 0020 0628 062E 062F 0645 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0637 0644 0627 0628 002E 0020 062A 0623 0643 062F 0020 0645 0646 0020 062A 0634 063A 064A 0644 0020 062A 0637 0628 064A 0642 0020 0047 0043 0056 060C") & vbCrLf & _
                  FromCodePoints("062B 0645 0020 0627 0636 063A 0637 0020 0632 0631 0020 00AB 062C 0644 0628 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0635 0641 0648 0641 0020 0627 0644 062C 062F 064A 062F 0629 00BB 002E") & vbCrLf & config.ApiBaseUrl
    ElseIf summary.LastHttpStatus <> 0 Then
        message = message & vbCrLf & vbCrLf & _
                  FromCodePoints("0623 0639 0627 062F 062A 0020 062E 062F 0645 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0637 0644 0627 0628 0020 062E 0637 0623 0020 0028 0048 0054 0054 0050 0020") & summary.LastHttpStatus & ")." & vbCrLf & _
                  FromCodePoints("062D 0627 0648 0644 0020 0644 0627 062D 0642 0627 064B 060C 0020 0623 0648 0020 0627 0633 062A 062E 062F 0645 0020 0632 0631 0020 00AB 0627 0644 0625 0628 0644 0627 063A 0020 0639 0646 0020 0645 0634 0643 0644 0629 00BB 002E")
    End If

    If summary.ConnectionDown Then LookupUi.ShowError message Else LookupUi.ShowWarning message
End Sub

' ----------------------------------------------------------------------------
' IDs
' ----------------------------------------------------------------------------

' Returns the ID as text. Numeric cell values lose leading zeros in Excel, so
' they are padded back to the configured length.
Public Function NormalizeIdValue(ByVal rawValue As Variant, ByVal idLength As Long) As String
    Dim text As String

    If IsEmpty(rawValue) Or IsError(rawValue) Then Exit Function

    If IsNumberType(rawValue) Then
        If rawValue >= 0 And rawValue = Int(rawValue) Then
            text = Format$(rawValue, "0")
            If Len(text) < idLength Then text = String$(idLength - Len(text), "0") & text
            NormalizeIdValue = text
            Exit Function
        End If
    End If

    NormalizeIdValue = Trim$(CStr(rawValue))
End Function

' Normalises the ID and writes it back as text when Excel stored it as a number.
Private Function NormalizeIdCell(ByVal idCell As Range, ByVal idLength As Long) As String
    Dim studentId As String

    studentId = NormalizeIdValue(idCell.Value2, idLength)
    If IsNumberType(idCell.Value2) And Len(studentId) > 0 Then
        idCell.NumberFormat = TEXT_FORMAT
        idCell.Value = studentId
    End If
    NormalizeIdCell = studentId
End Function

Public Function ValidateStudentId(ByVal studentId As String, ByVal idLength As Long, _
                                  ByVal checkDigitRequired As Boolean) As IdValidation
    If Len(studentId) = 0 Then
        ValidateStudentId = ivEmpty
    ElseIf Not studentId Like String$(Len(studentId), "#") Then
        ValidateStudentId = ivNotDigits
    ElseIf Len(studentId) <> idLength Then
        ValidateStudentId = ivWrongLength
    ElseIf checkDigitRequired And Not IsValidCheckDigit(studentId) Then
        ValidateStudentId = ivBadCheckDigit
    Else
        ValidateStudentId = ivValid
    End If
End Function

' Luhn-style check: digits in even positions (2nd, 4th, ...) are doubled,
' digits of results above 9 are summed, and the total must divide by 10.
' TO VERIFY against the full Civil Registry before relying on it in schools.
Public Function IsValidCheckDigit(ByVal studentId As String) As Boolean
    Dim position As Long
    Dim digit As Long
    Dim total As Long

    If Len(studentId) = 0 Or Not studentId Like String$(Len(studentId), "#") Then Exit Function

    For position = 1 To Len(studentId)
        digit = CLng(Mid$(studentId, position, 1))
        If position Mod 2 = 0 Then digit = digit * 2
        If digit > 9 Then digit = digit - 9
        total = total + digit
    Next position

    IsValidCheckDigit = (total Mod 10 = 0)
End Function

' ----------------------------------------------------------------------------
' Writing student data
' ----------------------------------------------------------------------------

Private Function StudentFields() As FieldSpec()
    Dim fields() As FieldSpec
    ReDim fields(1 To 11)

    SetField fields(1), "C", "full_name", fkText
    SetField fields(2), "D", "english_name", fkText
    SetField fields(3), "E", "birthdate", fkIsoDate
    SetField fields(4), "H", "gender", fkGender
    SetField fields(5), "M", "father_name", fkText
    SetField fields(6), "N", "father_id", fkIdText
    SetField fields(7), "P", "mother_name", fkText
    SetField fields(8), "Q", "mother_id", fkIdText
    SetField fields(9), "R", "place_before_displacement", fkText
    SetField fields(10), "T", "family_members_count", fkNumber
    SetField fields(11), "U", "order_among_siblings", fkNumber

    StudentFields = fields
End Function

Private Sub SetField(ByRef spec As FieldSpec, ByVal columnLetter As String, _
                     ByVal jsonKey As String, ByVal kind As FieldKind)
    spec.ColumnLetter = columnLetter
    spec.JsonKey = jsonKey
    spec.Kind = kind
End Sub

' Fills empty automatic cells only. A non-empty cell that differs from the
' registry value is treated as a manual correction and kept; the number of
' kept cells is returned so the status can show it. A new ID never reaches
' this point with old data, because the row is cleared first.
Private Function WriteStudent(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                              ByVal body As String, ByVal idLength As Long) As Long
    Dim fields() As FieldSpec
    Dim index As Long
    Dim target As Range
    Dim newValue As Variant
    Dim keptCount As Long

    fields = StudentFields()

    For index = LBound(fields) To UBound(fields)
        Set target = ws.Cells(rowNumber, fields(index).ColumnLetter)
        newValue = FieldValue(fields(index).Kind, JsonText(body, fields(index).JsonKey), idLength)

        If IsEmpty(target.Value2) Then
            If Not IsEmpty(newValue) Then WriteField target, fields(index).Kind, newValue
        ElseIf Not ValuesEqual(target.Value2, newValue) Then
            keptCount = keptCount + 1
        End If
    Next index

    WriteStudent = keptCount
End Function

Private Function FieldValue(ByVal kind As FieldKind, ByVal text As String, _
                            ByVal idLength As Long) As Variant
    FieldValue = Empty
    If Len(text) = 0 Then Exit Function

    Select Case kind
        Case fkText
            FieldValue = text
        Case fkNumber
            If IsNumeric(text) Then FieldValue = CDbl(text) Else FieldValue = text
        Case fkIsoDate
            FieldValue = ParseIsoDate(text)
        Case fkGender
            FieldValue = ArabicGender(text)
        Case fkIdText
            ' Parent IDs may arrive as JSON numbers that lost their leading zeros.
            If text Like String$(Len(text), "#") And Len(text) < idLength Then
                FieldValue = String$(idLength - Len(text), "0") & text
            Else
                FieldValue = text
            End If
        Case Else
            Err.Raise vbObjectError + 601, "StudentRegistry.FieldValue", "Unknown field kind: " & kind
    End Select
End Function

Private Sub WriteField(ByVal target As Range, ByVal kind As FieldKind, ByVal value As Variant)
    Select Case kind
        Case fkIdText
            target.NumberFormat = TEXT_FORMAT
            target.Value = value
        Case fkIsoDate
            target.Value = value
            target.NumberFormat = DATE_FORMAT
        Case Else
            target.Value = value
    End Select
End Sub

Private Function ValuesEqual(ByVal cellValue As Variant, ByVal newValue As Variant) As Boolean
    If IsEmpty(newValue) Then
        ValuesEqual = False
        Exit Function
    End If

    If VarType(newValue) = vbDate Then newValue = CDbl(newValue)

    If IsNumberType(cellValue) And IsNumberType(newValue) Then
        ValuesEqual = (CDbl(cellValue) = CDbl(newValue))
    Else
        ValuesEqual = (Trim$(CStr(cellValue)) = Trim$(CStr(newValue)))
    End If
End Function

' Returns a Date, or Empty when the text is not a valid yyyy-mm-dd date.
Public Function ParseIsoDate(ByVal isoDate As String) As Variant
    Dim datePart As String
    Dim yearPart As Integer
    Dim monthPart As Integer
    Dim dayPart As Integer
    Dim parsed As Date

    ParseIsoDate = Empty
    If Len(isoDate) < ISO_DATE_LENGTH Then Exit Function

    datePart = Left$(isoDate, ISO_DATE_LENGTH)
    If Not datePart Like "####-##-##" Then Exit Function

    yearPart = CInt(Left$(datePart, 4))
    monthPart = CInt(Mid$(datePart, 6, 2))
    dayPart = CInt(Mid$(datePart, 9, 2))
    If monthPart < 1 Or monthPart > 12 Or dayPart < 1 Then Exit Function

    parsed = DateSerial(yearPart, monthPart, dayPart)
    ' DateSerial rolls invalid days over (2015-02-30 -> 2 March); reject those.
    If Day(parsed) <> dayPart Then Exit Function
    ParseIsoDate = parsed
End Function

Public Function ArabicGender(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "male", "m"
            ArabicGender = FromCodePoints("0630 0643 0631")
        Case "female", "f"
            ArabicGender = FromCodePoints("0623 0646 062B 0649")
        Case Else
            ArabicGender = value
    End Select
End Function

' ----------------------------------------------------------------------------
' Row state: status column + hidden "checked ID" column. The hidden column
' holds the ID the row was last checked for (found, not found or invalid);
' it is empty while a lookup is still pending.
' ----------------------------------------------------------------------------

Private Sub SetRowState(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal config As RegistryConfig, _
                        ByVal statusText As String, ByVal loadedId As String)
    ws.Cells(rowNumber, config.StatusColumn).Value = statusText
    WriteLoadedId ws, rowNumber, config, loadedId
End Sub

Private Sub WriteLoadedId(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                          ByVal config As RegistryConfig, ByVal loadedId As String)
    With ws.Cells(rowNumber, config.LoadedIdColumn)
        .NumberFormat = TEXT_FORMAT
        If Len(loadedId) = 0 Then .ClearContents Else .Value = loadedId
    End With
End Sub

Private Function LoadedIdOf(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                            ByVal config As RegistryConfig) As String
    LoadedIdOf = NormalizeIdValue(ws.Cells(rowNumber, config.LoadedIdColumn).Value2, config.IdLength)
End Function

Private Sub ClearRowState(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal config As RegistryConfig)
    ClearAutomaticFields ws, rowNumber
    ws.Cells(rowNumber, config.StatusColumn).ClearContents
    ws.Cells(rowNumber, config.LoadedIdColumn).ClearContents
End Sub

Private Sub ClearAutomaticFields(ByVal ws As Worksheet, ByVal rowNumber As Long)
    Dim fields() As FieldSpec
    Dim index As Long

    fields = StudentFields()
    For index = LBound(fields) To UBound(fields)
        ws.Cells(rowNumber, fields(index).ColumnLetter).ClearContents
    Next index
End Sub

Private Function HasAutomaticData(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim fields() As FieldSpec
    Dim index As Long

    fields = StudentFields()
    For index = LBound(fields) To UBound(fields)
        If Not IsEmpty(ws.Cells(rowNumber, fields(index).ColumnLetter).Value2) Then
            HasAutomaticData = True
            Exit Function
        End If
    Next index
End Function

' ----------------------------------------------------------------------------
' Ranges
' ----------------------------------------------------------------------------

' Limits the changed cells to the ID column inside the used area, so clearing
' a whole column does not loop over a million empty rows.
Public Function ChangedIdCells(ByVal ws As Worksheet, ByVal target As Range, _
                                ByVal config As RegistryConfig) As Range
    Dim lastRow As Long

    ' Not UsedRange: formatting whole columns stretches it to the last row.
    lastRow = LastFilledRow(ws, ID_COLUMN)
    If LastFilledRow(ws, config.StatusColumn) > lastRow Then lastRow = LastFilledRow(ws, config.StatusColumn)
    If LastFilledRow(ws, config.LoadedIdColumn) > lastRow Then lastRow = LastFilledRow(ws, config.LoadedIdColumn)
    If lastRow < config.FirstDataRow Then Exit Function

    Set ChangedIdCells = Intersect(target, _
        ws.Range(ws.Cells(config.FirstDataRow, ID_COLUMN), ws.Cells(lastRow, ID_COLUMN)))
End Function

Private Function LastFilledRow(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    LastFilledRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function

Private Function LastIdRow(ByVal ws As Worksheet, ByVal config As RegistryConfig) As Long
    LastIdRow = ws.Cells(ws.Rows.Count, ID_COLUMN).End(xlUp).Row
    If LastIdRow < config.FirstDataRow Then LastIdRow = config.FirstDataRow - 1
End Function

Private Function CellCount(ByVal rng As Range) As Long
    Dim area As Range
    For Each area In rng.Areas
        CellCount = CellCount + area.Cells.Count
    Next area
End Function

Private Function IsNumberType(ByVal value As Variant) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNumberType = True
    End Select
End Function

' ----------------------------------------------------------------------------
' Status texts (Arabic)
' ----------------------------------------------------------------------------

Public Function StatusHeader() As String
    StatusHeader = FromCodePoints("062D 0627 0644 0629 0020 0627 0644 062C 0644 0628")                         ' "Lookup status"
End Function

Public Function StatusFound() As String
    StatusFound = FromCodePoints("062A 0645 0020 062C 0644 0628 0020 0627 0644 0628 064A 0627 0646 0627 062A")                      ' "Data retrieved"
End Function

Public Function StatusFoundKeptManual(ByVal keptCount As Long) As String
    StatusFoundKeptManual = FromCodePoints("062A 0645 0020 0627 0644 062C 0644 0628 0020 002D 0020 0642 064A 0645 0020 064A 062F 0648 064A 0629 0020 0645 062D 0641 0648 0638 0629 003A 0020") & keptCount   ' "Retrieved - manual values kept: n"
End Function

Public Function StatusNotFound() As String
    StatusNotFound = FromCodePoints("063A 064A 0631 0020 0645 0648 062C 0648 062F 0020 0641 064A 0020 0627 0644 0633 062C 0644 0020 0627 0644 0645 062F 0646 064A")         ' "Not found in Civil Registry"
End Function

Public Function StatusInvalidId() As String
    StatusInvalidId = FromCodePoints("0631 0642 0645 0020 0647 0648 064A 0629 0020 063A 064A 0631 0020 0635 062D 064A 062D")                ' "Invalid ID number"
End Function

Public Function StatusPending() As String
    StatusPending = FromCodePoints("0628 0627 0646 062A 0638 0627 0631 0020 0625 0639 0627 062F 0629 0020 0627 0644 0645 062D 0627 0648 0644 0629")             ' "Waiting for retry"
End Function

Public Function StatusFailedKeptData() As String
    StatusFailedKeptData = FromCodePoints("062A 0639 0630 0631 0020 0627 0644 0627 062A 0635 0627 0644 0020 002D 0020 0627 0644 0628 064A 0627 0646 0627 062A 0020 0627 0644 0633 0627 0628 0642 0629 0020 0645 062D 0641 0648 0638 0629")   ' "No connection - previous data kept"
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim index As Long

    parts = Split(codePoints, " ")
    For index = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(index)))
    Next index
End Function

' ----------------------------------------------------------------------------
' Minimal JSON reader for the flat student object
' ----------------------------------------------------------------------------

Public Function JsonText(ByVal json As String, ByVal key As String) As String
    Dim raw As Variant
    raw = JsonValue(json, key)

    If IsEmpty(raw) Or IsNull(raw) Then
        JsonText = vbNullString
    Else
        JsonText = CStr(raw)
    End If
End Function

' Returns Empty when the key is missing, Null for JSON null, otherwise the value.
Private Function JsonValue(ByVal json As String, ByVal key As String) As Variant
    Dim valuePosition As Long
    Dim currentCharacter As String
    Dim result As String
    Dim escaped As Boolean

    valuePosition = JsonValueStart(json, key)
    If valuePosition = 0 Then
        JsonValue = Empty
        Exit Function
    End If

    If LCase$(Mid$(json, valuePosition, 4)) = "null" Then
        JsonValue = Null
        Exit Function
    End If

    If Mid$(json, valuePosition, 1) = Chr$(34) Then
        valuePosition = valuePosition + 1
        Do While valuePosition <= Len(json)
            currentCharacter = Mid$(json, valuePosition, 1)
            If escaped Then
                result = result & DecodeJsonEscape(json, valuePosition)
                escaped = False
            ElseIf currentCharacter = "\" Then
                escaped = True
            ElseIf currentCharacter = Chr$(34) Then
                JsonValue = result
                Exit Function
            Else
                result = result & currentCharacter
            End If
            valuePosition = valuePosition + 1
        Loop
        Err.Raise vbObjectError + 602, "StudentRegistry.JsonValue", "Unterminated string for key " & key
    End If

    Do While valuePosition <= Len(json)
        currentCharacter = Mid$(json, valuePosition, 1)
        Select Case currentCharacter
            Case ",", "}", " ", vbTab, vbCr, vbLf: Exit Do
        End Select
        result = result & currentCharacter
        valuePosition = valuePosition + 1
    Loop

    result = Trim$(result)
    Select Case LCase$(result)
        Case "true": JsonValue = True
        Case "false": JsonValue = False
        Case Else: JsonValue = result
    End Select
End Function

' Finds "key" used as a key (followed by a colon), not as a value.
Private Function JsonValueStart(ByVal json As String, ByVal key As String) As Long
    Dim quotedKey As String
    Dim searchFrom As Long
    Dim keyPosition As Long
    Dim position As Long

    quotedKey = Chr$(34) & key & Chr$(34)
    searchFrom = 1

    Do
        keyPosition = InStr(searchFrom, json, quotedKey, vbBinaryCompare)
        If keyPosition = 0 Then Exit Function

        position = SkipWhitespace(json, keyPosition + Len(quotedKey))
        If Mid$(json, position, 1) = ":" Then
            JsonValueStart = SkipWhitespace(json, position + 1)
            Exit Function
        End If
        searchFrom = keyPosition + 1
    Loop
End Function

Private Function SkipWhitespace(ByVal json As String, ByVal position As Long) As Long
    Do While position <= Len(json)
        Select Case Mid$(json, position, 1)
            Case " ", vbTab, vbCr, vbLf
                position = position + 1
            Case Else
                Exit Do
        End Select
    Loop
    SkipWhitespace = position
End Function

Private Function DecodeJsonEscape(ByVal json As String, ByRef position As Long) As String
    Select Case Mid$(json, position, 1)
        Case Chr$(34): DecodeJsonEscape = Chr$(34)
        Case "\": DecodeJsonEscape = "\"
        Case "/": DecodeJsonEscape = "/"
        Case "b": DecodeJsonEscape = Chr$(8)
        Case "f": DecodeJsonEscape = Chr$(12)
        Case "n": DecodeJsonEscape = vbLf
        Case "r": DecodeJsonEscape = vbCr
        Case "t": DecodeJsonEscape = vbTab
        Case "u"
            If position + 4 <= Len(json) Then
                DecodeJsonEscape = ChrW$(CLng("&H" & Mid$(json, position + 1, 4)))
                position = position + 4
            End If
        Case Else
            DecodeJsonEscape = Mid$(json, position, 1)
    End Select
End Function
