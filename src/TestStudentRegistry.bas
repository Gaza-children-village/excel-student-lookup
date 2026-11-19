Attribute VB_Name = "TestStudentRegistry"
Option Explicit

' Tests for StudentRegistry. Run RunStudentLookupTests (Alt+F8) in the development
' workbook; results go to the Immediate window (Ctrl+G).
' Uses a temporary sheet and FakeStudentApi: no network calls.
' Do NOT ship this module or FakeStudentApi in the secretaries' workbook.

' Fictional IDs (valid check digit unless named otherwise).
Private Const ID_A As String = "123456782"
Private Const ID_B As String = "400000006"
Private Const ID_C As String = "411111115"
Private Const ID_D As String = "402222228"
Private Const ID_BAD_CHECK As String = "123456789"
Private Const PARENT_ID_LEADING_ZERO As String = "087654323"

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunStudentLookupTests()
    Dim ws As Worksheet

    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "StudentRegistry tests, module " & MODULE_VERSION
    On Error GoTo Crash

    ' Pure functions
    Test_NormalizeId_PadsNumericIdsWithLeadingZeros
    Test_NormalizeId_KeepsTextIdsAndTrimsSpaces
    Test_ValidateId_ClassifiesEachKindOfProblem
    Test_ValidateId_SkipsCheckDigitWhenDisabled
    Test_CheckDigit_AcceptsValidAndRejectsInvalidIds
    Test_Json_ReadsNullsEscapesAndIgnoresKeyNamesUsedAsValues
    Test_ParseIsoDate_ReturnsEmptyForMalformedDates
    Test_ArabicGender_ConvertsKnownValuesOnly

    ' Sheet behaviour
    Set ws = CreateTestSheet()

    Test_Found_FillsFieldsStatusAndLoadedId ws
    Test_SameId_KeepsManualEditsAndFillsBlanks ws
    Test_NewId_ReplacesPreviousStudentData ws
    Test_ConnectionFailure_KeepsDataForSameStudent ws
    Test_ConnectionFailure_ClearsDataOfPreviousIdAndMarksPending ws
    Test_ConnectionFailure_StopsFurtherRequestsInBatch ws
    Test_HttpError_MarksPendingButContinuesBatch ws
    Test_NotFound_ClearsFields ws
    Test_ClearingId_ClearsRowState ws
    Test_InvalidId_IsNotSentToApi ws
    Test_NumericIdCell_IsRewrittenAsTextWithLeadingZeros ws
    Test_ParentId_KeepsLeadingZero ws
    Test_Retry_SelectsOnlyRowsNeedingLookup ws
    Test_Prepare_RecordsLoadedIdForExistingRows ws
    Test_ClearingWholeIdColumn_ProcessesOnlyFilledRows ws
    Test_Lookup_LeavesTemplateHealthAndNotesColumnsUntouched ws
    Test_Logging_RecordsRowResultsWithoutTheId ws
    Test_GcvWorkbook_UnmarkedSheetIsIgnored ws
    Test_GcvWorkbook_MarkedSheetIsFoundAfterRename ws
    Test_SelectedIdCells_TakesTheIdCellOfEverySelectedRow ws

Finish:
    GcvLogClient.CaptureLogForTests False
    DeleteTestSheet ws
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Student lookup tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

' ----------------------------------------------------------------------------
' Pure function tests
' ----------------------------------------------------------------------------

Private Sub Test_NormalizeId_PadsNumericIdsWithLeadingZeros()
    StartTest "NormalizeId pads numeric IDs with leading zeros"
    AssertEqual "087654323", NormalizeIdValue(87654323#, 9), "8-digit number"
    AssertEqual ID_A, NormalizeIdValue(123456782#, 9), "9-digit number"
    AssertEqual "", NormalizeIdValue(Empty, 9), "empty cell"
End Sub

Private Sub Test_NormalizeId_KeepsTextIdsAndTrimsSpaces()
    StartTest "NormalizeId keeps text IDs and trims spaces"
    AssertEqual "087654323", NormalizeIdValue(" 087654323 ", 9), "text with spaces"
    AssertEqual "12A", NormalizeIdValue("12A", 9), "non-digit text unchanged"
End Sub

Private Sub Test_ValidateId_ClassifiesEachKindOfProblem()
    StartTest "ValidateId classifies each kind of problem"
    AssertEqual ivEmpty, ValidateStudentId("", 9, True), "empty"
    AssertEqual ivNotDigits, ValidateStudentId("12345678X", 9, True), "letters"
    AssertEqual ivWrongLength, ValidateStudentId("12345678", 9, True), "8 digits"
    AssertEqual ivBadCheckDigit, ValidateStudentId(ID_BAD_CHECK, 9, True), "bad check digit"
    AssertEqual ivValid, ValidateStudentId(ID_A, 9, True), "valid"
End Sub

Private Sub Test_ValidateId_SkipsCheckDigitWhenDisabled()
    StartTest "ValidateId skips the check digit when disabled"
    AssertEqual ivValid, ValidateStudentId(ID_BAD_CHECK, 9, False), "check digit off"
End Sub

Private Sub Test_CheckDigit_AcceptsValidAndRejectsInvalidIds()
    StartTest "Check digit accepts valid IDs and rejects invalid ones"
    AssertEqual True, IsValidCheckDigit(ID_A), ID_A
    AssertEqual True, IsValidCheckDigit(ID_B), ID_B
    AssertEqual True, IsValidCheckDigit(PARENT_ID_LEADING_ZERO), PARENT_ID_LEADING_ZERO
    AssertEqual False, IsValidCheckDigit(ID_BAD_CHECK), ID_BAD_CHECK
    AssertEqual False, IsValidCheckDigit(""), "empty"
End Sub

Private Sub Test_Json_ReadsNullsEscapesAndIgnoresKeyNamesUsedAsValues()
    Dim json As String
    StartTest "JSON reader handles nulls, escapes and key names used as values"
    json = "{""note"": ""full_name"", ""full_name"" : ""Sami \""Abu\"" Ali"", " & _
           """english_name"": null, ""city"": ""Abc"", ""count"": 7}"
    AssertEqual "Sami ""Abu"" Ali", JsonText(json, "full_name"), "escaped quotes + spaced colon"
    AssertEqual "", JsonText(json, "english_name"), "null"
    AssertEqual "Abc", JsonText(json, "city"), "unicode escape"
    AssertEqual "7", JsonText(json, "count"), "number"
    AssertEqual "", JsonText(json, "missing"), "missing key"
    AssertEqual "2", JsonText("{" & vbCrLf & "  ""order"": 2" & vbCrLf & "}", "order"), "pretty-printed last value"
End Sub

Private Sub Test_ParseIsoDate_ReturnsEmptyForMalformedDates()
    StartTest "ParseIsoDate returns Empty for malformed dates"
    AssertEqual CDbl(DateSerial(2015, 3, 10)), CDbl(ParseIsoDate("2015-03-10")), "date"
    AssertEqual CDbl(DateSerial(2015, 3, 10)), CDbl(ParseIsoDate("2015-03-10T00:00:00")), "datetime"
    AssertEqual True, IsEmpty(ParseIsoDate("10/03/2015")), "wrong format"
    AssertEqual True, IsEmpty(ParseIsoDate("")), "empty"
    AssertEqual True, IsEmpty(ParseIsoDate("2015-02-30")), "impossible day"
    AssertEqual True, IsEmpty(ParseIsoDate("2015-00-10")), "month zero"
End Sub

Private Sub Test_ArabicGender_ConvertsKnownValuesOnly()
    StartTest "ArabicGender converts known values only"
    AssertEqual &H630, AscW(ArabicGender("male")), "male starts with Arabic dhal"
    AssertEqual &H623, AscW(ArabicGender("F")), "female starts with Arabic alef hamza"
    AssertEqual "x", ArabicGender("x"), "unknown unchanged"
End Sub

' ----------------------------------------------------------------------------
' Sheet behaviour tests
' ----------------------------------------------------------------------------

Private Sub Test_Found_FillsFieldsStatusAndLoadedId(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Found ID fills fields, status and loaded ID"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "Student A En")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api

    AssertEqual "Student A", ws.Range("C4").Value2, "Arabic name column"
    AssertEqual "Student A En", ws.Range("D4").Value2, "English name"
    AssertEqual CDbl(DateSerial(2015, 3, 10)), ws.Range("E4").Value2, "birthdate"
    AssertEqual 5, ws.Range("T4").Value2, "family members"
    AssertEqual StatusFound(), ws.Range("AE4").Value2, "status"
    AssertEqual ID_A, ws.Range("AF4").Value2, "loaded ID"
End Sub

Private Sub Test_SameId_KeepsManualEditsAndFillsBlanks(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Re-entering the same ID keeps manual edits and fills blanks"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    ws.Range("D4").Value = "Manual English"   ' secretary adds a spelling
    ws.Range("M4").ClearContents              ' a blank to be refilled

    Process ws, ws.Range("B4"), api

    AssertEqual "Manual English", ws.Range("D4").Value2, "manual value kept"
    AssertEqual "Father A", ws.Range("M4").Value2, "blank refilled"
    AssertEqual StatusFoundKeptManual(1), ws.Range("AE4").Value2, "status counts kept value"
End Sub

Private Sub Test_NewId_ReplacesPreviousStudentData(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "A new ID replaces the previous student's data"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "Manual-looking")
    api.RespondFound ID_B, StudentJson("Student B", "")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    ws.Range("B4").Value = ID_B
    Process ws, ws.Range("B4"), api

    AssertEqual "Student B", ws.Range("C4").Value2, "name replaced"
    AssertEqual True, IsEmpty(ws.Range("D4").Value2), "old English name removed"
    AssertEqual ID_B, ws.Range("AF4").Value2, "loaded ID updated"
End Sub

Private Sub Test_ConnectionFailure_KeepsDataForSameStudent(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Connection failure keeps data for the same student"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    api.FailAllConnections = True
    Process ws, ws.Range("B4"), api

    AssertEqual "Student A", ws.Range("C4").Value2, "data kept"
    AssertEqual StatusFailedKeptData(), ws.Range("AE4").Value2, "status"
    AssertEqual ID_A, ws.Range("AF4").Value2, "loaded ID kept"
End Sub

Private Sub Test_ConnectionFailure_ClearsDataOfPreviousIdAndMarksPending(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Connection failure clears the previous ID's data and marks the row pending"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    api.FailAllConnections = True
    ws.Range("B4").Value = ID_B
    Process ws, ws.Range("B4"), api

    AssertEqual True, IsEmpty(ws.Range("C4").Value2), "previous student's name removed"
    AssertEqual StatusPending(), ws.Range("AE4").Value2, "status pending"
    AssertEqual True, IsEmpty(ws.Range("AF4").Value2), "no loaded ID"
End Sub

Private Sub Test_ConnectionFailure_StopsFurtherRequestsInBatch(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    Dim summary As BatchSummary
    StartTest "Connection failure stops further requests in the batch"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.FailAllConnections = True

    ws.Range("B4:B6").NumberFormat = "@"
    ws.Range("B4").Value = ID_A
    ws.Range("B5").Value = ID_B
    ws.Range("B6").Value = ID_C
    Set summary = Process(ws, ws.Range("B4:B6"), api)

    AssertEqual 1, api.CallCount, "only one request sent"
    AssertEqual 3, summary.Pending, "all rows pending"
    AssertEqual True, summary.ConnectionDown, "connection flagged"
    AssertEqual StatusPending(), ws.Range("AE6").Value2, "last row status"
End Sub

Private Sub Test_HttpError_MarksPendingButContinuesBatch(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    Dim summary As BatchSummary
    StartTest "HTTP error marks the row pending but the batch continues"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondHttpError ID_A, 500
    api.RespondFound ID_B, StudentJson("Student B", "")

    ws.Range("B4:B5").NumberFormat = "@"
    ws.Range("B4").Value = ID_A
    ws.Range("B5").Value = ID_B
    Set summary = Process(ws, ws.Range("B4:B5"), api)

    AssertEqual 2, api.CallCount, "both requested"
    AssertEqual StatusPending() & " (HTTP 500)", ws.Range("AE4").Value2, "error row status"
    AssertEqual "Student B", ws.Range("C5").Value2, "next row filled"
    AssertEqual 500, summary.LastHttpStatus, "status recorded"
End Sub

Private Sub Test_NotFound_ClearsFields(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Not found clears the automatic fields"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    ws.Range("B4").Value = ID_D          ' not registered in the fake
    Process ws, ws.Range("B4"), api

    AssertEqual True, IsEmpty(ws.Range("C4").Value2), "fields cleared"
    AssertEqual StatusNotFound(), ws.Range("AE4").Value2, "status"
    AssertEqual ID_D, ws.Range("AF4").Value2, "checked ID recorded"
End Sub

Private Sub Test_ClearingId_ClearsRowState(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Clearing the ID clears fields, status and loaded ID"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")
    ws.Range("F4").Value = "manual column"

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    ws.Range("B4").ClearContents
    Process ws, ws.Range("B4"), api

    AssertEqual True, IsEmpty(ws.Range("C4").Value2), "fields cleared"
    AssertEqual True, IsEmpty(ws.Range("AE4").Value2), "status cleared"
    AssertEqual True, IsEmpty(ws.Range("AF4").Value2), "loaded ID cleared"
    AssertEqual "manual column", ws.Range("F4").Value2, "manual column untouched"
End Sub

Private Sub Test_InvalidId_IsNotSentToApi(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "An invalid ID is not sent to the API"
    ResetSheet ws
    Set api = New FakeStudentApi

    ws.Range("B4").Value = ID_BAD_CHECK
    Process ws, ws.Range("B4"), api

    AssertEqual 0, api.CallCount, "no request"
    AssertEqual StatusInvalidId(), ws.Range("AE4").Value2, "status"
    AssertEqual ID_BAD_CHECK, ws.Range("AF4").Value2, "checked ID recorded"
End Sub

Private Sub Test_NumericIdCell_IsRewrittenAsTextWithLeadingZeros(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "A numeric ID cell is rewritten as text with leading zeros"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound PARENT_ID_LEADING_ZERO, StudentJson("Student Z", "")

    ws.Range("B4").NumberFormat = "General"
    ws.Range("B4").Value = 87654323#
    Process ws, ws.Range("B4"), api

    AssertEqual PARENT_ID_LEADING_ZERO, ws.Range("B4").Value2, "ID rewritten"
    AssertEqual "Student Z", ws.Range("C4").Value2, "lookup used padded ID"
End Sub

Private Sub Test_ParentId_KeepsLeadingZero(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Parent IDs keep their leading zero"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "", "87654323")

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api

    AssertEqual PARENT_ID_LEADING_ZERO, ws.Range("N4").Value2, "father ID"
End Sub

Private Sub Test_Retry_SelectsOnlyRowsNeedingLookup(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    Dim pending As Range
    StartTest "Retry selects only rows that need a lookup"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")
    api.RespondFound ID_C, StudentJson("Student C", "")

    ws.Range("B4:B6").NumberFormat = "@"
    ws.Range("B4").Value = ID_A
    ws.Range("B6").Value = ID_C
    Process ws, ws.Range("B4"), api           ' row 4 found
    Process ws, ws.Range("B6"), api           ' row 6 found
    api.FailAllConnections = True
    Process ws, ws.Range("B6"), api           ' row 6: failed, data kept
    ws.Range("B5").Value = ID_B               ' row 5: never looked up
    api.FailAllConnections = False
    ws.Range("B7").Value = ID_D               ' row 7: not found -> skipped
    Process ws, ws.Range("B7"), api
    ws.Range("B8").Value = ID_BAD_CHECK       ' row 8: invalid -> skipped
    Process ws, ws.Range("B8"), api

    Set pending = PendingIdCells(ws, TestConfig())

    AssertEqual "$B$5:$B$6", pending.Address, "rows 5 and 6 only"
End Sub

Private Sub Test_Prepare_RecordsLoadedIdForExistingRows(ByVal ws As Worksheet)
    StartTest "Preparing the sheet records loaded IDs for rows that already have data"
    ResetSheet ws
    ws.Range("B4").Value = ID_A
    ws.Range("C4").Value = "Existing name"
    ws.Range("B5").Value = ID_B           ' ID only, no data yet

    PrepareStudentSheet ws, TestConfig()

    AssertEqual ID_A, ws.Range("AF4").Value2, "existing row recorded"
    AssertEqual True, IsEmpty(ws.Range("AF5").Value2), "row without data not recorded"
    AssertEqual StatusHeader(), ws.Range("AE3").Value2, "status header"
    AssertEqual True, ws.Columns("AF").Hidden, "loaded-ID column hidden"
End Sub

Private Sub Test_ClearingWholeIdColumn_ProcessesOnlyFilledRows(ByVal ws As Worksheet)
    Dim config As RegistryConfig
    Dim api As IStudentApi
    Dim summary As BatchSummary
    StartTest "Clearing the whole ID column processes only rows that hold data"
    ResetSheet ws
    Set config = TestConfig()
    Set api = New FakeStudentApi
    ws.Range("B4").Value = ID_A
    ws.Range("C4").Value = "Existing name"
    PrepareStudentSheet ws, config          ' formats B and W down to the last row

    ws.Columns("B").ClearContents
    Set summary = ProcessIdCells(ws, ChangedIdCells(ws, ws.Columns("B"), config), api, config)

    AssertEqual 1, summary.Total, "one row processed"
    AssertEqual True, IsEmpty(ws.Range("C4").Value2), "row cleared"
End Sub

Private Sub Test_Lookup_LeavesTemplateHealthAndNotesColumnsUntouched(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    StartTest "Lookup leaves the template's health (V:AC) and notes (AD) columns untouched"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")
    ws.Range("V4").Value = True                 ' chronic illness: yes
    ws.Range("W4").Value = "Health detail"
    ws.Range("AD4").Value = "Note"

    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api

    AssertEqual True, ws.Range("V4").Value2, "health flag kept"
    AssertEqual "Health detail", ws.Range("W4").Value2, "health detail kept"
    AssertEqual "Note", ws.Range("AD4").Value2, "notes kept"
    AssertEqual StatusFound(), ws.Range("AE4").Value2, "status after the template columns"
End Sub

Private Sub Test_Logging_RecordsRowResultsWithoutTheId(ByVal ws As Worksheet)
    Dim api As FakeStudentApi
    Dim content As String
    StartTest "Logging records row results but never the ID"
    ResetSheet ws
    Set api = New FakeStudentApi
    api.RespondFound ID_A, StudentJson("Student A", "")
    ' The lines that would go to GCV Logger are kept in memory (all levels).
    GcvLogClient.CaptureLogForTests True
    ws.Range("B4").Value = ID_A
    Process ws, ws.Range("B4"), api
    content = GcvLogClient.CapturedLog()
    GcvLogClient.CaptureLogForTests False

    AssertEqual True, InStr(content, "[student-lookup] - DEBUG:") > 0, "debug line written"
    AssertEqual True, InStr(content, "Row 4: found") > 0, "row result"
    AssertEqual 0, InStr(content, ID_A), "no ID in the log"
    AssertEqual 0, InStr(content, "Student A"), "no name in the log"
End Sub

Private Sub Test_GcvWorkbook_UnmarkedSheetIsIgnored(ByVal ws As Worksheet)
    StartTest "A sheet without the GCV marker is ignored"
    AssertEqual False, GcvWorkbook.IsStudentSheet(ws), "not a student sheet"
    AssertEqual False, GcvWorkbook.IsGcvWorkbook(ws.Parent), "not a GCV workbook"
    AssertEqual False, GcvWorkbook.IsStudentSheet(Nothing), "Nothing"
End Sub

Private Sub Test_GcvWorkbook_MarkedSheetIsFoundAfterRename(ByVal ws As Worksheet)
    Dim other As Worksheet
    StartTest "The marked student sheet is recognised, also after renaming"
    Set other = ws.Parent.Worksheets.Add(After:=ws)
    GcvWorkbook.MarkStudentSheet ws
    AssertEqual True, GcvWorkbook.IsStudentSheet(ws), "marked sheet"
    AssertEqual False, GcvWorkbook.IsStudentSheet(other), "other sheet"
    ws.Name = "zz_renamed"
    AssertEqual True, GcvWorkbook.IsStudentSheet(ws), "after rename"
    AssertEqual GcvWorkbook.TEMPLATE_TEXT, Replace(Replace(ws.Parent.Names(GcvWorkbook.TEMPLATE_NAME).RefersTo, "=", ""), """", ""), "template marker"
    Application.DisplayAlerts = False
    other.Delete
    Application.DisplayAlerts = True
End Sub

Private Sub Test_SelectedIdCells_TakesTheIdCellOfEverySelectedRow(ByVal ws As Worksheet)
    Dim config As RegistryConfig
    Dim cells As Range
    StartTest "Fetch selected rows: the ID cell of each selected data row"
    ResetSheet ws
    Set config = TestConfig()
    ws.Range("B4").Value = ID_A
    ws.Range("B5").Value = ID_B
    ws.Range("B6").Value = ID_C
    Set cells = SelectedIdCells(ws, ws.Range("D5:F6"), config)
    AssertEqual "$B$5:$B$6", cells.Address, "rows 5 and 6"
    Set cells = SelectedIdCells(ws, ws.Range("A1:A2"), config)
    AssertEqual True, cells Is Nothing, "header rows ignored"
    Set cells = SelectedIdCells(ws, ws.Columns("C"), config)
    AssertEqual "$B$4:$B$6", cells.Address, "whole column limited to the data"
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function Process(ByVal ws As Worksheet, ByVal idCells As Range, _
                         ByVal fake As FakeStudentApi) As BatchSummary
    Dim api As IStudentApi
    Set api = fake
    Set Process = ProcessIdCells(ws, idCells, api, TestConfig())
End Function

Private Function TestConfig() As RegistryConfig
    Dim config As RegistryConfig
    Set config = New RegistryConfig    ' defaults: 9 digits, check digit on, AE/AF, row 4
    Set TestConfig = config
End Function

Private Function StudentJson(ByVal fullName As String, ByVal englishName As String, _
                             Optional ByVal fatherId As String = "400000006") As String
    Dim englishValue As String
    If Len(englishName) = 0 Then englishValue = "null" Else englishValue = """" & englishName & """"

    StudentJson = "{""full_name"":""" & fullName & """," & _
                  """english_name"":" & englishValue & "," & _
                  """birthdate"":""2015-03-10""," & _
                  """gender"":""male""," & _
                  """father_name"":""Father A""," & _
                  """father_id"":" & fatherId & "," & _
                  """mother_name"":""Mother A""," & _
                  """mother_id"":""402222228""," & _
                  """place_before_displacement"":""Place A""," & _
                  """family_members_count"":5," & _
                  """order_among_siblings"":2}"
End Function

' The tests run inside the add-in, so they use their own temporary workbook
' (closed without saving at the end). It is not a GCV workbook, so the
' Excel-wide events ignore the test writes.
Private Function CreateTestSheet() As Worksheet
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "zz_tests"
    Set CreateTestSheet = wb.Worksheets(1)
End Function

Private Sub DeleteTestSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Parent.Close SaveChanges:=False
End Sub

Private Sub ResetSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Columns.Hidden = False
    ws.Range("B4:B20").NumberFormat = "@"
End Sub

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean

    If IsEmpty(expected) Or IsEmpty(actual) Then
        same = (IsEmpty(expected) And IsEmpty(actual))
    ElseIf IsNumeric(expected) And IsNumeric(actual) _
           And VarType(expected) <> vbString And VarType(actual) <> vbString Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If

    If same Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & _
                    " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub
