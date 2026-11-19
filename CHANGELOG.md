# Changelog - gcv/student-lookup (GCV Student Lookup add-in)

## 3.0.0

- Separate Excel add-in "GCV Student Lookup.xlam"; the school workbook is a plain .xlsx without code.
- On demand only: ribbon buttons "fetch selected rows" and "fetch new rows". Typing an ID no longer starts a lookup, so Ctrl+Z keeps working and editing or saving never sends requests.
- Works on the student sheet of GCV workbooks (hidden name GCV_STUDENT_SHEET, follows renames); buttons are greyed out elsewhere.
- Messages in Arabic.

## 1.2.0

- Logging through gcv/logger; the request address (it contains the ID) is no longer printed.

## 1.1.0

- No data loss when the connection fails; one summary message for many rows; manual corrections kept; Arabic status column; settings sheet; ID check digit; leading zeros kept.
