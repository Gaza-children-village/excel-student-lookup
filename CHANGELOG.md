# Changelog - gcv/student-lookup (GCV Student Lookup add-in)

## 3.1.0

- Ready for the local API of the GCV desktop app (v2.0.0, which also installs and updates the add-ins): every request sends `X-GCV-Excel-Addin: gcv/student-lookup/<version>` (replaces the old `X-GCV-Excel-Module-Version` header, which still said 3.0.0) and, when the app has written one, its key `X-GCV-Api-Key` (from `HKCU\Software\GCV Tools\LocalApiKey`; never logged).
- 401 (key refused): the batch stops, rows wait for retry, and the message says to open or update the GCV application.
- 503 (registry not ready): the batch stops, rows wait for retry, and the message says the registry is not ready yet.
- After a failed batch the add-in asks `GET /api/health` once and adds the app version to the message.
- The check digit rule is verified (60,000 IDs of a 2023 registry sample); no longer marked "to verify".

## 3.0.1

- The two buttons are always active: on another sheet they say to open the student sheet. Before, they could stay greyed out on the student sheet because the ribbon was not refreshed (add-in loaded after the workbook, or after a VBA reset).

## 3.0.0

- Separate Excel add-in "GCV Student Lookup.xlam"; the school workbook is a plain .xlsx without code.
- On demand only: ribbon buttons "fetch selected rows" and "fetch new rows". Typing an ID no longer starts a lookup, so Ctrl+Z keeps working and editing or saving never sends requests.
- Works on the student sheet of GCV workbooks (hidden name GCV_STUDENT_SHEET, follows renames); buttons are greyed out elsewhere.
- Messages in Arabic.

## 1.2.0

- Logging through gcv/logger; the request address (it contains the ID) is no longer printed.

## 1.1.0

- No data loss when the connection fails; one summary message for many rows; manual corrections kept; Arabic status column; settings sheet; ID check digit; leading zeros kept.
