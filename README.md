# GCV Student Lookup

![Release](https://img.shields.io/github/v/release/Gaza-children-village/excel-student-lookup?filter=student-lookup-v*&label=release)
![Excel](https://img.shields.io/badge/Excel-2016%2B%20Windows-217346)
![VBA](https://img.shields.io/badge/VBA-add--in-blue)

Excel add-in (`GCV Student Lookup.xlam`) that fills the student rows of the GCV Administration workbook from the Civil Registry served by the local GCV application.

## Ribbon

| Button | What it does |
|---|---|
| **جلب بيانات الصفوف المحددة** (fetch selected rows) | Looks up the selected rows; refreshes filled rows without overwriting values corrected by hand |
| **جلب بيانات الصفوف الجديدة** (fetch new rows) | Every row whose ID has no data yet, has changed, or failed before because of the connection |

The buttons work on the student sheet of a GCV workbook; on any other sheet they say which sheet to open.

**On demand only.** Nothing happens while the secretary types, so **Ctrl+Z** keeps working and editing or saving never sends requests.

## How it works

- The workbook is a plain `.xlsx` without code. Two hidden names identify it: `GCV_TEMPLATE` and `GCV_STUDENT_SHEET` (points to the student sheet and follows renames). `SetupStudentSheet` adds them to older copies.
- Request: `GET http://127.0.0.1:5555/api/student/<id>` (changeable in the workbook's hidden Settings sheet, names `cfg_*`).
- IDs: 9 digits with a check digit; kept as text so leading zeros stay.
- Status column **AE** shows the result of each row in Arabic; hidden column **AF** remembers which ID the row's data belongs to.
- No data loss: when the service cannot be reached, a row keeps the data of the same student and is marked "waiting for retry"; one summary message for the whole batch.
- Logs (through GCV Logger) contain row numbers and counts, never IDs or names.

## GCV application (local API)

- Requests carry `X-GCV-Excel-Addin: gcv/student-lookup/<version>` and, once the app has written one, its key `X-GCV-Api-Key` (read from `HKCU\Software\GCV Tools\LocalApiKey`, never logged). The app requires the key only after all PCs have 3.1.0 or later.
- `401`: the app refuses the key → the batch stops, rows wait for retry, the message says to open or update the GCV application.
- `503`: the registry is not ready (not installed yet, locked, or out of date) → the batch stops, rows wait for retry.
- After a failed batch, one `GET /api/health` adds the app version and the registry state to the message.
- Contract: `docs/local-api.md` in the GCV desktop app repository.

## Requirements

- The GCV application running on the same PC (Civil Registry API on port 5555).
- GCV Logger (optional but recommended).

## Installing

Secretaries do not install add-ins one by one. The **GCV Tools installer** (see [excel-updater](https://github.com/Gaza-children-village/excel-updater#installing-on-a-pc)) installs all of them for the current Windows user, and the **update tools** button keeps them current.

## Building from source

Excel cannot load a VBA project created outside Excel, so the add-in is assembled in Excel from the files in this repository:

1. Copy `GCV Student Lookup - shell.xlam` and rename the copy to `GCV Student Lookup.xlam` (the shell already contains the add-in file type and the ribbon).
2. Open it, press **Alt+F11** and select its VBAProject.
3. **File › Import File**: every file in `src/`.
4. Double-click **ThisWorkbook** and paste the contents of `ThisWorkbook.code`.
5. **Debug › Compile VBAProject**, then save with **Ctrl+S**. Saving records the add-in's id, version and file name in the file.

Source files are ASCII with CRLF line endings. Arabic text is stored as Unicode code points (`FromCodePoints("0627 ...")`), because the VBA editor cannot keep Arabic characters.

## Tests

Run `RunStudentLookupTests` from the VBA editor (click inside the procedure and press **F5**). Results appear in a message box and in the Immediate window (**Ctrl+G**). The tests use fakes and temporary files: no network, no real data.

Before a release, remove the development modules (`TestStudentRegistry`, `FakeStudentApi`), compile and save.

## Releasing

Releases are checked and published by GitHub Actions ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)); only the build happens by hand, because it needs Excel.

1. Raise `ADDIN_VERSION` in `src/GcvAddin.bas` and add a `## <version>` section to `CHANGELOG.md`.
2. Build the add-in in Excel (see above) **without** the test modules, and save it.
3. Upload it as a **draft** release:
   ```
   gh release create student-lookup-v<version> "GCV Student Lookup.xlam" --draft --repo Gaza-children-village/excel-student-lookup
   ```
4. Merge to `main` and push the tag `student-lookup-v<version>` (through Gitea; the mirror carries it to GitHub).
5. The workflow checks the uploaded file (same id and version as the tag, add-in type, ribbon, no test modules, CHANGELOG section), adds `manifest.json` (SHA-256 and size), renames the file to its asset name, sets the release notes from `CHANGELOG.md` and **publishes**. If anything is wrong, the release stays a draft (the updater ignores drafts) and the job shows why.

Uploaded the draft after pushing the tag? Re-run the failed **release** job in the Actions tab, or `gh workflow run ci.yml -f tag=student-lookup-v<version> --repo Gaza-children-village/excel-student-lookup`.

Every push also runs the source checks: VBA files ASCII with CRLF, module names equal to file names, `Option Explicit`, the shared ribbon tab, a valid identity in `GcvAddin.bas` with a CHANGELOG section. Run them locally with `python .github/scripts/gcv_ci.py check`.

## Files

| File | Purpose |
|---|---|
| `src/StudentRegistry.bas` | Lookup logic: validation, writing rows, status, summary |
| `src/LookupUi.bas` | Ribbon and messages |
| `src/GcvWorkbook.bas` | Recognising GCV workbooks and the student sheet |
| `src/HttpStudentApi.cls`, `src/IStudentApi.cls` | Registry API (WinHTTP) behind an interface |
| `src/RegistryConfig.cls`, `src/BatchSummary.cls`, `src/StudentLookupResult.cls` | Settings, batch counts, one result |
| `src/GcvLocalApi.bas` | Headers (add-in, key) and health check for the GCV app; same file in English Names |
| `src/GcvLogClient.bas`, `src/GcvAddin.bas` | Logging through GCV Logger; identity of the add-in |
| `src/TestStudentRegistry.bas`, `src/FakeStudentApi.cls`, `src/TestGcvLocalApi.bas`, `src/FakeHeaderSink.cls` | Tests (development only) |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog and loader |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |

See [CHANGELOG.md](CHANGELOG.md).
