# Daily Data Records (DDR)

Desktop app for logging daily service-counter transactions. Each entry is saved
in the app and mirrored into an Excel workbook at
`Documents\DDR\service_entries_database.xlsx` (installed app) or
`Documents\DDR-dev\...` (running from source via `npm start`).

## How your data is protected

| Layer | What it protects | Where |
|---|---|---|
| Archive (`entries_archive.json`) | Every entry, permanently, the moment it's added or edited | `%APPDATA%\DDR\` |
| Local daily backup | The whole Excel workbook, including sheets you add yourself (`D_repo`, `Guidelines`, etc.) | `Documents\DDR\Backups\service_entries_database_backup.xlsx` |
| Cloud daily backup | Same as above, off-site | `G:\My Drive\DDR Backups\service_entries_database_backup.xlsx` (only if Google Drive for Desktop is signed in and running) |

Both backups are a single file, refreshed at most once a day (the first time
DDR writes the workbook that day), holding whatever the workbook looked like
*before* that day's changes.

## Incident: Excel file corrupted to 0 bytes (Sep 12, 2026)

### Problem
`service_entries_database.xlsx` was found truncated to 0 bytes. The `Entries`
sheet is backed by the archive JSON and was fully recoverable, but the sheets
that live only in the Excel file (`D_repo`, `Guidelines`, `Pending`, `Links`)
had no backup at the time, and data entered directly into `D_repo` between
Sep 9–12 could not be recovered from any source (no later AutoRecover
snapshot, no shadow copies, no OneDrive version history existed for this
file).

### Cause (confirmed via Windows Event Log)
The old code wrote directly onto the live file
(`workbook.xlsx.writeFile(EXCEL_FILE)`), with no atomicity — any interruption
mid-write leaves the file in whatever state that write had reached.

The Windows System event log pinpoints the exact interruption on Sep 12:

| Time | Event |
|---|---|
| 16:59:17 | `service_entries_database.xlsx` last modified (the write that got corrupted) |
| 16:59:34 | A process called `SetSuspendState` / `SetSystemPowerState` |
| 16:59:35 | **System entered sleep** |
| 16:59:37 | **System resumed** (~2 seconds later) |

The PC went to sleep in the middle of DDR's write to the workbook and resumed
before it finished, leaving the file truncated. (OneDrive Known Folder Move
and Excel AutoSave were both investigated and ruled out — neither is active
on this machine.)

### Solution (`main.js`)
1. **Atomic writes** — `atomicWriteWorkbook()` writes to a temp file in the
   same folder and only `fs.renameSync`s it over the real workbook after the
   write fully succeeds. A sleep/crash/kill/power-loss mid-write can now only
   ever leave a stray `.tmp-*` file (auto-cleaned on the next write) — the
   real workbook is always either the old good copy or the fully-written new
   one, never truncated.
2. **Self-heal** — `excelFileMissingOrEmpty()` treats a 0-byte file the same
   as "no file yet," so if this ever happens again DDR rebuilds the `Entries`
   sheet from the archive automatically instead of refusing to touch the file
   forever.
3. **Daily backups** — `backupExcelFileIfDue()` (local) and
   `backupToGoogleDriveIfDue()` (cloud) snapshot the workbook once a day
   before the day's first write, specifically so sheets not covered by the
   archive JSON have a real recovery point. Both are best-effort and never
   block the real write if the backup destination is unavailable.

### Recovery performed
The Excel file was rebuilt from an Excel AutoRecover snapshot
(`%APPDATA%\Roaming\Microsoft\Excel\service_entries_database (version 1).xlsb`,
Sep 8 7:14 PM) for the sheet structure, with `Entries` rebuilt fresh from the
158-row archive on top of it. 62 additional entries that existed only in a
later cached copy of the Excel file (dated through Sep 14) were merged back
into the archive so they're protected going forward. `D_repo` data for
Sep 9–12 was not recoverable from any source and would need to be re-entered
manually if available elsewhere.
