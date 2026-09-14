const { app, BrowserWindow, ipcMain, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const DATA_FILE = path.join(app.getPath('userData'), 'entries.json');
const ARCHIVE_FILE = path.join(app.getPath('userData'), 'entries_archive.json');
// Rows the user deleted from the working list + Excel. The archive itself is
// never trimmed (append-only permanent record) - this is just a "don't show
// these in Excel any more" list, so a deleted row can still be recovered from
// the archive by hand if it's ever needed.
const DELETED_FILE = path.join(app.getPath('userData'), 'deleted_rows.json');
// Installed app writes to Documents\DDR; a source/dev run (npm start) writes to a
// separate folder so the two never fight over the same workbook.
const DOCS_DIR = path.join(app.getPath('documents'), app.isPackaged ? 'DDR' : 'DDR-dev');
const EXCEL_FILE = path.join(DOCS_DIR, 'service_entries_database.xlsx');
const EXCEL_SHEET = 'Entries';
const BACKUPS_DIR = path.join(DOCS_DIR, 'Backups');
const BACKUP_FILE = path.join(BACKUPS_DIR, 'service_entries_database_backup.xlsx');
// Off-site copy: Google Drive is mounted as G:\ (Drive File Stream) on this PC.
// Only ever read from here to check it's mounted - never rely on it being there.
const GDRIVE_ROOT = 'G:\\My Drive';
const GDRIVE_BACKUPS_DIR = path.join(GDRIVE_ROOT, 'DDR Backups');
const GDRIVE_BACKUP_FILE = path.join(GDRIVE_BACKUPS_DIR, 'service_entries_database_backup.xlsx');

const HEADERS =['Date','Time','Name','Mobile','Type','Source','Method','Location','Main Service','Sub Service',
  'Charge','Payment Mode','Pending','Pending Rec.','Deductions','Deduction Amount','Commission','Ref','Remarks'];

function readJSON(file, fallback) {
  try {
    if (fs.existsSync(file)) {
      return JSON.parse(fs.readFileSync(file, 'utf-8'));
    }
  } catch (e) { /* fall through to default */ }
  return fallback;
}

function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data), 'utf-8');
}

function ensureDocsDir() {
  if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR, { recursive: true });
}

// --- Excel mirror (ExcelJS) --------------------------------------------------
// ExcelJS round-trips other worksheets with their styling, number formats,
// column widths, merged cells, data validation and formulas intact. Charts and
// embedded images are still not preserved by any free library.

// Blank fields become a real empty-string cell (not a missing cell), so every
// data row spans the full HEADERS width instead of stopping at its last value.
function cellValue(v) {
  return (v === undefined || v === null) ? '' : v;
}

// Give existing data rows a cell in every column (fills gaps left by rows that
// were written before this — only adds empty cells, never changes a value).
function fillRowGaps(ws) {
  let changed = false;
  const N = HEADERS.length;
  for (let rn = 2; rn <= ws.rowCount; rn++) {
    const row = ws.getRow(rn);
    if (row.actualCellCount === 0) continue; // leave entirely-empty rows alone
    for (let c = 1; c <= N; c++) {
      const cell = row.getCell(c);
      if (cell.type === ExcelJS.ValueType.Null) { cell.value = ''; changed = true; }
    }
  }
  return changed;
}

function cellToPlain(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return v.result == null ? '' : v.result; // formula -> cached value
    if ('text' in v) return v.text;                              // hyperlink
    if (v instanceof Date) return v.toISOString();
    return '';
  }
  return v;
}

function normCell(v) { return v == null ? '' : v; }
function rowKey(row, n) {
  const out = [];
  for (let i = 0; i < n; i++) out.push(normCell(row[i]));
  return JSON.stringify(out);
}

function readDeletedKeys() {
  return new Set(readJSON(DELETED_FILE, []));
}
function addDeletedKey(key) {
  const keys = readJSON(DELETED_FILE, []);
  if (!keys.includes(key)) { keys.push(key); writeJSON(DELETED_FILE, keys); }
}

// Map every data row currently in the sheet, by content, to the row number(s)
// it's on (usually one, but a Map of arrays copes if a row is ever duplicated).
function sheetRowIndex(ws) {
  const N = HEADERS.length;
  const map = new Map();
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (row.actualCellCount === 0) continue;
    const vals = row.values;
    const arr = [];
    for (let c = 1; c <= N; c++) arr.push(cellToPlain(vals[c]));
    const k = rowKey(arr, N);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  }
  return map;
}

async function openWorkbook() {
  if (!fs.existsSync(EXCEL_FILE)) return null;
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_FILE);
    return wb;
  } catch (e) {
    return null; // locked / unreadable -> caller decides
  }
}

// A 0-byte file (an interrupted write got killed before it wrote anything)
// has nothing in it to protect, so treat it the same as "no file yet" and let
// DDR rebuild the Entries sheet from the archive instead of refusing forever.
// Anything else that exists but can't be read (locked, genuinely corrupt) is
// left alone - see the "unreadable" guard below.
function excelFileMissingOrEmpty() {
  try {
    if (!fs.existsSync(EXCEL_FILE)) return true;
    return fs.statSync(EXCEL_FILE).size === 0;
  } catch (e) {
    return true;
  }
}

// Keep exactly ONE local backup file, refreshed at most once per day (first
// write of the day) - never a growing list of dated copies. The archive JSON
// only protects the Entries data - custom sheets (D_repo, Guidelines, ...)
// live nowhere but this file, so this is their only safety net if the live
// copy is ever damaged. Best-effort: a backup failure must never block the
// real write.
function backupExcelFileIfDue() {
  try {
    if (excelFileMissingOrEmpty()) return; // nothing good to preserve yet
    if (!fs.existsSync(BACKUPS_DIR)) fs.mkdirSync(BACKUPS_DIR, { recursive: true });

    let alreadyToday = false;
    try { alreadyToday = fs.statSync(BACKUP_FILE).mtime.toDateString() === new Date().toDateString(); } catch (e) { /* no backup yet */ }
    if (!alreadyToday) fs.copyFileSync(EXCEL_FILE, BACKUP_FILE);
  } catch (e) { /* best-effort only */ }
}

// Off-site copy of the same single backup into Google Drive (G:\My Drive), so
// a disk failure / theft / fire that takes out this PC doesn't take the backup
// with it. One file, refreshed at most once per day. Silently does nothing if
// Drive isn't mounted or signed in - this must never block or fail the real write.
function backupToGoogleDriveIfDue() {
  try {
    if (excelFileMissingOrEmpty()) return;
    if (!fs.existsSync(GDRIVE_ROOT)) return; // Drive for Desktop not running/signed in - skip quietly
    if (!fs.existsSync(GDRIVE_BACKUPS_DIR)) fs.mkdirSync(GDRIVE_BACKUPS_DIR, { recursive: true });

    let alreadyToday = false;
    try { alreadyToday = fs.statSync(GDRIVE_BACKUP_FILE).mtime.toDateString() === new Date().toDateString(); } catch (e) { /* no backup yet */ }
    if (!alreadyToday) fs.copyFileSync(EXCEL_FILE, GDRIVE_BACKUP_FILE);
  } catch (e) { /* best-effort only */ }
}

// Write to a temp file in the same folder, then rename over the real workbook.
// A rename is a near-instant metadata swap, so a crash/kill/power-loss during
// the (slower) write only ever leaves a stray .tmp file - the real workbook is
// either the old good copy or the new one, never a truncated 0-byte file.
async function atomicWriteWorkbook(wb) {
  backupExcelFileIfDue();      // snapshot the pre-change state before it's replaced
  backupToGoogleDriveIfDue();  // ...and the same snapshot, off-site

  // best-effort sweep of any .tmp file orphaned by a previous crash/kill
  try {
    const prefix = path.basename(EXCEL_FILE) + '.tmp-';
    for (const f of fs.readdirSync(DOCS_DIR)) {
      if (f.startsWith(prefix)) { try { fs.unlinkSync(path.join(DOCS_DIR, f)); } catch (e) { /* ignore */ } }
    }
  } catch (e) { /* ignore */ }

  const tmp = EXCEL_FILE + '.tmp-' + process.pid + '-' + Date.now();
  try {
    await wb.xlsx.writeFile(tmp);
    fs.renameSync(tmp, EXCEL_FILE);
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (e2) { /* best effort cleanup */ }
    throw e;
  }
}

// True while the workbook is open in Excel. Excel reads the file into memory and
// releases the handle, so we can't rely on a file lock - instead we look for the
// "~$<name>" owner file Excel keeps beside it for the whole session. DDR must not
// write the workbook while this is true, or Excel drops the disk change / turns
// the open copy read-only.
function isExcelFileOpen() {
  const base = path.basename(EXCEL_FILE);
  const owners = ['~$' + base, '~$' + base.slice(2)];
  return owners.some((name) => {
    try { return fs.existsSync(path.join(DOCS_DIR, name)); } catch (e) { return false; }
  });
}

function newEntriesSheet(wb) {
  const ws = wb.addWorksheet(EXCEL_SHEET);
  ws.columns = HEADERS.map((h) => ({ header: h, width: 14 }));
  return ws;
}

// Files touched by SheetJS (and re-emitted by ExcelJS) carry
// defaultRowHeight="0", which collapses every row that has no explicit height -
// so the grid below the data looks like an empty grey void and its cells can't
// be clicked. Restore Excel's normal default; leave any real custom height alone.
function fixRowHeights(wb) {
  let changed = false;
  for (const ws of wb.worksheets) {
    const h = ws.properties && ws.properties.defaultRowHeight;
    if (!h || h < 5) {
      ws.properties.defaultRowHeight = 15;
      changed = true;
    }
  }
  return changed;
}

// Mirror the archive into the "Entries" sheet: append rows it doesn't have yet
// and remove rows that were deleted (see DELETED_FILE) but couldn't be removed
// at delete-time (e.g. Excel was open). Matching is by row CONTENT, not
// position, so this stays correct even after a deletion shifts everything
// below it up by one row. Existing (non-deleted) rows are never rewritten and
// the user's other sheets / extra columns / styling are never touched. A
// fresh sheet (or workbook) is created only when none exists yet.
async function writeExcelFromArchive(archive) {
  ensureDocsDir();
  const fileExists = !excelFileMissingOrEmpty();

  // Never write while the user has the workbook open - the entry is safe in the
  // archive and the pending-sync loop applies it once Excel is closed.
  if (fileExists && isExcelFileOpen()) {
    throw new Error('Excel file is open - changes will sync when you close it');
  }

  const wb = fileExists ? await openWorkbook() : new ExcelJS.Workbook();

  // Present but unreadable (open in Excel, or an unexpected format). Never
  // overwrite it blind - let the caller retry or stay pending.
  if (fileExists && !wb) throw new Error('Excel file exists but could not be opened');

  const N = HEADERS.length;
  const deletedKeys = readDeletedKeys();

  let ws = wb.getWorksheet(EXCEL_SHEET);
  if (!ws) {
    ws = newEntriesSheet(wb);
    for (const row of archive) {
      if (deletedKeys.has(rowKey(row, N))) continue; // don't resurrect a deleted entry
      ws.addRow(row.map(cellValue));
    }
    fixRowHeights(wb);
    await atomicWriteWorkbook(wb);
    return;
  }

  const index = sheetRowIndex(ws); // content-key -> row number(s) currently in the sheet

  const toRemove = [];
  for (const [k, rows] of index) if (deletedKeys.has(k)) toRemove.push(...rows);
  toRemove.sort((a, b) => b - a); // bottom-up, so earlier removals don't shift later row numbers
  for (const rn of toRemove) ws.spliceRows(rn, 1);

  const missing = archive.filter((row) => {
    const k = rowKey(row, N);
    return !index.has(k) && !deletedKeys.has(k);
  });

  const gapsFilled = fillRowGaps(ws);
  const heightsFixed = fixRowHeights(wb);
  for (const row of missing) ws.addRow(row.map(cellValue));

  if (missing.length === 0 && toRemove.length === 0 && !gapsFilled && !heightsFixed) return; // already up to date
  await atomicWriteWorkbook(wb);
}

// Overwrite one row of the "Entries" sheet in place (for an edited entry),
// keeping that row's formatting. Matches on the row's previous values (exact,
// then a loose date/time/name/mobile key); if it still cannot be found the
// corrected row is appended so the edit is never lost. Other sheets and any
// extra columns are left alone.
async function updateExcelRow(oldRow, newRow) {
  ensureDocsDir();
  if (isExcelFileOpen()) {
    throw new Error('Excel file is open - changes will sync when you close it');
  }
  const wb = await openWorkbook();
  if (!wb) throw new Error('workbook unavailable');
  const ws = wb.getWorksheet(EXCEL_SHEET);
  if (!ws) throw new Error('no Entries sheet');

  const N = HEADERS.length;
  const exact = rowKey(oldRow, N);
  const loose = JSON.stringify([0, 1, 2, 3].map((j) => String(normCell(oldRow[j]))));

  let exactRow = -1;
  let looseRow = -1;
  for (let rn = ws.rowCount; rn >= 2; rn--) {
    const vals = ws.getRow(rn).values; // 1-indexed; [0] is empty
    const arr = [];
    for (let c = 1; c <= N; c++) arr.push(cellToPlain(vals[c]));
    if (exactRow === -1 && rowKey(arr, N) === exact) exactRow = rn;
    if (looseRow === -1 &&
        JSON.stringify([0, 1, 2, 3].map((j) => String(normCell(arr[j])))) === loose) looseRow = rn;
    if (exactRow !== -1) break;
  }

  const rn = exactRow !== -1 ? exactRow : looseRow;
  if (rn === -1) {
    ws.addRow(newRow.map(cellValue)); // not found -> append
  } else {
    const row = ws.getRow(rn);
    for (let c = 1; c <= N; c++) row.getCell(c).value = cellValue(newRow[c - 1]);
  }
  fixRowHeights(wb);
  await atomicWriteWorkbook(wb);
  return rn !== -1;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function withRetry(fn, attempts = 6) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      lastErr = e;
      await delay(400);
    }
  }
  throw lastErr;
}

function writeExcelWithRetry(archive, attempts = 6) {
  return withRetry(() => writeExcelFromArchive(archive), attempts);
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1300,
    height: 850,
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  Menu.setApplicationMenu(null);
  win.loadFile('index.html');
}

app.whenReady().then(async () => {
  try {
    const archive = readJSON(ARCHIVE_FILE, []);
    if (archive.length > 0) await writeExcelFromArchive(archive);
  } catch (e) { /* will retry on next entry */ }
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('get-entries', () => readJSON(DATA_FILE, []));
ipcMain.handle('set-entries', (event, entries) => { writeJSON(DATA_FILE, entries); return true; });

ipcMain.handle('append-excel-row', async (event, row) => {
  // Record into the permanent archive first (plain JSON, can't be locked by Excel).
  const archive = readJSON(ARCHIVE_FILE, []);
  archive.push(row);
  writeJSON(ARCHIVE_FILE, archive);

  try {
    await writeExcelFromArchive(archive);
    return { ok: true, path: EXCEL_FILE };
  } catch (e) {
    // Row is safe in the archive regardless; it will be written on the next
    // successful sync (e.g. once the Excel file is closed).
    return { ok: false, recorded: true, error: String(e) };
  }
});

// Merge the working list into the permanent archive instead of only seeding an
// absent one. The archive is append-only history and must never lose rows, so:
//  - archive missing            -> create it from the working list
//  - archive present            -> append only working-list rows not already in it
//  - archive present but broken -> leave it untouched (rebuilding could drop rows)
// The common case (archive already a superset of the working list) writes nothing.
ipcMain.handle('seed-archive-if-empty', async (event, rows) => {
  const fileExists = fs.existsSync(ARCHIVE_FILE);
  const existing = readJSON(ARCHIVE_FILE, []);

  if (fileExists && !Array.isArray(existing)) {
    return { seeded: false, merged: 0, error: 'archive unreadable' };
  }

  const archive = Array.isArray(existing) ? existing.slice() : [];
  const seen = new Set(archive.map((r) => JSON.stringify(r)));
  let merged = 0;
  for (const row of (rows || [])) {
    const key = JSON.stringify(row);
    if (seen.has(key)) continue;
    seen.add(key);
    archive.push(row);
    merged++;
  }

  if (!fileExists || merged > 0) {
    writeJSON(ARCHIVE_FILE, archive);
    try { await writeExcelFromArchive(archive); } catch (e) { /* will retry on next entry */ }
  }
  return { seeded: !fileExists, merged, count: archive.length };
});

// Update an existing row in the permanent archive (and mirror it to Excel).
// The row is located by exact content match on its previous value; if buildRow
// output has drifted, fall back to a loose key (date/time/name/mobile/service/
// charge). If it still cannot be found the edited row is appended rather than
// dropped, so the corrected data always reaches the permanent record.
function looseRowKey(r) {
  return JSON.stringify([r[0], r[1], r[2], r[3], r[8], r[9], r[10]]);
}

ipcMain.handle('update-excel-row', async (event, oldRow, newRow) => {
  const archive = readJSON(ARCHIVE_FILE, null);
  if (!Array.isArray(archive)) return { ok: false, error: 'archive unreadable' };

  const exact = JSON.stringify(oldRow);
  let idx = -1;
  for (let i = archive.length - 1; i >= 0; i--) {
    if (JSON.stringify(archive[i]) === exact) { idx = i; break; }
  }
  if (idx === -1) {
    const loose = looseRowKey(oldRow);
    for (let i = archive.length - 1; i >= 0; i--) {
      if (looseRowKey(archive[i]) === loose) { idx = i; break; }
    }
  }

  let matched = true;
  if (idx === -1) { archive.push(newRow); matched = false; }
  else { archive[idx] = newRow; }

  writeJSON(ARCHIVE_FILE, archive);

  try {
    if (!excelFileMissingOrEmpty()) {
      await withRetry(() => updateExcelRow(oldRow, newRow)); // overwrite that row in place
    } else {
      await writeExcelWithRetry(archive);                    // no file yet (or 0 bytes) -> (re)create it
    }
    return { ok: true, path: EXCEL_FILE, matched };
  } catch (e) {
    // Archive already updated; the renderer's pending-sync loop will finish it.
    return { ok: false, recorded: true, matched, error: String(e) };
  }
});

// Delete an entry from the working list + Excel. The archive is NEVER trimmed
// (permanent, recoverable record) - this only tombstones the row's content-key
// so future reconciles skip and remove it. The tombstone is written first and
// unconditionally, so the deletion sticks even if Excel can't be reached right
// now (it gets cleaned up on the next successful writeExcelFromArchive).
ipcMain.handle('delete-excel-row', async (event, row) => {
  addDeletedKey(rowKey(row, HEADERS.length));

  const archive = readJSON(ARCHIVE_FILE, []);
  try {
    await writeExcelWithRetry(archive);
    return { ok: true };
  } catch (e) {
    return { ok: false, recorded: true, error: String(e) };
  }
});

ipcMain.handle('resync-excel', async () => {
  const archive = readJSON(ARCHIVE_FILE, []);
  if (archive.length === 0) return { ok: true, path: EXCEL_FILE, count: 0 };
  try {
    await writeExcelWithRetry(archive);
    return { ok: true, path: EXCEL_FILE, count: archive.length };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
});

ipcMain.handle('open-excel-file', async () => {
  const archive = readJSON(ARCHIVE_FILE, []);
  if (archive.length > 0) {
    try { await writeExcelWithRetry(archive); } catch (e) { /* open whatever's on disk */ }
  }
  ensureDocsDir();
  if (!fs.existsSync(EXCEL_FILE)) return false;
  shell.openPath(EXCEL_FILE);
  return true;
});

ipcMain.handle('get-excel-path', () => EXCEL_FILE);
