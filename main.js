const { app, BrowserWindow, ipcMain, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');
const XLSX = require('xlsx');

const DATA_FILE = path.join(app.getPath('userData'), 'entries.json');
const ARCHIVE_FILE = path.join(app.getPath('userData'), 'entries_archive.json');
const DOCS_DIR = path.join(app.getPath('documents'), 'DDR');
const EXCEL_FILE = path.join(DOCS_DIR, 'service_entries_database.xlsx');
const EXCEL_SHEET = 'Entries';
const CLOSE_SCRIPT = app.isPackaged
  ? path.join(process.resourcesPath, 'app.asar.unpacked', 'close_excel.ps1')
  : path.join(__dirname, 'close_excel.ps1');

const HEADERS = ['Date','Time','Name','Mobile','Type','Source','Method','Location','Main Service','Sub Service',
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

function writeExcelFromArchive(archive) {
  ensureDocsDir();
  const ws = XLSX.utils.aoa_to_sheet([HEADERS, ...archive]);
  ws['!cols'] = HEADERS.map(() => ({ wch: 14 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, EXCEL_SHEET);
  XLSX.writeFile(wb, EXCEL_FILE);
}

function closeExcelIfOpen() {
  return new Promise((resolve) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', CLOSE_SCRIPT, '-Path', EXCEL_FILE],
      { timeout: 6000 },
      () => resolve() // best-effort: proceed regardless of outcome
    );
  });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function writeExcelWithRetry(archive, attempts = 6) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      writeExcelFromArchive(archive);
      return;
    } catch (e) {
      lastErr = e;
      await delay(400);
    }
  }
  throw lastErr;
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

app.whenReady().then(() => {
  try {
    const archive = readJSON(ARCHIVE_FILE, []);
    if (archive.length > 0) writeExcelFromArchive(archive);
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

ipcMain.handle('append-excel-row', (event, row) => {
  // Record into the permanent archive first (plain JSON, can't be locked by Excel).
  const archive = readJSON(ARCHIVE_FILE, []);
  archive.push(row);
  writeJSON(ARCHIVE_FILE, archive);

  try {
    writeExcelFromArchive(archive);
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
ipcMain.handle('seed-archive-if-empty', (event, rows) => {
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
    try { writeExcelFromArchive(archive); } catch (e) { /* will retry on next entry */ }
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

  await closeExcelIfOpen();
  try {
    await writeExcelWithRetry(archive);
    return { ok: true, path: EXCEL_FILE, matched };
  } catch (e) {
    // Archive already updated; the renderer's pending-sync loop will finish it.
    return { ok: false, recorded: true, matched, error: String(e) };
  }
});

ipcMain.handle('resync-excel', async () => {
  const archive = readJSON(ARCHIVE_FILE, []);
  if (archive.length === 0) return { ok: true, path: EXCEL_FILE, count: 0 };
  await closeExcelIfOpen();
  try {
    await writeExcelWithRetry(archive);
    return { ok: true, path: EXCEL_FILE, count: archive.length };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
});

ipcMain.handle('open-excel-file', async () => {
  const archive = readJSON(ARCHIVE_FILE, []);
  await closeExcelIfOpen();
  if (archive.length > 0) {
    try { await writeExcelWithRetry(archive); } catch (e) { /* open whatever's on disk */ }
  }
  ensureDocsDir();
  if (!fs.existsSync(EXCEL_FILE)) return false;
  shell.openPath(EXCEL_FILE);
  return true;
});

ipcMain.handle('get-excel-path', () => EXCEL_FILE);
