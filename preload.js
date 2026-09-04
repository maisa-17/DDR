const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getEntries: () => ipcRenderer.invoke('get-entries'),
  setEntries: (entries) => ipcRenderer.invoke('set-entries', entries),
  appendExcelRow: (row) => ipcRenderer.invoke('append-excel-row', row),
  updateExcelRow: (oldRow, newRow) => ipcRenderer.invoke('update-excel-row', oldRow, newRow),
  seedArchiveIfEmpty: (rows) => ipcRenderer.invoke('seed-archive-if-empty', rows),
  resyncExcel: () => ipcRenderer.invoke('resync-excel'),
  openExcelFile: () => ipcRenderer.invoke('open-excel-file'),
  getExcelPath: () => ipcRenderer.invoke('get-excel-path')
});
