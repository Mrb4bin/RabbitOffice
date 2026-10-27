// Safe bridge between Rabbit Office web apps and the desktop shell.
const { contextBridge, ipcRenderer, webUtils } = require('electron');
let beforeClose = null;
ipcRenderer.on('before-close', async () => { const ok = beforeClose ? await beforeClose() : true; if (ok) ipcRenderer.send('close-ok'); });
contextBridge.exposeInMainWorld('rabbitNative', {
  platform: process.platform,
  openWindow: (url) => ipcRenderer.invoke('open-window', url),
  openFile: (o) => ipcRenderer.invoke('open-file', o),
  readFile: (p) => ipcRenderer.invoke('read-file', p),
  writeFile: (p, b64) => ipcRenderer.invoke('write-file', p, b64),
  saveDialog: (o) => ipcRenderer.invoke('save-dialog', o),
  printToPDF: (p, o) => ipcRenderer.invoke('print-pdf', p, o),
  setTitleBar: (o) => ipcRenderer.send('set-titlebar', o),
  setDirty: (d) => ipcRenderer.send('set-dirty', d),
  onBeforeClose: (fn) => { beforeClose = fn; },
  pathForFile: (f) => { try { return webUtils.getPathForFile(f); } catch { return null; } },
  desk: {
    host: (o) => ipcRenderer.invoke('desk-host', o),
    stop: () => ipcRenderer.invoke('desk-stop'),
    status: () => ipcRenderer.invoke('desk-status'),
    discover: (cb) => { ipcRenderer.removeAllListeners('desk-found'); ipcRenderer.on('desk-found', (e, d) => cb(d)); ipcRenderer.send('desk-discover'); },
  },
});
