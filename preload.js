const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  minimize: () => ipcRenderer.send('window-minimize'),
  maximize: () => ipcRenderer.send('window-maximize'),
  close: () => ipcRenderer.send('window-close'),
  openFileDialog: () => ipcRenderer.invoke('dialog:openFile'),
  onOpenFile: (callback) => ipcRenderer.on('open-file', (event, filePath, fileData, ext) => callback(filePath, fileData, ext))
});
