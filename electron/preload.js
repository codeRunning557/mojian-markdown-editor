const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('markdownApp', {
  ping: () => 'pong',
  openMarkdownFile: () => ipcRenderer.invoke('dialog:openMarkdownFile'),
  saveMarkdownFile: (payload) => ipcRenderer.invoke('dialog:saveMarkdownFile', payload),
  readMarkdownFile: (filePath) => ipcRenderer.invoke('dialog:readMarkdownFile', filePath),
  exportHtmlFile: (payload) => ipcRenderer.invoke('dialog:exportHtmlFile', payload),
  exportPdfFile: (payload) => ipcRenderer.invoke('dialog:exportPdfFile', payload),
  saveImageAsset: (payload) => ipcRenderer.invoke('dialog:saveImageAsset', payload),
  pickImageReference: (payload) => ipcRenderer.invoke('dialog:pickImageReference', payload),
  onMenuAction: (callback) => {
    const listener = (_event, action) => callback(action);
    ipcRenderer.on('menu-action', listener);

    return () => {
      ipcRenderer.removeListener('menu-action', listener);
    };
  }
});
