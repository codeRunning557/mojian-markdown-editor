const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('markdownApp', {
  ping: () => 'pong',
  openMarkdownFile: () => ipcRenderer.invoke('dialog:openMarkdownFile'),
  openDocxFile: () => ipcRenderer.invoke('dialog:openDocxFile'),
  saveMarkdownFile: (payload) => ipcRenderer.invoke('dialog:saveMarkdownFile', payload),
  saveImportedMarkdownFile: (payload) => ipcRenderer.invoke('dialog:saveImportedMarkdownFile', payload),
  materializeImportedMarkdown: (payload) => ipcRenderer.invoke('dialog:materializeImportedMarkdown', payload),
  readMarkdownFile: (filePath) => ipcRenderer.invoke('dialog:readMarkdownFile', filePath),
  exportHtmlFile: (payload) => ipcRenderer.invoke('dialog:exportHtmlFile', payload),
  exportPdfFile: (payload) => ipcRenderer.invoke('dialog:exportPdfFile', payload),
  saveImageAsset: (payload) => ipcRenderer.invoke('dialog:saveImageAsset', payload),
  pickImageReference: (payload) => ipcRenderer.invoke('dialog:pickImageReference', payload),
  readAiConfigDocument: () => ipcRenderer.invoke('config:readAiConfigDocument'),
  writeAiConfigDocument: (payload) => ipcRenderer.invoke('config:writeAiConfigDocument', payload),
  openExternalLink: (url) => ipcRenderer.invoke('shell:openExternalLink', url),
  syncDocumentState: (payload) => ipcRenderer.send('document-state:sync', payload),
  onMenuAction: (callback) => {
    const listener = (_event, action) => callback(action);
    ipcRenderer.on('menu-action', listener);

    return () => {
      ipcRenderer.removeListener('menu-action', listener);
    };
  }
});
