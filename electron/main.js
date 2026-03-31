const path = require('node:path');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const { app, BrowserWindow, dialog, ipcMain, Menu } = require('electron');

let mainWindow = null;

process.on('uncaughtException', (error) => {
  console.error('[main] uncaught exception');
  console.error(error);
});

function emitMenuAction(action) {
  if (!mainWindow) {
    return;
  }

  mainWindow.webContents.send('menu-action', action);
}

function createMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [{ role: 'about' }, { type: 'separator' }, { role: 'quit' }]
          }
        ]
      : []),
    {
      label: '文件',
      submenu: [
        { label: '新建', accelerator: 'CmdOrCtrl+N', click: () => emitMenuAction('new-file') },
        { label: '打开...', accelerator: 'CmdOrCtrl+O', click: () => emitMenuAction('open-file') },
        { label: '保存', accelerator: 'CmdOrCtrl+S', click: () => emitMenuAction('save-file') },
        { label: '另存为...', accelerator: 'CmdOrCtrl+Shift+S', click: () => emitMenuAction('save-file-as') },
        { label: '导出 HTML...', accelerator: 'CmdOrCtrl+E', click: () => emitMenuAction('export-html') },
        { label: '导出 PDF...', accelerator: 'CmdOrCtrl+Shift+E', click: () => emitMenuAction('export-pdf') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' }
      ]
    },
    {
      label: '主题',
      submenu: [
        { label: '米白', click: () => emitMenuAction('theme-paper') },
        { label: '雾蓝', click: () => emitMenuAction('theme-mist') },
        { label: '灰砚', click: () => emitMenuAction('theme-slate') },
        { label: '石墨', click: () => emitMenuAction('theme-graphite') }
      ]
    },
    {
      label: '视图',
      submenu: [
        { label: '统一工作区', accelerator: 'Alt+1', click: () => emitMenuAction('view-preview') },
        { type: 'separator' },
        { role: 'reload' },
        { role: 'toggledevtools' }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1080,
    minHeight: 720,
    show: false,
    backgroundColor: '#efe7d8',
    title: '墨笺 Markdown 编辑器',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow?.show();
    mainWindow?.focus();
  });

  mainWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    console.error('[main] failed to load window', errorCode, errorDescription);
  });

  mainWindow.webContents.on('render-process-gone', (_event, details) => {
    console.error('[main] render process gone', details);
  });

  const devServerUrl = process.env.VITE_DEV_SERVER_URL;
  console.log('[main] createWindow', { devServerUrl: devServerUrl ?? null });
  if (devServerUrl) {
    mainWindow.loadURL(devServerUrl);
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

ipcMain.handle('dialog:openMarkdownFile', async () => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const { canceled, filePaths } = await dialog.showOpenDialog(targetWindow, {
    title: '打开 Markdown 文件',
    properties: ['openFile'],
    filters: [
      { name: 'Markdown', extensions: ['md', 'markdown', 'mdown', 'txt'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (canceled || filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = filePaths[0];
  const content = await fs.readFile(filePath, 'utf8');
  return { canceled: false, filePath, content };
});

ipcMain.handle('dialog:saveMarkdownFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const shouldShowDialog = !payload?.filePath || payload?.forceDialog;

  let filePath = payload?.filePath ?? '';
  if (shouldShowDialog) {
    const result = await dialog.showSaveDialog(targetWindow, {
      title: '保存 Markdown 文件',
      defaultPath: filePath || 'untitled.md',
      filters: [
        { name: 'Markdown', extensions: ['md'] },
        { name: '文本', extensions: ['txt'] }
      ]
    });

    if (result.canceled || !result.filePath) {
      return { canceled: true };
    }

    filePath = result.filePath;
  }

  await fs.writeFile(filePath, payload?.content ?? '', 'utf8');
  return { canceled: false, filePath };
});

ipcMain.handle('dialog:readMarkdownFile', async (_event, filePath) => {
  const content = await fs.readFile(filePath, 'utf8');
  return { filePath, content };
});

ipcMain.handle('dialog:exportHtmlFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showSaveDialog(targetWindow, {
    title: '导出 HTML 文件',
    defaultPath: payload?.defaultPath || 'untitled.html',
    filters: [{ name: 'HTML', extensions: ['html'] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  await fs.writeFile(result.filePath, payload?.html ?? '', 'utf8');
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle('dialog:exportPdfFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showSaveDialog(targetWindow, {
    title: '导出 PDF 文件',
    defaultPath: payload?.defaultPath || 'untitled.pdf',
    filters: [{ name: 'PDF', extensions: ['pdf'] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const pdfData = await targetWindow.webContents.printToPDF({
    printBackground: true,
    preferCSSPageSize: true
  });
  await fs.writeFile(result.filePath, pdfData);
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle('dialog:saveImageAsset', async (_event, payload) => {
  const originalName = payload?.originalName || `image-${Date.now()}.png`;
  const safeName = originalName.replace(/[<>:"/\\|?*\x00-\x1f]/g, '-');
  const parsed = path.parse(safeName);
  const fileName = parsed.ext ? safeName : `${safeName}.png`;

  let targetDir = '';
  if (payload?.documentPath) {
    targetDir = path.join(path.dirname(payload.documentPath), 'assets');
  } else {
    const chosenDir = await dialog.showOpenDialog(BrowserWindow.getFocusedWindow() ?? mainWindow, {
      title: '选择图片保存目录',
      properties: ['openDirectory', 'createDirectory']
    });

    if (chosenDir.canceled || chosenDir.filePaths.length === 0) {
      return { canceled: true };
    }

    targetDir = chosenDir.filePaths[0];
  }

  if (!fsSync.existsSync(targetDir)) {
    await fs.mkdir(targetDir, { recursive: true });
  }

  const filePath = path.join(targetDir, fileName);
  const base64Data = String(payload?.dataUrl || '').replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '');
  const buffer = Buffer.from(base64Data, 'base64');
  await fs.writeFile(filePath, buffer);

  return { canceled: false, filePath };
});

ipcMain.handle('dialog:pickImageReference', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showOpenDialog(targetWindow, {
    title: '选择要引用的图片',
    properties: ['openFile'],
    filters: [
      { name: '图片', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = result.filePaths[0];
  let markdownPath = filePath.replace(/\\/g, '/');

  if (payload?.documentPath) {
    const relativePath = path.relative(path.dirname(payload.documentPath), filePath).replace(/\\/g, '/');
    markdownPath = relativePath.startsWith('.') ? relativePath : `./${relativePath}`;
  } else {
    markdownPath = `file:///${markdownPath.replace(/^\/+/, '')}`;
  }

  return {
    canceled: false,
    filePath,
    markdownPath,
    displayName: path.basename(filePath)
  };
});

app.whenReady().then(() => {
  createMenu();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
