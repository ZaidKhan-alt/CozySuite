const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs').promises;

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 1024,
    minHeight: 680,
    title: 'Cozy Office',
    frame: false, // frameless window for custom soft UI titlebar
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  // Load the built dist file, or locally if not built yet (assuming built for production)
  mainWindow.loadFile(path.join(__dirname, 'dist/index.html'));
  
  // Custom Window Controls
  ipcMain.on('window-minimize', () => mainWindow.minimize());
  ipcMain.on('window-maximize', () => {
    if (mainWindow.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow.maximize();
    }
  });
  ipcMain.on('window-close', () => mainWindow.close());
  
  mainWindow.webContents.on('did-finish-load', () => {
    handleProcessArgs(process.argv);
  });
}

// IPC handler for open file dialog
ipcMain.handle('dialog:openFile', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Office Files', extensions: ['docx', 'xlsx', 'pptx', 'pdf', 'csv'] }
    ]
  });
  
  if (!canceled && filePaths.length > 0) {
    await sendFileToRenderer(filePaths[0]);
  }
});

async function sendFileToRenderer(filePath) {
  if (!mainWindow) return;
  try {
    const ext = path.extname(filePath).toLowerCase();
    // For pdf we might just need the path or buffer, let's send buffer as base64 or arraybuffer
    const data = await fs.readFile(filePath);
    mainWindow.webContents.send('open-file', filePath, data.buffer, ext);
  } catch (err) {
    console.error('Error reading file:', err);
  }
}

function handleProcessArgs(argv) {
  if (argv.length >= 2) {
    const filePath = argv[1];
    if (filePath && filePath !== '.' && !filePath.startsWith('--')) {
      sendFileToRenderer(filePath);
    }
  }
}

// macOS open-file event
app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (app.isReady()) {
    sendFileToRenderer(filePath);
  } else {
    app.once('ready', () => {
      sendFileToRenderer(filePath);
    });
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});
