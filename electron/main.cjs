const { app, BrowserWindow, ipcMain, dialog, Notification, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

let mainWindow = null;
let serverProcess = null;

// Prevent any uncaught exceptions from showing error dialogs
process.on('uncaughtException', (err) => {
  console.error('[Main Process Error]:', err);
});

process.on('unhandledRejection', (reason) => {
  console.error('[Main Process Rejection]:', reason);
});

function startBackendServer() {
  const isDev = !app.isPackaged && process.env.NODE_ENV !== 'production';
  if (!isDev) {
    // In production, desktop app runs client-side offline engines without dev server CLI
    return;
  }

  try {
    const cmd = process.platform === 'win32' ? 'npx.cmd' : 'npx';
    serverProcess = spawn(cmd, ['tsx', 'server/index.ts'], {
      cwd: path.join(__dirname, '..'),
      env: { ...process.env, PORT: '5000' },
    });

    serverProcess.on('error', (err) => {
      console.warn('[Backend Spawn Notice]:', err.message);
    });

    serverProcess.stdout?.on('data', (data) => {
      console.log(`[Backend]: ${data}`);
    });

    serverProcess.stderr?.on('data', (data) => {
      console.error(`[Backend Error]: ${data}`);
    });
  } catch (err) {
    console.warn('Backend server not started:', err);
  }
}

function createWindow() {
  let appIcon = null;
  const icoPath = path.join(__dirname, '../public/icon.ico');
  const pngPath = path.join(__dirname, '../dist/logo.png');

  try {
    if (fs.existsSync(icoPath)) {
      appIcon = nativeImage.createFromPath(icoPath);
    } else if (fs.existsSync(pngPath)) {
      appIcon = nativeImage.createFromPath(pngPath);
    }
  } catch (e) {
    console.warn('Icon load notice:', e);
  }

  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 1024,
    minHeight: 700,
    title: 'CA Copilot',
    icon: appIcon || undefined,
    autoHideMenuBar: true,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#F8F9FA',
      symbolColor: '#1F2937',
      height: 36,
    },
    show: true,
    backgroundColor: '#F8F9FA',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  const isDev = !app.isPackaged && process.env.NODE_ENV !== 'production';

  if (isDev && process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    // In production, load the built static dist/index.html
    const indexPath = path.join(__dirname, '../dist/index.html');
    mainWindow.loadFile(indexPath).catch((err) => {
      console.error('Error loading dist/index.html:', err);
    });
  }

  mainWindow.show();
  mainWindow.focus();

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// IPC Handlers for Native Windows OS Integration
ipcMain.handle('dialog:showSaveDialog', async (_event, options) => {
  if (!mainWindow) return null;
  return await dialog.showSaveDialog(mainWindow, options);
});

ipcMain.handle('dialog:showOpenDialog', async (_event, options) => {
  if (!mainWindow) return null;
  return await dialog.showOpenDialog(mainWindow, options);
});

ipcMain.on('app:notification', () => {
  // All notifications disabled per user request
});

app.whenReady().then(() => {
  startBackendServer();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (serverProcess) {
    try {
      serverProcess.kill();
    } catch (e) {
      // Ignored
    }
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
