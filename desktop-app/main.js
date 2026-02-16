const { app, BrowserWindow, shell } = require('electron');

const APP_URL = 'https://confirmations3d.online/';

async function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    backgroundColor: '#f0f2f5',
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  // Force latest domain content on each app start.
  await win.webContents.session.clearCache();
  const version = app.getVersion();
  const sep = APP_URL.includes('?') ? '&' : '?';
  const liveUrl = `${APP_URL}${sep}appVersion=${encodeURIComponent(version)}&t=${Date.now()}`;
  win.loadURL(liveUrl);

  win.once('ready-to-show', () => {
    win.show();
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
