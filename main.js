const { app, BrowserWindow } = require('electron');
const path = require('path');

// ✅ @electron/remote 등록 추가
const remoteMain = require('@electron/remote/main');
remoteMain.initialize();

function createWindow() {
  const win = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true // ✅ 필요
    }
  });

  win.loadFile('index.html');

  // ✅ Remote 활성화
  remoteMain.enable(win.webContents);
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
