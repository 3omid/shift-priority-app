const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");

function createWindow() {
  const isDev = !app.isPackaged;
  // In dev, Vite serves everything under public/ as-is; once built, Vite
  // copies public/ into the root of dist/, so the packaged app (which only
  // ships the dist/ and electron/ folders, not public/ itself) finds the
  // icon there instead.
  const iconPath = isDev
    ? path.join(__dirname, "../public/icon-512.png")
    : path.join(__dirname, "../dist/icon-512.png");

  const win = new BrowserWindow({
    width: 1200,
    height: 820,
    icon: iconPath,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.cjs"),
    },
  });

  if (isDev) {
    win.loadURL("http://localhost:5173");
  } else {
    win.loadFile(path.join(__dirname, "../dist/index.html"));
  }
}

ipcMain.on("app-quit", () => {
  app.quit();
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
