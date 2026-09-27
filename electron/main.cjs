const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");

// Pin userData to a fixed folder name, independent of package.json's
// "name"/"productName". Without this, running `electron .` (which uses
// "name": "shift-priority-app") and running the built portable .exe (which
// uses "productName": "Shift Priority") would each get their OWN separate
// data folder — and if productName is ever renamed later, every existing
// user's saved data would appear to vanish (it would still be on disk, just
// under the old folder). "Shift Priority" matches the current productName,
// so anyone who already has the app installed keeps their existing data
// with no migration needed.
app.setName("Shift Priority");

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
