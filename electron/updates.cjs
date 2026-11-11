// Automatic updates from GitHub Releases (electron-updater).
// Checks shortly after start and every 6 hours, downloads in the background and
// installs on the next quit, or right away when the user clicks "Restart".
// Off in development, when KERTAS_NO_UPDATE is set, or when turned off in the app.
const { app, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

const SETTINGS = () => path.join(app.getPath('userData'), 'settings.json');
const readSettings = () => { try { return JSON.parse(fs.readFileSync(SETTINGS(), 'utf8')); } catch { return {}; } };
const writeSettings = (s) => fs.writeFileSync(SETTINGS(), JSON.stringify({ ...readSettings(), ...s }, null, 2));

let updater = null;
let state = { state: 'idle', version: app.getVersion() };
let send = () => {};

function setState(s) {
  state = { ...state, ...s };
  send(state);
}

// Builds made without a publish target have no app-update.yml and cannot update
const configured = () => !app.isPackaged || fs.existsSync(path.join(process.resourcesPath, 'app-update.yml'));

function enabled() {
  if (process.env.KERTAS_NO_UPDATE) return false;
  if (!app.isPackaged && !process.env.KERTAS_UPDATE_TEST) return false;
  if (!configured()) return false;
  return readSettings().autoUpdate !== false;
}

function load() {
  if (updater) return updater;
  ({ autoUpdater: updater } = require('electron-updater'));
  updater.autoDownload = true;
  updater.autoInstallOnAppQuit = true;
  updater.allowPrerelease = false;
  if (process.env.KERTAS_UPDATE_TEST) updater.forceDevUpdateConfig = true; // reads dev-app-update.yml
  updater.logger = { info: () => {}, warn: console.warn, error: console.error, debug: () => {} };
  updater.on('checking-for-update', () => setState({ state: 'checking', message: null }));
  updater.on('update-available', (i) => setState({ state: 'downloading', available: i.version, percent: 0, notes: typeof i.releaseNotes === 'string' ? i.releaseNotes.replace(/<[^>]+>/g, '').slice(0, 600) : null }));
  updater.on('update-not-available', () => setState({ state: 'current', checkedAt: Date.now() }));
  updater.on('download-progress', (p) => setState({ state: 'downloading', percent: Math.round(p.percent || 0) }));
  updater.on('update-downloaded', (i) => setState({ state: 'ready', available: i.version, percent: 100 }));
  updater.on('error', (e) => setState({ state: 'error', message: String(e?.message || e).split('\n')[0].slice(0, 200) }));
  return updater;
}

async function check() {
  if (!app.isPackaged && !process.env.KERTAS_UPDATE_TEST) { setState({ state: 'dev' }); return state; }
  if (!configured()) { setState({ state: 'unconfigured' }); return state; }
  try { await load().checkForUpdates(); } catch (e) { setState({ state: 'error', message: String(e?.message || e).slice(0, 200) }); }
  return state;
}

/**
 * hooks.send(state) delivers state to the window; hooks.beforeInstall() lets
 * the app close its windows without the unsaved-changes prompt.
 */
function setupUpdates(hooks) {
  send = hooks.send;
  ipcMain.handle('kertas:update', async (_e, op, arg) => {
    switch (op) {
      case 'state': return { ...state, auto: readSettings().autoUpdate !== false, packaged: app.isPackaged };
      case 'check': return check();
      case 'install':
        if (state.state !== 'ready') return false;
        hooks.beforeInstall?.();
        setImmediate(() => load().quitAndInstall(false, true));
        return true;
      case 'set-auto':
        writeSettings({ autoUpdate: !!arg });
        if (arg) check();
        return true;
      default: return null;
    }
  });
  if (!enabled()) return;
  setTimeout(check, 8000);
  setInterval(check, 6 * 60 * 60 * 1000).unref?.();
}

module.exports = { setupUpdates };
