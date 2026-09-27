// The app's ONLY data access point. Nothing outside this file (and
// backup.js/migrations.js, which it composes) should touch localStorage or
// indexedDB directly. localStorage is read exactly once, by
// migrations.js, to seed IndexedDB the first time this runs — after that
// it is never read or written again.
//
// Every getter here is synchronous and reads from an in-memory cache that
// is hydrated before the app renders (see ready()). Every setter updates
// the cache immediately (so React sees the change with no delay) and then
// persists to IndexedDB asynchronously, broadcasts the change to any other
// open tab, and lets backup.js decide whether this warrants a snapshot.

import { openDB, tx, getAll, get, put, del } from "./db.js";
import { DB_NAME, DB_VERSION, STORES, DEFAULTS, CURRENT_SCHEMA_VERSION, upgrade } from "./schema.js";
import { migrateLegacyIfNeeded, readLegacyForDisplay } from "./migrations.js";

const cache = { kv: { ...DEFAULTS }, dailyLog: [], schedule: {} };
const listeners = new Set();
const writeListeners = new Set(); // backup.js hooks in here — no circular import needed

let db = null;
let readOnly = false;
let readOnlyReason = null; // "downgrade" | "migration-failed" | "unavailable" | null
let initReport = null;
let readyResolve;
const readyPromise = new Promise((res) => { readyResolve = res; });

function notify() { listeners.forEach((fn) => { try { fn(); } catch { /* listener's problem */ } }); }
function notifyWrite(key) { writeListeners.forEach((fn) => { try { fn(key); } catch { /* ignore */ } }); }

let channel = null;
try { channel = new BroadcastChannel("shift-priority-sync"); } catch { /* older browser */ }
if (channel) {
  channel.onmessage = (e) => {
    if (e.data?.type === "kv" && e.data.key) { hydrateKvKey(e.data.key).then(notify); }
    else if (e.data?.type === "dailyLog") { hydrateDailyLog().then(notify); }
    else if (e.data?.type === "reload-required") { notify(); }
  };
}

async function hydrateKvKey(key) {
  const row = await get(dbStore(STORES.kv), key).catch(() => null);
  cache.kv[key] = row ? row.value : DEFAULTS[key];
}
async function hydrateDailyLog() {
  const rows = await getAll(dbStore(STORES.dailyLog)).catch(() => []);
  cache.dailyLog = rows.sort((a, b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
}
function dbStore(name, mode = "readonly") {
  return db.transaction(name, mode).objectStore(name);
}

export async function init() {
  readOnly = false; readOnlyReason = null; initReport = null;
  try {
    db = await openDB(DB_NAME, DB_VERSION, (idb, oldV) => upgrade(idb, oldV));
  } catch (e) {
    readOnly = true; readOnlyReason = "unavailable";
    initReport = { error: String(e) };
    Object.assign(cache.kv, DEFAULTS);
    readyResolve();
    return getInitReport();
  }

  db.onversionchange = () => { db.close(); channel?.postMessage({ type: "reload-required" }); notify(); };

  let migration = { ranMigration: false };
  try {
    migration = await migrateLegacyIfNeeded(db);
  } catch (e) {
    readOnly = true; readOnlyReason = "migration-failed";
    initReport = { migrationError: String(e) };
    // Fall back to showing the person their old data, read-only, straight
    // from localStorage — never silently show them an empty app.
    const legacy = readLegacyForDisplay();
    cache.kv.profile = legacy.profile ?? DEFAULTS.profile;
    cache.kv.crewNames = legacy.crewNames ?? DEFAULTS.crewNames;
    cache.kv.theme = legacy.theme ?? DEFAULTS.theme;
    cache.kv.dailyLogAccess = legacy.dailyLogAccess ?? DEFAULTS.dailyLogAccess;
    cache.kv.sound = DEFAULTS.sound;
    cache.kv.lang = DEFAULTS.lang;
    cache.dailyLog = (legacy.dailyLogEntries || []).slice().sort((a, b) => a.date.localeCompare(b.date));
    cache.schedule = legacy.lastFile || {};
    readyResolve();
    return getInitReport();
  }

  const versionRow = await get(dbStore(STORES.meta), "schemaVersion").catch(() => null);
  const storedVersion = versionRow ? versionRow.value : CURRENT_SCHEMA_VERSION;
  if (storedVersion > CURRENT_SCHEMA_VERSION) {
    readOnly = true; readOnlyReason = "downgrade";
    initReport = { storedVersion, currentVersion: CURRENT_SCHEMA_VERSION };
  }

  const kvRows = await getAll(dbStore(STORES.kv)).catch(() => []);
  for (const k of Object.keys(DEFAULTS)) cache.kv[k] = DEFAULTS[k];
  for (const row of kvRows) cache.kv[row.key] = row.value;
  await hydrateDailyLog();
  cache.schedule = (await get(dbStore(STORES.schedule), "current").catch(() => null)) || {};

  initReport = initReport || { ranMigration: migration.ranMigration, ...migration };
  readyResolve();
  return getInitReport();
}

export function ready() { return readyPromise; }
export function isReadOnly() { return readOnly; }
export function readOnlyReasonCode() { return readOnlyReason; }
export function getInitReport() { return initReport; }
export function subscribe(fn) { listeners.add(fn); return () => listeners.delete(fn); }
export function onWrite(fn) { writeListeners.add(fn); return () => writeListeners.delete(fn); }

// ---- generic kv (sync read, async persisted write) ----
function getKv(key) { return cache.kv[key]; }
async function setKv(key, value) {
  if (readOnly) return;
  cache.kv[key] = value;
  notify();
  try {
    await put(dbStore(STORES.kv, "readwrite"), { key, value });
    channel?.postMessage({ type: "kv", key });
    notifyWrite(key);
  } catch (e) {
    notify(); // let the UI know the write failed by re-emitting; caller checks lastWriteError()
    lastWriteError = String(e);
  }
}
let lastWriteError = null;
export function getLastWriteError() { return lastWriteError; }

// ---- profile ----
export function loadProfile() { return getKv("profile"); }
export function saveProfile(p) { return setKv("profile", p); }
export function clearProfileStorage() { return setKv("profile", DEFAULTS.profile); }

// ---- crew name directory ----
export function loadCrewNames() { return getKv("crewNames") || {}; }
export function saveCrewNames(map) { return setKv("crewNames", map); }

// ---- theme ----
export function loadThemePrefs() { return getKv("theme") || DEFAULTS.theme; }
export function saveThemePrefs(style, mode) { return setKv("theme", { style, mode }); }

// ---- sound ----
export function loadSoundSettings() { return getKv("sound") || DEFAULTS.sound; }
export function saveSoundSettings(s) { return setKv("sound", s); }

// ---- language ----
export function loadLang() { return getKv("lang") || DEFAULTS.lang; }
export function saveLang(l) { return setKv("lang", l); }

// ---- daily log access list (legacy admin feature, kept for compatibility) ----
export function loadDailyLogAccess() { return getKv("dailyLogAccess") || []; }
export function saveDailyLogAccess(list) { return setKv("dailyLogAccess", list); }

// ---- daily log entries ----
// Mirrors the app's existing usage: callers hold the full array in React
// state and call this with the full next array after any change.
export function loadDailyLogEntries() { return cache.dailyLog.slice(); }
export async function saveDailyLogEntries(nextList) {
  if (readOnly) return;
  const prevIds = new Set(cache.dailyLog.map((e) => e.id));
  const nextIds = new Set(nextList.map((e) => e.id));
  cache.dailyLog = nextList.slice().sort((a, b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
  notify();
  try {
    await tx(db, [STORES.dailyLog], "readwrite", async (t) => {
      const store = t.objectStore(STORES.dailyLog);
      for (const id of prevIds) if (!nextIds.has(id)) await del(store, id);
      for (const e of nextList) await put(store, { ...e, updatedAt: e.updatedAt || Date.now() });
    });
    channel?.postMessage({ type: "dailyLog" });
    notifyWrite("dailyLog");
  } catch (e) {
    lastWriteError = String(e);
    notify();
  }
}

// ---- schedule (current parsed Excel board) ----
export function loadLastFile() {
  const s = cache.schedule;
  return s && s.parsed ? { fileName: s.fileName, sheetName: s.sheetName, parsed: s.parsed } : null;
}
export async function saveLastFile(entry) {
  if (readOnly) return;
  // Keep one previous board around too, in case a bad upload needs undoing.
  const prevCurrent = cache.schedule;
  cache.schedule = { ...entry, savedAt: Date.now() };
  notify();
  try {
    await tx(db, [STORES.schedule], "readwrite", async (t) => {
      const store = t.objectStore(STORES.schedule);
      if (prevCurrent && prevCurrent.parsed) await put(store, { ...prevCurrent, key: "previous" });
      await put(store, { key: "current", ...entry, savedAt: Date.now() });
    });
    notifyWrite("schedule");
  } catch (e) { lastWriteError = String(e); notify(); }
}
export async function clearLastFileStorage() {
  if (readOnly) return;
  cache.schedule = {};
  notify();
  try { await del(dbStore(STORES.schedule, "readwrite"), "current"); notifyWrite("schedule"); }
  catch (e) { lastWriteError = String(e); notify(); }
}

// ---- exposed for backup.js / diagnostics only ----
export function _internal() { return { db, cache, readOnly, readOnlyReason }; }
