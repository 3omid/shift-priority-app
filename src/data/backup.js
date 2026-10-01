// Backup layer built on top of store.js. Nothing here talks to
// localStorage or IndexedDB directly — it only calls store.js's public
// functions, plus a small amount of direct snapshot-store access for
// listing/pruning/restoring snapshots.

import pkg from "../../package.json" with { type: "json" };
import { STORES, CURRENT_SCHEMA_VERSION } from "./schema.js";
import { getAll, put, del } from "./db.js";
import * as Store from "./store.js";

const APP_VERSION = pkg.version;
const FORMAT = "shift-priority-backup";
const DEBOUNCE_MS = 5000;
const RETENTION = { auto: 10, daily: 7, weekly: 4, guard: 30 * 864e5 }; // guard = pre-migration/pre-restore, in ms

let debounceTimer = null;

// ---- payload ----
export function buildUserDataPayload() {
  return {
    profile: Store.loadProfile(),
    crewNames: Store.loadCrewNames(),
    // Admin's Employee-ID directory (used by the Shift Exchange form). Older
    // backups don't have it; restore simply leaves the current one alone.
    crewEmployeeIds: Store.loadCrewEmployeeIds(),
    dailyLogAccess: Store.loadDailyLogAccess(),
    // Drivers directory and dispatch extra shifts. Older backups have
    // neither; restore then leaves the current ones alone.
    crewServiceTypes: Store.loadCrewServiceTypes(),
    drivers: Store.loadDrivers(),
    extraShifts: Store.loadExtraShifts(),
    theme: Store.loadThemePrefs(),
    sound: Store.loadSoundSettings(),
    lang: Store.loadLang(),
    dailyLog: Store.loadDailyLogEntries(),
  };
}

// A plain JSON.stringify(obj, arrayOfKeys) only filters top-level keys, so
// it would silently ignore changes made to any nested field (which is most
// of this payload) — sort keys recursively instead, so every field at
// every depth actually participates in the checksum.
function stableStringify(value) {
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}
async function checksumOf(obj) {
  const json = stableStringify(obj);
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(json));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export async function buildBackupFile({ includeSchedule = false } = {}) {
  const data = buildUserDataPayload();
  if (includeSchedule) {
    const internal = Store._internal();
    if (internal.cache.schedule?.parsed) data.schedule = internal.cache.schedule;
  }
  const checksum = await checksumOf(data);
  return {
    format: FORMAT,
    appVersion: APP_VERSION,
    schemaVersion: CURRENT_SCHEMA_VERSION,
    createdAt: new Date().toISOString(),
    checksum,
    data,
  };
}

// ---- export (download / share) ----
export async function exportBackup(opts) {
  const file = await buildBackupFile(opts);
  const json = JSON.stringify(file, null, 2);
  const stamp = file.createdAt.slice(0, 10);
  const filename = `shift-priority-backup-${stamp}.json`;
  if (typeof window !== "undefined") {
    const blob = new Blob([json], { type: "application/json" });
    const shareFile = new File([blob], filename, { type: "application/json" });
    if (navigator.canShare && navigator.canShare({ files: [shareFile] })) {
      try { await navigator.share({ files: [shareFile], title: "Shift Priority backup" }); return { json, shared: true }; }
      catch { /* user cancelled the share sheet — fall through to download */ }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }
  await Store.ready();
  await setMeta("lastExportAt", Date.now());
  return { json, shared: false };
}

// ---- validate an incoming file before touching anything ----
export async function validateBackupFile(rawText) {
  let file;
  try { file = JSON.parse(rawText); } catch { return { ok: false, reason: "not-json" }; }
  if (!file || file.format !== FORMAT) return { ok: false, reason: "wrong-format" };
  if (typeof file.schemaVersion !== "number") return { ok: false, reason: "wrong-format" };
  if (file.schemaVersion > CURRENT_SCHEMA_VERSION) return { ok: false, reason: "too-new", fileVersion: file.schemaVersion };
  const expected = file.checksum;
  const actual = await checksumOf(file.data);
  if (expected !== actual) return { ok: false, reason: "checksum-mismatch" };
  const d = file.data || {};
  return {
    ok: true,
    data: d,
    summary: {
      createdAt: file.createdAt,
      appVersion: file.appVersion,
      hasProfile: !!(d.profile && (d.profile.firstName || d.profile.crewNumber)),
      crewNameCount: d.crewNames ? Object.keys(d.crewNames).length : 0,
      dailyLogCount: Array.isArray(d.dailyLog) ? d.dailyLog.length : 0,
      hasSchedule: !!d.schedule,
    },
  };
}

// ---- restore ----
// mode: "replace" (default) restores every field as-is from the backup.
// "mergeDailyLog" restores everything else as-is too, but combines the
// daily log entry-by-entry (by date), keeping whichever side of each date
// was edited more recently, instead of dropping the other side's entries.
export async function restoreBackup(data, { mode = "replace" } = {}) {
  if (!data || typeof data !== "object") return { ok: false, reason: "no-data" };
  await createSnapshot("pre-restore");

  if (data.profile !== undefined) await Store.saveProfile(data.profile);
  if (data.crewNames !== undefined) await Store.saveCrewNames(data.crewNames);
  if (data.crewEmployeeIds !== undefined) await Store.saveCrewEmployeeIds(data.crewEmployeeIds);
  if (data.dailyLogAccess !== undefined) await Store.saveDailyLogAccess(data.dailyLogAccess);
  if (data.crewServiceTypes !== undefined) await Store.saveCrewServiceTypes(data.crewServiceTypes);
  if (data.drivers !== undefined) await Store.saveDrivers(data.drivers);
  if (data.extraShifts !== undefined) await Store.saveExtraShifts(data.extraShifts);
  if (data.theme !== undefined) await Store.saveThemePrefs(data.theme.style, data.theme.mode);
  if (data.sound !== undefined) await Store.saveSoundSettings(data.sound);
  if (data.lang !== undefined) await Store.saveLang(data.lang);

  const incoming = Array.isArray(data.dailyLog) ? data.dailyLog : [];
  if (mode === "mergeDailyLog") {
    const current = Store.loadDailyLogEntries();
    const byDate = new Map(current.map((e) => [e.date, e]));
    for (const inc of incoming) {
      const existing = byDate.get(inc.date);
      if (!existing || (inc.updatedAt || 0) > (existing.updatedAt || 0)) byDate.set(inc.date, inc);
    }
    await Store.saveDailyLogEntries([...byDate.values()]);
  } else {
    await Store.saveDailyLogEntries(incoming);
  }
  return { ok: true, restored: true };
}

export async function undoLastRestore() {
  const snaps = await listSnapshots();
  const last = snaps.filter((s) => s.reason === "pre-restore").sort((a, b) => b.createdAt - a.createdAt)[0];
  if (!last) return { ok: false, reason: "no-snapshot" };
  await restoreSnapshotById(last.id);
  return { ok: true };
}

// ---- snapshots ----
function dbInternal() { return Store._internal().db; }

async function setMeta(key, value) {
  const db = dbInternal();
  if (!db) return;
  await put(db.transaction(STORES.meta, "readwrite").objectStore(STORES.meta), { key, value });
}
async function getMeta(key) {
  const db = dbInternal();
  if (!db) return null;
  const row = await new Promise((res) => {
    const r = db.transaction(STORES.meta, "readonly").objectStore(STORES.meta).get(key);
    r.onsuccess = () => res(r.result); r.onerror = () => res(null);
  });
  return row ? row.value : null;
}

export async function createSnapshot(reason) {
  const db = dbInternal();
  if (!db || Store.isReadOnly()) return null;
  const payload = buildUserDataPayload();
  const id = `${reason}-${Date.now()}`;
  await put(db.transaction(STORES.snapshots, "readwrite").objectStore(STORES.snapshots), {
    id, reason, createdAt: Date.now(), data: payload,
  });
  await pruneSnapshots();
  return id;
}

// Debounced, de-duplicated snapshot for routine edits. Called from the
// store's write hook — never call this directly for user-triggered saves
// like Restore or migration, which use createSnapshot() immediately instead.
function scheduleAutoSnapshot() {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(async () => {
    const db = dbInternal();
    if (!db || Store.isReadOnly()) return;
    const payload = buildUserDataPayload();
    const hash = await checksumOf(payload);
    const lastHash = await getMeta("lastSnapshotHash");
    if (hash === lastHash) return; // nothing actually changed — skip
    await put(db.transaction(STORES.snapshots, "readwrite").objectStore(STORES.snapshots), {
      id: `auto-${Date.now()}`, reason: "auto", createdAt: Date.now(), data: payload,
    });
    await setMeta("lastSnapshotHash", hash);
    await setMeta("lastSnapshotAt", Date.now());
    await pruneSnapshots();
  }, DEBOUNCE_MS);
}

Store.onWrite((key) => { if (key !== "schedule") scheduleAutoSnapshot(); });

// Call once, after Storage.ready(), from the app's bootstrap.
export async function checkPeriodicSnapshot() {
  if (Store.isReadOnly()) return;
  const now = Date.now();
  const lastDaily = (await getMeta("lastDailySnapshotAt")) || 0;
  const lastWeekly = (await getMeta("lastWeeklySnapshotAt")) || 0;
  if (now - lastWeekly > 7 * 864e5) { await createSnapshot("weekly"); await setMeta("lastWeeklySnapshotAt", now); }
  else if (now - lastDaily > 864e5) { await createSnapshot("daily"); await setMeta("lastDailySnapshotAt", now); }
}

export async function listSnapshots() {
  const db = dbInternal();
  if (!db) return [];
  const rows = await getAll(db.transaction(STORES.snapshots, "readonly").objectStore(STORES.snapshots));
  return rows.map((r) => ({ id: r.id, reason: r.reason, createdAt: r.createdAt })).sort((a, b) => b.createdAt - a.createdAt);
}

export async function restoreSnapshotById(id) {
  const db = dbInternal();
  const row = await new Promise((res) => {
    const r = db.transaction(STORES.snapshots, "readonly").objectStore(STORES.snapshots).get(id);
    r.onsuccess = () => res(r.result); r.onerror = () => res(null);
  });
  if (!row) return { ok: false, reason: "not-found" };
  // The one-time "pre-migration" snapshot stores the raw old localStorage
  // strings (legacyRaw), not a payload — convert it instead of crashing on
  // the missing .data.
  const data = row.data || (row.legacyRaw ? legacyRawToPayload(row.legacyRaw) : null);
  if (!data) return { ok: false, reason: "no-data" };
  return restoreBackup(data, { mode: "replace" });
}

function legacyRawToPayload(raw) {
  const parse = (v) => { if (v === null || v === undefined) return undefined; try { return JSON.parse(v); } catch { return undefined; } };
  const out = {};
  const profile = parse(raw.profile); if (profile !== undefined) out.profile = profile;
  const crewNames = parse(raw.crewNames); if (crewNames && typeof crewNames === "object") out.crewNames = crewNames;
  const access = parse(raw.dailyLogAccess); if (Array.isArray(access)) out.dailyLogAccess = access;
  const theme = parse(raw.theme); if (theme && theme.style) out.theme = theme;
  const log = parse(raw.dailyLogEntries);
  out.dailyLog = Array.isArray(log) ? log.filter((e) => e && e.id && e.date) : [];
  return out;
}

async function pruneSnapshots() {
  const db = dbInternal();
  const rows = await getAll(db.transaction(STORES.snapshots, "readonly").objectStore(STORES.snapshots));
  const byReason = {};
  for (const r of rows) (byReason[r.reason] ||= []).push(r);
  const toDelete = [];
  for (const [reason, list] of Object.entries(byReason)) {
    list.sort((a, b) => b.createdAt - a.createdAt);
    if (reason === "pre-migration" || reason === "pre-restore") {
      toDelete.push(...list.filter((r) => Date.now() - r.createdAt > RETENTION.guard));
    } else {
      const keep = RETENTION[reason] ?? RETENTION.auto;
      toDelete.push(...list.slice(keep));
    }
  }
  if (!toDelete.length) return;
  const store = db.transaction(STORES.snapshots, "readwrite").objectStore(STORES.snapshots);
  for (const r of toDelete) await del(store, r.id);
}

export async function daysSinceLastExport() {
  const last = await getMeta("lastExportAt");
  if (!last) return Infinity;
  return (Date.now() - last) / 864e5;
}
