// Single source of truth for the database name/version and every default
// value. Bump CURRENT_SCHEMA_VERSION and add a step in migrations.js
// whenever the *shape* of stored data changes — never for a bug fix that
// doesn't change the shape.

export const DB_NAME = "shift-priority"; // unique even though GitHub Pages
export const DB_VERSION = 1; // shares an origin with every other 3omid.github.io app
export const CURRENT_SCHEMA_VERSION = 1;

export const STORES = {
  kv: "kv", // keyPath: "key" — profile, settings, crewNames, dailyLogAccess
  dailyLog: "dailyLog", // keyPath: "id", index "date"
  schedule: "schedule", // keyPath: "key" — "current" | "previous"
  snapshots: "snapshots", // keyPath: "id" (timestamp string), index "reason"
  meta: "meta", // keyPath: "key" — schemaVersion, lastSnapshotAt, lastExportAt
};

export function upgrade(db, oldVersion) {
  if (oldVersion < 1) {
    db.createObjectStore(STORES.kv, { keyPath: "key" });
    const log = db.createObjectStore(STORES.dailyLog, { keyPath: "id" });
    log.createIndex("date", "date", { unique: true });
    db.createObjectStore(STORES.schedule, { keyPath: "key" });
    const snaps = db.createObjectStore(STORES.snapshots, { keyPath: "id" });
    snaps.createIndex("reason", "reason", { unique: false });
    db.createObjectStore(STORES.meta, { keyPath: "key" });
  }
  // if (oldVersion < 2) { ... add v1->v2 store/index changes here ... }
}

export const DEFAULTS = {
  profile: null, // { firstName, crewNumber, employeeId }
  crewNames: {}, // { [crewNumber]: name }
  crewEmployeeIds: {}, // { [crewNumber]: employeeId } -- admin-maintained, for the Shift Exchange form
  dailyLogAccess: [], // [crewNumber, ...]
  theme: { style: "universal", mode: "light" },
  sound: { enabled: false, type: "keyboard", pitch: 1 },
  lang: "en",
};

// localStorage keys this app has ever used, and where each one lands in the
// new store. Nothing here is ever deleted — see migrations.js.
export const LEGACY_KEYS = {
  profile: "shiftPriorityProfile",
  crewNames: "shiftPriorityCrewNames",
  dailyLogEntries: "shiftPriorityDailyLogEntries",
  dailyLogAccess: "shiftPriorityDailyLogAccess",
  theme: "shiftPriorityTheme",
  lastFile: "shiftPriorityLastFile",
};
