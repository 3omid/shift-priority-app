import { tx, put, get } from "./db.js";
import { STORES, LEGACY_KEYS, DEFAULTS, CURRENT_SCHEMA_VERSION } from "./schema.js";

function safeParse(raw) {
  if (raw === null || raw === undefined) return { ok: true, value: undefined };
  try { return { ok: true, value: JSON.parse(raw) }; }
  catch { return { ok: false, raw }; }
}

// Reads every legacy key exactly as it is on disk. Never throws, never
// deletes anything — this is purely a read.
function readLegacy() {
  const out = {};
  for (const [name, key] of Object.entries(LEGACY_KEYS)) {
    let raw = null;
    try { raw = localStorage.getItem(key); } catch { /* storage unavailable */ }
    out[name] = { key, raw, parsed: safeParse(raw) };
  }
  return out;
}

// One-time legacy -> v1 migration. Idempotent: if meta.schemaVersion is
// already set, this is a no-op. Never deletes the legacy localStorage keys.
export async function migrateLegacyIfNeeded(db) {
  const existing = await get(await storeHandle(db, STORES.meta), "schemaVersion").catch(() => null);
  if (existing && existing.value >= 1) return { ranMigration: false };

  const legacy = readLegacy();
  const corrupt = Object.entries(legacy).filter(([, v]) => v.raw !== null && !v.parsed.ok).map(([k]) => k);

  // Snapshot the raw legacy state first, before writing anything new, so a
  // botched migration can always be inspected/recovered from later.
  const snapshotId = `pre-migration-${Date.now()}`;
  const rawBundle = Object.fromEntries(Object.entries(legacy).map(([k, v]) => [k, v.raw]));

  let dailyLogCount = 0, dailyLogWritten = 0;

  await tx(db, [STORES.kv, STORES.dailyLog, STORES.schedule, STORES.snapshots, STORES.meta], "readwrite", async (t) => {
    const kv = t.objectStore(STORES.kv);
    const log = t.objectStore(STORES.dailyLog);
    const sched = t.objectStore(STORES.schedule);
    const snaps = t.objectStore(STORES.snapshots);
    const meta = t.objectStore(STORES.meta);

    await put(snaps, { id: snapshotId, reason: "pre-migration", createdAt: Date.now(), legacyRaw: rawBundle });

    const profile = legacy.profile.parsed.ok ? legacy.profile.parsed.value : DEFAULTS.profile;
    await put(kv, { key: "profile", value: profile ?? DEFAULTS.profile });

    const crewNames = legacy.crewNames.parsed.ok ? legacy.crewNames.parsed.value : DEFAULTS.crewNames;
    await put(kv, { key: "crewNames", value: crewNames || DEFAULTS.crewNames });

    const dailyLogAccess = legacy.dailyLogAccess.parsed.ok ? legacy.dailyLogAccess.parsed.value : DEFAULTS.dailyLogAccess;
    await put(kv, { key: "dailyLogAccess", value: Array.isArray(dailyLogAccess) ? dailyLogAccess : DEFAULTS.dailyLogAccess });

    const theme = legacy.theme.parsed.ok ? legacy.theme.parsed.value : DEFAULTS.theme;
    await put(kv, { key: "theme", value: theme || DEFAULTS.theme });

    await put(kv, { key: "sound", value: DEFAULTS.sound });
    await put(kv, { key: "lang", value: DEFAULTS.lang });

    const entries = legacy.dailyLogEntries.parsed.ok && Array.isArray(legacy.dailyLogEntries.parsed.value)
      ? legacy.dailyLogEntries.parsed.value : [];
    dailyLogCount = entries.length;
    for (const e of entries) {
      if (!e || !e.id || !e.date) continue;
      await put(log, { ...e, updatedAt: e.updatedAt || Date.now() });
      dailyLogWritten++;
    }

    if (legacy.lastFile.parsed.ok && legacy.lastFile.parsed.value) {
      await put(sched, { key: "current", ...legacy.lastFile.parsed.value, savedAt: Date.now() });
    }

    // Only mark the migration done once every write above has succeeded —
    // this line is what makes the whole migration "commit".
    await put(meta, { key: "schemaVersion", value: CURRENT_SCHEMA_VERSION });
    await put(meta, { key: "migratedAt", value: Date.now() });
  });

  return { ranMigration: true, dailyLogCount, dailyLogWritten, corrupt, snapshotId };
}

async function storeHandle(db, name) {
  return await new Promise((resolve, reject) => {
    try {
      const t = db.transaction(name, "readonly");
      resolve(t.objectStore(name));
    } catch (e) { reject(e); }
  });
}

// Used only when migration itself throws, so the person still sees their
// real data (read-only) instead of a blank app while we investigate.
export function readLegacyForDisplay() {
  const legacy = readLegacy();
  const val = (name, fallback) => (legacy[name].parsed.ok ? legacy[name].parsed.value ?? fallback : fallback);
  return {
    profile: val("profile", DEFAULTS.profile),
    crewNames: val("crewNames", DEFAULTS.crewNames),
    dailyLogAccess: val("dailyLogAccess", DEFAULTS.dailyLogAccess),
    theme: val("theme", DEFAULTS.theme),
    dailyLogEntries: val("dailyLogEntries", []),
    lastFile: val("lastFile", null),
  };
}

// Placeholder for the next schema bump:
// export async function migrateV1ToV2(db) { ... }
