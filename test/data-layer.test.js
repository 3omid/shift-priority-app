// npm install (fake-indexeddb + vitest are already in devDependencies), then:
//   npm test
import "fake-indexeddb/auto";
import { describe, it, expect } from "vitest";

// Vitest's default "node" environment has no localStorage — this suite only
// needs get/set/removeItem, so a tiny in-memory stand-in is enough (no need
// to pull in jsdom/happy-dom just for this).
if (typeof localStorage === "undefined") {
  const backing = new Map();
  globalThis.localStorage = {
    getItem: (k) => (backing.has(k) ? backing.get(k) : null),
    setItem: (k, v) => backing.set(k, String(v)),
    removeItem: (k) => backing.delete(k),
    clear: () => backing.clear(),
  };
}
import * as Store from "../src/data/store.js";
import * as Backup from "../src/data/backup.js";

// fake-indexeddb keeps its databases in a module-level map, so each test
// needs its own database name to stay isolated. store.js reads DB_NAME from
// schema.js at import time, so we can't easily rename it per-test without
// restructuring the modules — instead, tests run in one shared sequence
// (matching how a single browser tab actually behaves) and each block
// builds on the previous one's state, same as manual testing on a real page.
function seedLegacyLocalStorage() {
  localStorage.setItem("shiftPriorityProfile", JSON.stringify({ firstName: "Omid", crewNumber: "40" }));
  localStorage.setItem("shiftPriorityCrewNames", JSON.stringify({ "40": "Omid Farhadnia", "10": "Elahe Alamdar" }));
  localStorage.setItem("shiftPriorityTheme", JSON.stringify({ style: "sunset", mode: "dark" }));
  localStorage.setItem("shiftPriorityDailyLogEntries", JSON.stringify([
    { id: "1", date: "2026-09-01", startTime: "13:29", endTime: "21:10", startYard: "Newmarket", endYard: "Newmarket", description: "", totalHours: 7.68 },
    { id: "2", date: "2026-09-02", startTime: "13:30", endTime: "23:30", startYard: "Caldari", endYard: "Caldari", description: "", totalHours: 10 },
  ]));
  localStorage.setItem("shiftPriorityDailyLogAccess", "not valid json {{{");
}

describe("data layer: migration + storage (run in order, one shared IndexedDB)", () => {
  it("migrates legacy localStorage into IndexedDB on first init, without deleting it", async () => {
    seedLegacyLocalStorage();
    const report = await Store.init();
    expect(report.ranMigration).toBe(true);
    expect(Store.loadProfile()?.firstName).toBe("Omid");
    expect(Store.loadCrewNames()["10"]).toBe("Elahe Alamdar");
    expect(Store.loadThemePrefs().style).toBe("sunset");
    expect(Store.loadDailyLogEntries()).toHaveLength(2);
    expect(Store.loadDailyLogAccess()).toEqual([]); // corrupt legacy value -> default, not a crash
    expect(report.corrupt).toContain("dailyLogAccess");
    expect(localStorage.getItem("shiftPriorityProfile")).not.toBeNull();
  });

  it("does not re-run migration, and writes persist across a re-hydrate", async () => {
    await Store.init();
    expect(Store.getInitReport().ranMigration).toBe(false);
    await Store.saveProfile({ firstName: "Omid", crewNumber: "41" });
    await Store.init();
    expect(Store.loadProfile().crewNumber).toBe("41");
  });

  it("saveDailyLogEntries with a shorter array deletes the missing row", async () => {
    const list = Store.loadDailyLogEntries();
    await Store.saveDailyLogEntries(list.filter((e) => e.id !== "2"));
    await Store.init();
    expect(Store.loadDailyLogEntries().map((e) => e.id)).toEqual(["1"]);
  });

  it("refuses writes and clears once schema is back in range (downgrade guard)", async () => {
    const { db } = Store._internal();
    await new Promise((res) => {
      const t = db.transaction("meta", "readwrite");
      t.objectStore("meta").put({ key: "schemaVersion", value: 999 });
      t.oncomplete = res;
    });
    await Store.init();
    expect(Store.isReadOnly()).toBe(true);
    expect(Store.readOnlyReasonCode()).toBe("downgrade");
    const before = Store.loadProfile();
    await Store.saveProfile({ firstName: "nope" });
    expect(Store.loadProfile()).toEqual(before);

    await new Promise((res) => {
      const t = db.transaction("meta", "readwrite");
      t.objectStore("meta").put({ key: "schemaVersion", value: 1 });
      t.oncomplete = res;
    });
    await Store.init();
    expect(Store.isReadOnly()).toBe(false);
  });

  it("rejects a tampered or too-new backup file, and restores a valid one", async () => {
    const file = await Backup.buildBackupFile();
    const json = JSON.stringify(file);
    const v1 = await Backup.validateBackupFile(json);
    expect(v1.ok).toBe(true);

    const tampered = JSON.parse(json);
    tampered.data.profile.crewNumber = "999";
    expect((await Backup.validateBackupFile(JSON.stringify(tampered))).reason).toBe("checksum-mismatch");

    const tooNew = JSON.parse(json);
    tooNew.schemaVersion = 999;
    expect((await Backup.validateBackupFile(JSON.stringify(tooNew))).reason).toBe("too-new");

    await Store.saveProfile({ firstName: "temp", crewNumber: "0" });
    await Backup.restoreBackup(v1.data, { mode: "replace" });
    await Store.init();
    expect(Store.loadProfile().crewNumber).toBe("41");
    expect((await Backup.listSnapshots()).some((s) => s.reason === "pre-restore")).toBe(true);
  });

  it("mergeDailyLog keeps whichever side was edited more recently, per date", async () => {
    const current = Store.loadDailyLogEntries();
    const id = current[0].id;
    await Store.saveDailyLogEntries(current.map((e) => e.id === id ? { ...e, description: "local edit", updatedAt: Date.now() + 1e5 } : e));

    await Backup.restoreBackup({
      dailyLog: [
        { id, date: current[0].date, description: "stale from backup", totalHours: 1, updatedAt: 1 },
        { id: "new-from-backup", date: "2026-09-20", description: "only in backup", totalHours: 8, updatedAt: Date.now() },
      ],
    }, { mode: "mergeDailyLog" });

    await Store.init();
    const merged = Store.loadDailyLogEntries();
    expect(merged.find((e) => e.id === id).description).toBe("local edit");
    expect(merged.some((e) => e.id === "new-from-backup")).toBe(true);
  });

  it("backups include the admin Employee-ID directory and restore it", async () => {
    await Store.saveCrewEmployeeIds({ "40": "1599" });
    const file = await Backup.buildBackupFile();
    expect(file.data.crewEmployeeIds).toEqual({ "40": "1599" });
    await Store.saveCrewEmployeeIds({});
    const res = await Backup.restoreBackup(file.data, { mode: "replace" });
    expect(res.ok).toBe(true);
    await Store.init();
    expect(Store.loadCrewEmployeeIds()["40"]).toBe("1599");
  });

  it("restores the pre-migration snapshot (raw legacy data) without crashing", async () => {
    const snap = (await Backup.listSnapshots()).find((s) => s.reason === "pre-migration");
    expect(snap).toBeTruthy();
    await Store.saveProfile({ firstName: "changed", crewNumber: "1" });
    const res = await Backup.restoreSnapshotById(snap.id);
    expect(res.ok).toBe(true);
    await Store.init();
    expect(Store.loadProfile()).toEqual({ firstName: "Omid", crewNumber: "40" });
    expect(Store.loadCrewNames()["10"]).toBe("Elahe Alamdar");
    expect(Store.loadDailyLogEntries().map((e) => e.date).sort()).toEqual(["2026-09-01", "2026-09-02"]);
  });
});

describe("data layer: drivers directory, extra shifts, admin remember flag", () => {
  it("persist across a re-hydrate and ride along in backups (except the remember flag)", async () => {
    await Store.init();
    expect(Store.loadDrivers()).toBeNull(); // never set up -> app seeds it once
    await Store.saveDrivers([{ id: "d1", name: "Omid", employeeId: "1599", type: "full", crewNumber: "40", active: true }]);
    await Store.saveExtraShifts([{ id: "s1", driverId: "d1", date: "2099-01-01", run: "V7", start: "13:22", end: "" }]);
    await Store.saveAdminRemember(true);
    await Store.init();
    expect(Store.loadDrivers()[0].name).toBe("Omid");
    expect(Store.loadExtraShifts()[0].run).toBe("V7");
    expect(Store.loadAdminRemember()).toBe(true);
    const payload = Backup.buildUserDataPayload();
    expect(payload.drivers).toHaveLength(1);
    expect(payload.extraShifts).toHaveLength(1);
    expect(payload).not.toHaveProperty("adminRemember");
    await Store.saveAdminRemember(false);
    expect(Store.loadAdminRemember()).toBe(false);
  });
});
