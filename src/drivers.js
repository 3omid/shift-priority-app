// Drivers directory: the one list of people (name, Employee ID, full/part
// time, optional crew number, active or not) that every driver picker and
// Employee-ID lookup in the app reads from. Pure helpers only — saving goes
// through src/data/store.js like everything else.
//
// Driver shape: { id, name, employeeId, type: "full"|"part", crewNumber, active, createdAt }
// Drivers are never deleted, only deactivated, so old extra shifts and
// history that point at a driver id keep resolving.

export function newDriverId() {
  return `drv-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function makeDriver({ name, employeeId = "", type = "full", crewNumber = "", active = true }) {
  return {
    id: newDriverId(),
    name: String(name || "").trim(),
    employeeId: String(employeeId || "").trim(),
    type: type === "part" ? "part" : "full",
    crewNumber: String(crewNumber || "").trim(),
    active: active !== false,
    createdAt: Date.now(),
  };
}

// One-time move of the old per-crew name / Employee-ID maps into the
// directory: one driver per crew that has a name. A manual Admin entry wins
// over the built-in default list, same priority the app always used. A
// crew whose name was deliberately blanked is an open run, not a person.
export function seedDriversFromCrews({ crewNames = {}, crewEmployeeIds = {}, nameDefaults = {}, employeeIdDefaults = {} }) {
  const has = (o, k) => Object.prototype.hasOwnProperty.call(o || {}, k);
  const nums = new Set([...Object.keys(nameDefaults), ...Object.keys(crewNames), ...Object.keys(crewEmployeeIds)]);
  const out = [];
  const seen = new Set();
  const sorted = [...nums].sort((a, b) => (Number(a) - Number(b)) || a.localeCompare(b));
  for (const num of sorted) {
    const name = String(has(crewNames, num) ? crewNames[num] : nameDefaults[num] || "").trim();
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const employeeId = has(crewEmployeeIds, num) && crewEmployeeIds[num] ? crewEmployeeIds[num] : employeeIdDefaults[num] || "";
    out.push(makeDriver({ name, employeeId, type: "full", crewNumber: num }));
  }
  return out;
}

// The active driver currently on a crew (or null for an open run).
export function findDriverForCrew(drivers, crewNumber) {
  const key = String(crewNumber ?? "").trim();
  if (!key) return null;
  return (drivers || []).find((d) => d.active && String(d.crewNumber) === key) || null;
}

// Bulk add: one "Name, ID" per line. Blank lines are skipped. Returns the
// rows ready to add plus the lines that were skipped and why
// ("format" = not "Name, ID"; "duplicate" = that Employee ID already exists
// in the directory or earlier in the paste).
export function parseBulkDrivers(text, existing = []) {
  const ids = new Set((existing || []).map((d) => String(d.employeeId || "").trim()).filter(Boolean));
  const rows = [];
  const errors = [];
  String(text || "").split(/\r?\n/).forEach((raw, i) => {
    const line = raw.trim();
    if (!line) return;
    const m = line.match(/^(.+?)\s*[,;\t]\s*([A-Za-z0-9-]+)\s*$/);
    if (!m || !m[1].trim()) { errors.push({ line: i + 1, text: line, reason: "format" }); return; }
    const name = m[1].trim();
    const employeeId = m[2].trim();
    if (ids.has(employeeId)) { errors.push({ line: i + 1, text: line, reason: "duplicate" }); return; }
    ids.add(employeeId);
    rows.push({ name, employeeId });
  });
  return { rows, errors };
}

// Search helper for driver pickers: name, Employee ID or crew number.
export function matchDriver(d, query) {
  const q = String(query || "").trim().toLowerCase();
  if (!q) return true;
  return d.name.toLowerCase().includes(q) || String(d.employeeId).toLowerCase().includes(q) || String(d.crewNumber) === q;
}
