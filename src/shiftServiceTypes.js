// Service type of each SHIFT (a run on a day, e.g. run "21" out of the
// Newmarket yard): "OR" (On Request — regular riders) or "MOR" (Mobility On
// Request — riders with a disability, a physical or mental condition, or
// seniors). It belongs to the shift, not to the crew: the same crew can work
// MOR runs on some days and OR runs on others, so a crew's week shows the
// type day by day.
//
// Stored as { [shiftKey]: "OR" | "MOR" } through src/data/store.js, where
// shiftKey = "<RUNCODE>@<regionKey>" (e.g. "21@NMK"), or "<RUNCODE>@*" when
// the yard wasn't given (matches that run code at any yard).
// Pure helpers only, so the text import can be unit-tested.

import { LOCATIONS } from "./extraShiftParser.js";

export const SERVICE_TYPES = ["OR", "MOR"];
export const ANY_YARD = "*";

// "PRO 9 MRC" -> "PRO9MRC", "run 14" -> "14", "ru15" -> "15", " v7 " -> "V7".
export function normalizeRunCode(code) {
  let s = String(code ?? "").toUpperCase().replace(/[\s._#-]+/g, "");
  s = s.replace(/^(?:RUN|RU|RN)(?=\d)/, "");
  return s;
}

export function shiftKey(code, regionKey) {
  const c = normalizeRunCode(code);
  if (!c) return "";
  return `${c}@${regionKey && regionKey !== "OTHER" ? regionKey : ANY_YARD}`;
}

// The type of one day's shift: an exact run+yard entry first, then the
// same run code saved without a yard.
export function serviceTypeForShift(map, code, regionKey) {
  if (!map) return null;
  const c = normalizeRunCode(code);
  if (!c) return null;
  const exact = regionKey ? map[`${c}@${regionKey}`] : undefined;
  const v = SERVICE_TYPES.includes(exact) ? exact : map[`${c}@${ANY_YARD}`];
  return SERVICE_TYPES.includes(v) ? v : null;
}

// "MOR", "Mobility", "Mobility On Request" -> MOR; "OR", "On Request" -> OR.
// MOR is checked first because "Mobility On Request" also contains "On Request".
const MOR_RE = /\bmobility(?:\s+on[\s-]*request)?\b|\bM\.?O\.?R\b|موبیلیتی|مبیلیتی/gi;
const OR_RE = /\bon[\s-]*request\b|\bO\.?R\b|آن\s*ریکوئست|آن‌ریکوئست/gi;
export function detectServiceType(text) {
  const s = String(text || "");
  if (new RegExp(MOR_RE.source, "i").test(s)) return "MOR";
  if (new RegExp(OR_RE.source, "i").test(s)) return "OR";
  return null;
}

// Bulk import, one line per run (or list/range of runs), any order:
//   21 Newmarket MOR
//   Crew 12 Newmarket On Request
//   PRO 9 MRC, BRT, MOR
//   30-35 Newmarket MOR
//   V7<TAB>Newmarket<TAB>OR          (pasted from a spreadsheet)
// A line without a yard applies to that run code at any yard.
// Returns { updates: { [shiftKey]: type }, errors: [{ line, text, reason }] },
// reason "noType" or "noRun". A later line wins for the same shift.
export function parseServiceTypeText(text) {
  const updates = {};
  const errors = [];
  String(text || "").split(/\r?\n/).forEach((raw, i) => {
    const line = raw.trim();
    if (!line) return;
    const type = detectServiceType(line);
    if (!type) { errors.push({ line: i + 1, text: line, reason: "noType" }); return; }
    let rest = line.replace(new RegExp(MOR_RE.source, "gi"), " ").replace(new RegExp(OR_RE.source, "gi"), " ");
    let regionKey = null;
    for (const loc of LOCATIONS) {
      const re = new RegExp(loc.re.source, "gi");
      if (re.test(rest)) { regionKey = regionKey || loc.regionKey; rest = rest.replace(re, " "); }
    }
    rest = rest.replace(/\b(?:crew|shift|block)\b|کرو|شیفت|ران/gi, " ");
    const codes = [];
    for (const part of rest.split(/[,;\t|/]+/)) {
      const p = part.replace(/[:\-–—\s]+$/, "").replace(/^[:\-–—\s]+/, "").trim();
      if (!p) continue;
      const range = p.match(/^(\d{1,3})\s*(?:-|–|to|تا)\s*(\d{1,3})$/i);
      if (range && Number(range[2]) >= Number(range[1]) && Number(range[2]) - Number(range[1]) <= 100) {
        for (let n = Number(range[1]); n <= Number(range[2]); n++) codes.push(String(n));
      } else if (/^\d+(?:\s+\d+)+$/.test(p)) {
        codes.push(...p.split(/\s+/)); // "21 22 23"
      } else {
        codes.push(p.replace(/^-\s*|\s*-$/g, ""));
      }
    }
    const keys = codes.map((c) => shiftKey(c, regionKey)).filter(Boolean);
    if (!keys.length) { errors.push({ line: i + 1, text: line, reason: "noRun" }); return; }
    for (const k of keys) updates[k] = type;
  });
  return { updates, errors };
}

// "21@NMK" -> { code: "21", regionKey: "NMK" }
export function splitShiftKey(key) {
  const at = key.lastIndexOf("@");
  return { code: key.slice(0, at), regionKey: key.slice(at + 1) };
}
