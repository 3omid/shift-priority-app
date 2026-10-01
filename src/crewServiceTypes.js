// Service type of each crew: "OR" (On Request — regular riders) or "MOR"
// (Mobility On Request — riders with a disability, mobility needs or
// seniors). Admin-maintained per crew number; stored as
// { [crewNumber]: "OR" | "MOR" } through src/data/store.js.
// Pure helpers only, so the text import can be unit-tested.

export const SERVICE_TYPES = ["OR", "MOR"];

// "MOR", "Mobility", "Mobility On Request" -> MOR; "OR", "On Request" -> OR.
// MOR is checked first because "Mobility On Request" also contains "On Request".
export function detectServiceType(text) {
  const s = String(text || "");
  if (/\bM\.?O\.?R\b|\bmobility\b|موبیلیتی|مبیلیتی/i.test(s)) return "MOR";
  if (/\bO\.?R\b|\bon[\s-]*request\b|آن\s*ریکوئست|آن‌ریکوئست/i.test(s)) return "OR";
  return null;
}

// Crew numbers in one line: "21", "21, 22, 23", "21-25", "21 to 25",
// "Crew 21". Times like "5:22" and long numbers (bus/employee #) are ignored.
function crewNumbersIn(line) {
  const nums = [];
  const cleaned = line.replace(/\d{1,2}:\d{2}/g, " ");
  const re = /(?<![\d.])(\d{1,3})(?:\s*(?:-|–|to|تا)\s*(\d{1,3}))?(?![\d.])/gi;
  for (const m of cleaned.matchAll(re)) {
    const a = Number(m[1]);
    const b = m[2] !== undefined ? Number(m[2]) : a;
    if (b >= a && b - a <= 100) for (let n = a; n <= b; n++) nums.push(String(n));
    else nums.push(String(a));
  }
  return nums;
}

// Bulk import: one or more crews and a type per line, in any order, e.g.
//   21 Newmarket MOR
//   Crew 5 - On Request
//   30-35 MOR
//   12<TAB>Richmond Hill<TAB>OR      (pasted from a spreadsheet)
// Returns { updates: { [crew]: type }, errors: [{ line, text, reason }] }
// where reason is "noType" or "noCrew". A later line wins for the same crew.
export function parseServiceTypeText(text) {
  const updates = {};
  const errors = [];
  String(text || "").split(/\r?\n/).forEach((raw, i) => {
    const line = raw.trim();
    if (!line) return;
    const type = detectServiceType(line);
    if (!type) { errors.push({ line: i + 1, text: line, reason: "noType" }); return; }
    const nums = crewNumbersIn(line);
    if (!nums.length) { errors.push({ line: i + 1, text: line, reason: "noCrew" }); return; }
    for (const n of nums) updates[n] = type;
  });
  return { updates, errors };
}

export function serviceTypeOf(map, crewNumber) {
  const v = map ? map[String(crewNumber)] : null;
  return SERVICE_TYPES.includes(v) ? v : null;
}
