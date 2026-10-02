// Parses the free-text messages dispatch sends about extra/switched shifts
// ("shift switches for next week, Tuesday Sept.29 Pro8A BRT 5:22-9:37 and
// Thursday Oct.1 23STF BRT 4:52-15:07") into one row per shift:
//   { date, dateText, weekday, run, location, regionKey, start, end, flags }
// Pure functions only (no React, no storage) so it can be unit-tested.
// Nothing here is trusted blindly: anything doubtful is returned with a flag
// for the admin to check in the preview table before it is saved.

const MONTHS = {
  jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3, may: 4,
  jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7, sep: 8, sept: 8, september: 8,
  oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
};
const WEEKDAYS = {
  sun: 0, sunday: 0, mon: 1, monday: 1, tue: 2, tues: 2, tuesday: 2, wed: 3, wednesday: 3,
  thu: 4, thur: 4, thurs: 4, thursday: 4, fri: 5, friday: 5, sat: 6, saturday: 6,
};

// Yard names as dispatch writes them -> the app's region keys (see
// REGION_LABELS in App.jsx). "BRT" is the Richmond Hill yard.
export const LOCATIONS = [
  { re: /\bBRT\b/i, location: "Richmond Hill", regionKey: "RH" },
  { re: /\brichmond\s*hill\b|\bRH\b/i, location: "Richmond Hill", regionKey: "RH" },
  { re: /\bnew\s?market\b|\bNMK\b/i, location: "Newmarket", regionKey: "NMK" },
  { re: /\bcaldari\b/i, location: "Caldari", regionKey: "CLDR" },
  { re: /\bstouffville\b/i, location: "Stouffville", regionKey: "STF" },
  { re: /\bmaple\b/i, location: "Maple", regionKey: "MRG" },
];

export const FLAGS = {
  invalidStart: "invalidStart",
  invalidEnd: "invalidEnd",
  missingStart: "missingStart",
  missingEnd: "missingEnd",
  missingDate: "missingDate",
  relativeDate: "relativeDate",
  weekdayMismatch: "weekdayMismatch",
  pastDate: "pastDate",
};

// ---- date helpers (all local calendar dates as "YYYY-MM-DD", no TZ math) ----
export function ymd(y, m, d) {
  return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}
function parseYmd(s) {
  const [y, m, d] = s.split("-").map(Number);
  return { y, m: m - 1, d };
}
function toUtc(s) { const { y, m, d } = parseYmd(s); return Date.UTC(y, m, d); }
export function addDays(s, n) {
  const dt = new Date(toUtc(s) + n * 864e5);
  return ymd(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
}
export function weekdayOf(s) { return new Date(toUtc(s)).getUTCDay(); }
function isRealDate(y, m, d) {
  const dt = new Date(Date.UTC(y, m, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m && dt.getUTCDate() === d;
}
export function todayStr(now = new Date()) {
  return ymd(now.getFullYear(), now.getMonth(), now.getDate());
}

// ---- times ----
// "5:22", "13:52", "5.22", "1322" -> { text: "05:22", valid } (valid=false
// keeps the as-typed text, e.g. "5:62", so the admin sees what was sent).
export function normalizeTime(raw) {
  const s = String(raw || "").trim();
  if (!s) return { text: "", valid: false, empty: true };
  const m = s.match(/^(\d{1,2})[:.h]?(\d{2})$/);
  if (!m) return { text: s, valid: false };
  const h = Number(m[1]), mm = Number(m[2]);
  if (h > 23 || mm > 59) return { text: s, valid: false };
  return { text: `${String(h).padStart(2, "0")}:${String(mm).padStart(2, "0")}`, valid: true };
}

// Hours between two "HH:MM" strings; an end at/before the start runs past midnight.
export function shiftHours(start, end) {
  const a = normalizeTime(start), b = normalizeTime(end);
  if (!a.valid || !b.valid) return null;
  const toMin = (t) => Number(t.slice(0, 2)) * 60 + Number(t.slice(3));
  let diff = toMin(b.text) - toMin(a.text);
  if (diff <= 0) diff += 1440;
  return Math.round((diff / 60) * 100) / 100;
}

const MONTH_RE = "(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sept(?:ember)?|sep|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)";
const WEEKDAY_RE = "(sun(?:day)?|mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?)";
// [weekday[,]] month[.] day[st|nd|rd|th][,] [year]
const ABS_DATE_RE = new RegExp(`(?:\\b${WEEKDAY_RE}\\b[,\\s]*)?\\b${MONTH_RE}\\b\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\b(?:,?\\s*(20\\d{2})\\b)?`, "gi");
// [weekday] day month  (e.g. "Tuesday 29 Sept")
const ABS_DATE_DM_RE = new RegExp(`(?:\\b${WEEKDAY_RE}\\b[,\\s]*)?\\b(\\d{1,2})(?:st|nd|rd|th)?\\s+${MONTH_RE}\\b\\.?(?:,?\\s*(20\\d{2})\\b)?`, "gi");
const REL_DATE_RE = new RegExp(`\\b(today|tonight|tomorrow|tmrw|tmr)\\b(?:[,\\s]+${WEEKDAY_RE}\\b)?`, "gi");
const WEEKDAY_ONLY_RE = new RegExp(`\\b(?:(next|this)\\s+)?${WEEKDAY_RE}\\b`, "gi");

// Every date mentioned in `text`, with its character span.
export function findDates(text, today) {
  const found = [];
  const taken = [];
  const free = (s, e) => !taken.some(([a, b]) => s < b && e > a);
  const push = (m, info) => {
    const s = m.index, e = m.index + m[0].length;
    if (!free(s, e)) return;
    taken.push([s, e]);
    found.push({ start: s, end: e, text: m[0].trim(), ...info });
  };

  const resolveAbs = (wdWord, monWord, dayNum, yearNum) => {
    const month = MONTHS[monWord.toLowerCase()];
    const day = Number(dayNum);
    const t = parseYmd(today);
    const flags = [];
    let y = yearNum ? Number(yearNum) : t.y;
    if (!isRealDate(y, month, day)) return null;
    let date = ymd(y, month, day);
    if (!yearNum && date < today) {
      // No year given: the nearest upcoming date.
      y += 1;
      if (!isRealDate(y, month, day)) return null;
      date = ymd(y, month, day);
    }
    if (yearNum && date < today) flags.push(FLAGS.pastDate);
    const weekday = wdWord ? WEEKDAYS[wdWord.toLowerCase()] : null;
    if (weekday !== null && weekday !== undefined && weekdayOf(date) !== weekday) flags.push(FLAGS.weekdayMismatch);
    return { date, relative: false, flags };
  };

  for (const m of text.matchAll(ABS_DATE_RE)) {
    const r = resolveAbs(m[1], m[2], m[3], m[4]);
    if (r) push(m, r);
  }
  for (const m of text.matchAll(ABS_DATE_DM_RE)) {
    const r = resolveAbs(m[1], m[3], m[2], m[4]);
    if (r) push(m, r);
  }
  for (const m of text.matchAll(REL_DATE_RE)) {
    const word = m[1].toLowerCase();
    const date = word === "today" || word === "tonight" ? today : addDays(today, 1);
    const flags = [FLAGS.relativeDate];
    const weekday = m[2] ? WEEKDAYS[m[2].toLowerCase()] : null;
    if (weekday !== null && weekday !== undefined && weekdayOf(date) !== weekday) flags.push(FLAGS.weekdayMismatch);
    push(m, { date, relative: true, flags });
  }
  for (const m of text.matchAll(WEEKDAY_ONLY_RE)) {
    // A bare weekday ("Wednesday") -> the next such day from today (today
    // itself counts); "next Wednesday" -> the one in the following week.
    const weekday = WEEKDAYS[m[2].toLowerCase()];
    let diff = (weekday - weekdayOf(today) + 7) % 7;
    if (m[1] && m[1].toLowerCase() === "next" && diff === 0) diff = 7;
    push(m, { date: addDays(today, diff), relative: true, flags: [FLAGS.relativeDate] });
  }
  return found.sort((a, b) => a.start - b.start);
}

// Times: "5:22-9:37", "13:52 to 22:07", or a lone start ("1322", "5:22").
const TIME_TOKEN = "(\\d{1,2}[:.]\\d{2}|\\d{4})";
const TIME_RE = new RegExp(`(?<![\\d:.])${TIME_TOKEN}(?:\\s*(?:-|–|—|to|till|until)\\s*${TIME_TOKEN})?(?![\\d:])`, "gi");

function findTimes(text) {
  const out = [];
  for (const m of text.matchAll(TIME_RE)) {
    // "Bus 25005" never matches (5 digits); "Bus 2500" would, so skip a
    // 4-digit number right after "bus"/"#"/"unit".
    const before = text.slice(Math.max(0, m.index - 8), m.index);
    if (/(?:bus|unit|#|no\.?)\s*$/i.test(before)) continue;
    out.push({ start: m.index, end: m.index + m[0].length, startRaw: m[1], endRaw: m[2] || "" });
  }
  return out;
}

// Run codes: "run 14", "ru15", "run V7", or a stand-alone code mixing
// letters and digits ("Pro8A", "23STF", "V7").
const RUN_PREFIX_RE = /\b(?:run|ru|rn)\s*#?\s*([a-z]*\d+[a-z0-9]*)\b/gi;
const CODE_RE = /\b(?=[a-z0-9]*[a-z])(?=[a-z0-9]*\d)[a-z0-9]{2,8}\b/gi;

function formatRun(code) {
  const c = code.toUpperCase();
  return /^\d+$/.test(c) ? `RUN ${c}` : c;
}

function findRuns(text) {
  const out = [];
  const taken = [];
  for (const m of text.matchAll(RUN_PREFIX_RE)) {
    out.push({ start: m.index, end: m.index + m[0].length, run: formatRun(m[1]) });
    taken.push([m.index, m.index + m[0].length]);
  }
  for (const m of text.matchAll(CODE_RE)) {
    const s = m.index, e = s + m[0].length;
    if (taken.some(([a, b]) => s < b && e > a)) continue;
    if (/^\d+(?:st|nd|rd|th)$/i.test(m[0])) continue; // ordinal, not a run
    const before = text.slice(Math.max(0, s - 8), s);
    if (/(?:bus|unit)\s*#?\s*$/i.test(before)) continue;
    out.push({ start: s, end: e, run: formatRun(m[0]) });
  }
  return out.sort((a, b) => a.start - b.start);
}

function findLocations(text) {
  const out = [];
  for (const loc of LOCATIONS) {
    const re = new RegExp(loc.re.source, "gi");
    for (const m of text.matchAll(re)) out.push({ start: m.index, end: m.index + m[0].length, location: loc.location, regionKey: loc.regionKey });
  }
  return out.sort((a, b) => a.start - b.start);
}

// Blank out spans (keeping positions) so later scans can't re-read them.
function mask(text, spans) {
  let out = text;
  for (const { start, end } of spans) out = out.slice(0, start) + " ".repeat(end - start) + out.slice(end);
  return out;
}

// Splits a message into clauses at " and ", ";", new lines and sentence
// ends, so "Tuesday ... 5:22-9:37 and Thursday ... 4:52-15:07" keeps each
// shift with its own date/run/yard.
function clauseBounds(text) {
  const cuts = [0];
  const re = /\band\b|\balso\b|[;\n]|[.!?](?=\s+[A-Z])|,\s*(?=(?:sun|mon|tue|wed|thu|fri|sat)[a-z]*\b)/gi;
  for (const m of text.matchAll(re)) cuts.push(m.index);
  cuts.push(text.length);
  const out = [];
  for (let i = 0; i < cuts.length - 1; i++) if (cuts[i + 1] > cuts[i]) out.push([cuts[i], cuts[i + 1]]);
  return out;
}

function inside(tok, [a, b]) { return tok.start >= a && tok.start < b; }
function nearest(list, pos) {
  let best = null, bestD = Infinity;
  for (const x of list) {
    const d = Math.abs(x.start - pos);
    if (d < bestD) { best = x; bestD = d; }
  }
  return best;
}

// Main entry. `today` is "YYYY-MM-DD" (defaults to the device's today).
export function parseDispatchMessage(text, today = todayStr()) {
  const src = String(text || "");
  if (!src.trim()) return [];

  const dates = findDates(src, today);
  const noDates = mask(src, dates);
  const times = findTimes(noDates);
  const runs = findRuns(mask(noDates, times));
  const locs = findLocations(src);

  // Each time (or time range) is one shift. With no times at all, each date
  // is one shift with its times left for the admin to fill in.
  let anchors = times.map((tm) => ({ pos: tm.start, time: tm }));
  if (!anchors.length) anchors = dates.map((d) => ({ pos: d.start, date: d }));
  if (!anchors.length && runs.length) anchors = [{ pos: runs[0].start }];
  if (!anchors.length) return [];

  const clauses = clauseBounds(src);
  const clauseOf = (pos) => clauses.find(([a, b]) => pos >= a && pos < b) || [0, src.length];
  // A message-wide value is used when a clause doesn't name its own and the
  // message only ever mentions one (e.g. "...Newmarket tomorrow ... V7 1322").
  const only = (list) => (list.length > 0 && list.every((x) => (x.date ?? x.run ?? x.regionKey) === (list[0].date ?? list[0].run ?? list[0].regionKey)) ? list[0] : null);

  const pick = (list, anchor, clause) => {
    const local = list.filter((x) => inside(x, clause));
    if (local.length) return nearest(local, anchor.pos);
    const one = only(list);
    if (one) return one;
    // Several candidates elsewhere in the message: take the closest one
    // before this shift (dispatch usually states the context first).
    const before = list.filter((x) => x.start < anchor.pos);
    return before.length ? before[before.length - 1] : null;
  };

  return anchors.map((anchor) => {
    const clause = clauseOf(anchor.pos);
    const date = anchor.date || pick(dates, anchor, clause);
    const run = pick(runs, anchor, clause);
    const loc = pick(locs, anchor, clause);
    const st = normalizeTime(anchor.time?.startRaw);
    const en = normalizeTime(anchor.time?.endRaw);
    const flags = [];
    if (!date) flags.push(FLAGS.missingDate);
    else flags.push(...date.flags);
    if (!anchor.time) flags.push(FLAGS.missingStart);
    else if (!st.valid) flags.push(FLAGS.invalidStart);
    if (anchor.time && en.empty) flags.push(FLAGS.missingEnd);
    else if (anchor.time && !en.valid) flags.push(FLAGS.invalidEnd);
    if (!anchor.time) flags.push(FLAGS.missingEnd);
    return {
      date: date ? date.date : "",
      parsedDate: date ? date.date : "",
      dateFlags: date ? date.flags.slice() : [],
      dateText: date ? date.text : "",
      weekday: date ? weekdayOf(date.date) : null,
      run: run ? run.run : "",
      location: loc ? loc.location : "",
      regionKey: loc ? loc.regionKey : "",
      start: st.text,
      end: en.text,
      flags: [...new Set(flags)],
    };
  });
}

// Re-checks one (possibly hand-edited) preview row. The date warnings from
// parsing (relative date, weekday mismatch) stay until the admin changes
// the date by hand.
export function validateShiftRow(row, today = todayStr()) {
  const flags = [];
  if (!row.date || !/^\d{4}-\d{2}-\d{2}$/.test(row.date)) flags.push(FLAGS.missingDate);
  else if (row.date < today) flags.push(FLAGS.pastDate);
  if (row.date && row.date === row.parsedDate) flags.push(...(row.dateFlags || []).filter((f) => f !== FLAGS.pastDate));
  const st = normalizeTime(row.start), en = normalizeTime(row.end);
  if (st.empty) flags.push(FLAGS.missingStart);
  else if (!st.valid) flags.push(FLAGS.invalidStart);
  if (en.empty) flags.push(FLAGS.missingEnd);
  else if (!en.valid) flags.push(FLAGS.invalidEnd);
  return [...new Set(flags)];
}

// Flags that block saving a row (the others are warnings to double-check).
export const BLOCKING_FLAGS = [FLAGS.missingDate, FLAGS.pastDate, FLAGS.missingStart, FLAGS.invalidStart, FLAGS.invalidEnd];

// ---- applying saved extra shifts to the weekly board ----
// The weekly board is a template week (dayIdx 0-6). An extra shift on a real
// date is shown on that weekday only for the Sun–Sat week containing it.
// Rows are copied, never mutated, so the original board stays intact.
export function weekBounds(dateStr) {
  const start = addDays(dateStr, -weekdayOf(dateStr));
  return { start, end: addDays(start, 6) };
}

export function extraCrewKey(driver) {
  if (driver.crewNumber) return String(driver.crewNumber);
  return `P${driver.employeeId || driver.id}`;
}

export function applyExtraShifts(crews, extraShifts, drivers, refDate) {
  const base = crews || [];
  if (!extraShifts || !extraShifts.length || !refDate) return base;
  const { start, end } = weekBounds(refDate);
  const inWeek = extraShifts.filter((s) => s.date >= start && s.date <= end);
  if (!inWeek.length) return base;
  const byId = new Map((drivers || []).map((d) => [d.id, d]));
  const out = base.slice();
  const indexOf = new Map(out.map((c, i) => [String(c.crew), i]));
  for (const s of inWeek) {
    const driver = byId.get(s.driverId);
    if (!driver) continue;
    const key = extraCrewKey(driver);
    let i = indexOf.get(key);
    if (i === undefined) {
      out.push({ crew: key, type: driver.type === "part" ? "PT" : "FT", shiftRaw: "", totalHours: 0, workedCount: 0, days: [], driverName: driver.name, employeeId: driver.employeeId || "", extraOnly: true });
      i = out.length - 1;
      indexOf.set(key, i);
    }
    const row = { ...out[i], days: out[i].days.slice() };
    const dayIdx = weekdayOf(s.date);
    const hours = shiftHours(s.start, s.end);
    const prev = row.days.find((d) => d.dayIdx === dayIdx);
    row.days = row.days.filter((d) => d.dayIdx !== dayIdx);
    row.days.push({ dayIdx, code: s.run || "", hours: hours ?? 0, start: s.start, end: s.end || "", regionKey: s.regionKey || "OTHER", extra: true, extraDate: s.date, replaced: prev || null });
    row.days.sort((a, b) => a.dayIdx - b.dayIdx);
    row.workedCount = row.days.length;
    row.totalHours = Math.round(row.days.reduce((sum, d) => sum + (Number(d.hours) || 0), 0) * 100) / 100;
    out[i] = row;
  }
  return out;
}

// Extra shifts whose date has passed are dropped.
export function pruneExpiredShifts(list, today = todayStr()) {
  return (list || []).filter((s) => s.date >= today);
}
