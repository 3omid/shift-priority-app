import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import pkg from "../package.json";
import {
  Upload, RotateCcw, ListChecks, AlertCircle, Check, Printer,
  FileSpreadsheet, GitCompare, X, Trophy, Medal, Award, ArrowUp, ArrowDown, ArrowLeft, Plus,
  Menu, Sun, Moon, HelpCircle, Trash2, Users, Info, Mail, LogOut,
  User, Star, CalendarOff, Shield, Lock, Search, ClipboardList, Pencil,
} from "lucide-react";

const APP_VERSION = pkg.version;
const DAY_NAMES = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
const CATCHALL = "OTHER";
const OFFICE_THEME = ["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
// Personal profile (name + "my crew number") — lives only in this
// browser/device's localStorage. Nothing here is ever sent anywhere, so one
// person's saved profile can never show up for anyone else who opens this
// same app/build.
const PROFILE_KEY = "shiftPriorityProfile";
// Crew-number -> driver-name directory. A manually-entered/edited name
// (added or fixed from the Admin panel) always wins over a name auto-read
// from an uploaded file's "Driver Name" column — see resolveCrewName().
// Also per-device localStorage only, same as PROFILE_KEY above.
const CREW_NAMES_KEY = "shiftPriorityCrewNames";
// Whether this browser tab is currently unlocked as Admin. Deliberately
// sessionStorage (not localStorage): it clears itself when the tab/app
// closes, instead of leaving Admin unlocked forever on a shared device.
const ADMIN_SESSION_KEY = "shiftPriorityAdminSession";
// NOTE ON SECURITY: this is a purely client-side app with no server, so
// there is no real way to keep a password secret here — anyone who opens
// this source file (or the browser's dev tools) can read these two
// constants in plain text. This gate only hides the Admin panel from
// casual/curious users; it is NOT protection against someone who actually
// wants in. Never reuse a real/sensitive password for this.
const ADMIN_USERNAME = "omid";
const ADMIN_PASSWORD = "ABC2020";

// Built-in default crew-number -> driver-name directory, seeded from the
// printed crew list posted at the workplace (since the uploaded Excel
// itself doesn't have a Driver Name column). This is the lowest-priority
// name source: an auto-read "Driver Name" column in an uploaded file wins
// over it, and a manual Admin entry wins over both — see resolveCrewName().
// Each crew number belongs to exactly one driver — never repeat a name
// under two different crew numbers here. Crew numbers deliberately left
// OUT of this list (2, 4, 11, 19, ...) are open/unassigned runs on the
// current board: no driver, and none should be guessed. Leave a newly-open
// crew out entirely rather than adding it with an empty string —
// resolveCrewName() already treats "not in this list" as "no default name".
// To update this list later (someone moves crews, a new hire, etc.), either
// edit it here and redeploy, or just fix it per-crew from the Admin panel —
// a manual edit there always overrides this list, and the Admin panel
// itself refuses to save the same name under two different crew numbers.
const CREW_NAME_DEFAULTS = {
  "1": "Sam Singh",
  "3": "Pirapa Maheswaran",
  "5": "Mojtaba Hosseini",
  "6": "Bosco James",
  "7": "Jun Wang",
  "8": "Behnaz Kheirandish",
  "9": "Jeremy Range",
  "10": "Elahe Alamdar",
  "12": "Hussein Jiwan",
  "13": "Greg Black",
  "14": "Durley Torres",
  "15": "Siavash Rafiee",
  "16": "On Phan",
  "17": "Ibrar Khan",
  "18": "Aneesa Salmon",
  "20": "Jasmine Majaski",
  "21": "Protacio Mapanao",
  "22": "Jibril Abdinasir",
  "23": "Sameh Fanous",
  "24": "Kevin Pang",
  "25": "Manjit Rai",
  "26": "Adel Nassim",
  "27": "Sonny Ogley",
  "28": "Garth Mantock",
  "29": "Oliver Pelboos",
  "30": "Danilo Fuentes",
  "31": "Harwinder Singh",
  "32": "Sahar DadgarAzad",
  "33": "Joe Jiang",
  "34": "Amani Asaad",
  "35": "Geeta Lal",
  "36": "Mozhgan Fatehi",
  "37": "Moyosore Alabi",
  "38": "Becka Dennie",
  "39": "Julie Edwards",
  "40": "Omid Farhadnia",
  "41": "Candy Lin",
  "42": "Morteza Hosseini",
};

// The last successfully-parsed schedule (already-parsed crew data, not the
// raw Excel file) so re-opening the app — closing and reopening the tab,
// relaunching the installed PWA, restarting the Electron app — shows the
// same schedule again instead of forcing a re-upload every single time.
// Also localStorage-only: same one-device-only rule as everything above.
const LAST_FILE_KEY = "shiftPriorityLastFile";

// ---------- Excel parsing ----------

function normalize(v) {
  return String(v ?? "").trim();
}

function findLayout(aoa) {
  let bestMissing = null;
  for (let r = 0; r < Math.min(aoa.length, 8); r++) {
    const row = (aoa[r] || []).map((c) => String(c ?? "").trim().toLowerCase());
    const crewIdx = row.findIndex((c) => c === "crew #" || c === "crew#" || c.startsWith("crew"));
    if (crewIdx === -1) continue;
    const typeIdx = row.findIndex((c) => c === "type");
    const totalIdx = row.findIndex((c) => c.includes("total"));
    const dayIdx = DAY_NAMES.map((dn) => row.findIndex((c) => c === dn));
    // Optional — not required for the sheet to be usable, just read in when
    // present. "Driver Name" is the header TOK Transit sheets actually use;
    // a couple of common variants are matched too.
    const nameIdx = row.findIndex((c) => c === "driver name" || c === "name" || c === "employee name" || c === "employee");
    const missing = [];
    if (typeIdx === -1) missing.push("Type");
    if (totalIdx === -1) missing.push("TOTAL HRS");
    DAY_NAMES.forEach((dn, i) => { if (dayIdx[i] === -1) missing.push(dn.toUpperCase()); });
    if (missing.length === 0) {
      return { ok: true, layout: { headerRow: r, crewIdx, typeIdx, shiftIdx: typeIdx + 1, dayIdx, totalIdx, nameIdx } };
    }
    if (!bestMissing || missing.length < bestMissing.length) bestMissing = missing;
  }
  return { ok: false, missing: bestMissing || ["Crew #"] };
}

function extractTextTag(code) {
  const s = normalize(code);
  if (!s) return CATCHALL;
  const m = s.match(/([A-Za-z]{2,})\s*$/);
  if (m) return m[1].toUpperCase();
  return CATCHALL;
}

function applyTint(hex, tint) {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
  const adj = (c) => {
    let v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(v)));
  };
  const toHex = (c) => c.toString(16).padStart(2, "0").toUpperCase();
  return toHex(adj(r)) + toHex(adj(g)) + toHex(adj(b));
}

function cellColorInfo(cell) {
  if (!cell || !cell.s) return null;
  const fg = cell.s.fgColor || cell.s.bgColor;
  if (!fg) return null;
  if (fg.rgb) {
    const hex = fg.rgb.length === 8 ? fg.rgb.slice(2) : fg.rgb;
    return { key: `rgb-${hex}`, hex };
  }
  if (fg.theme !== undefined) {
    const base = OFFICE_THEME[fg.theme] || "808080";
    const tint = fg.tint || 0;
    const hex = applyTint(base, tint);
    return { key: `theme-${fg.theme}-${tint.toFixed(2)}`, hex };
  }
  return null;
}

const KNOWN_COLORS = { F8CBAD: "RH", E2EFDA: "NMK", "70AD47": "RPT", D9E1F2: "CLDR" };

function hexToRgb(hex) {
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function nearestKnownColor(hex) {
  if (!hex) return null;
  const [r, g, b] = hexToRgb(hex);
  let bestKey = null, bestDist = Infinity;
  for (const [khex, key] of Object.entries(KNOWN_COLORS)) {
    const [kr, kg, kb] = hexToRgb(khex);
    const dist = Math.sqrt((r - kr) ** 2 + (g - kg) ** 2 + (b - kb) ** 2);
    if (dist < bestDist) { bestDist = dist; bestKey = key; }
  }
  return bestDist <= 40 ? bestKey : null;
}

function regionKeyFor(textTag, colorInfo) {
  if (textTag === "STF") return "STF";
  if (textTag === "MRG") return "MRG";
  if (textTag === "RPT") return "RPT";
  if (colorInfo) return nearestKnownColor(colorInfo.hex) || CATCHALL;
  return CATCHALL;
}

// Some workbooks keep old/helper sheets hidden; only offer visible tabs,
// matching what the user actually sees in Excel.
function visibleSheetNames(wb) {
  const meta = wb.Workbook && wb.Workbook.Sheets;
  if (!meta) return wb.SheetNames;
  const visible = wb.SheetNames.filter((_, idx) => !meta[idx] || !meta[idx].Hidden);
  return visible.length ? visible : wb.SheetNames;
}

function parseSchedule(sheet) {
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const layoutResult = findLayout(aoa);
  if (!layoutResult.ok) return { ok: false, reason: "no_layout", missing: layoutResult.missing };
  const { headerRow, crewIdx, typeIdx, shiftIdx, dayIdx, totalIdx, nameIdx } = layoutResult.layout;

  const crews = [];
  let coloredCells = 0;
  let totalWorkedCells = 0;

  for (let r = headerRow + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const crewVal = row[crewIdx];
    if (typeof crewVal !== "number" && !(typeof crewVal === "string" && /^\d+$/.test(crewVal.trim()))) break;

    const type = normalize(row[typeIdx]);
    const shiftRaw = normalize(row[shiftIdx]).toUpperCase();
    const totalHours = typeof row[totalIdx] === "number" ? row[totalIdx] : parseFloat(row[totalIdx]) || 0;
    const driverName = nameIdx !== -1 ? normalize(row[nameIdx]) : "";

    const days = [];
    let workedCount = 0;
    for (let d = 0; d < 7; d++) {
      const col = dayIdx[d];
      const code = row[col];
      const hoursCell = row[col + 3];
      const startCell = row[col + 1];
      const endCell = row[col + 2];
      const hasHours = typeof hoursCell === "number" && hoursCell > 0;
      if (!hasHours) continue;
      totalWorkedCells++;
      workedCount++;

      const cellRef = XLSX.utils.encode_cell({ r, c: col });
      const cellObj = sheet[cellRef];
      const colorInfo = cellColorInfo(cellObj);
      const textTag = extractTextTag(code);
      const regionKey = regionKeyFor(textTag, colorInfo);
      if (colorInfo) coloredCells++;

      days.push({ dayIdx: d, code: normalize(code), hours: hoursCell, start: startCell, end: endCell, regionKey });
    }
    if (workedCount === 0 && !type) continue;
    crews.push({ crew: crewVal, type, shiftRaw, totalHours, workedCount, days, driverName });
  }

  if (crews.length === 0) return { ok: false, reason: "no_data" };
  const colorsDetected = totalWorkedCells > 0 && coloredCells / totalWorkedCells > 0.4;
  return { ok: true, crews, colorsDetected };
}

// ---------- Criteria catalog ----------

function singleShiftScore(shiftRaw, pref) {
  const hasAM = shiftRaw.includes("AM");
  const hasPM = shiftRaw.includes("PM");
  if (pref === "morning") return hasAM && !hasPM ? 100 : hasAM && hasPM ? 55 : 0;
  if (pref === "afternoon") return hasPM && !hasAM ? 100 : hasAM && hasPM ? 55 : 0;
  if (pref === "mixed") return hasAM && hasPM ? 100 : hasAM || hasPM ? 40 : 0;
  return 50;
}

function regionPct(c, key) {
  if (c.days.length === 0) return 0;
  return Math.round((100 * c.days.filter((d) => d.regionKey === key).length) / c.days.length);
}

const REGION_COLORS = { NMK: "#70AD47", RH: "#F4A261", CLDR: "#5B8DEF", STF: "#C9A227", MRG: "#8E7CC3", RPT: "#B3432A", [CATCHALL]: "#9AA0A6" };

const CRITERIA_CATALOG = [
  { id: "shift_am", group: "shift", label: { fa: "شیفت AM", en: "AM shift", hi: "AM शिफ्ट" }, score: (c) => singleShiftScore(c.shiftRaw, "morning") },
  { id: "shift_pm", group: "shift", label: { fa: "شیفت PM", en: "PM shift", hi: "PM शिफ्ट" }, score: (c) => singleShiftScore(c.shiftRaw, "afternoon") },
  { id: "shift_mixed", group: "shift", label: { fa: "شیفت ترکیبی AM/PM", en: "AM/PM mixed shift", hi: "AM/PM मिश्रित शिफ्ट" }, score: (c) => singleShiftScore(c.shiftRaw, "mixed") },
  { id: "days_4", group: "days", label: { fa: "۴ روز کاری", en: "4 working days", hi: "४ कार्य दिवस" }, score: (c) => (c.workedCount === 4 ? 100 : 0) },
  { id: "days_5", group: "days", label: { fa: "۵ روز کاری", en: "5 working days", hi: "५ कार्य दिवस" }, score: (c) => (c.workedCount === 5 ? 100 : 0) },
  { id: "region_NMK", group: "region", label: { fa: "نیومارکت", en: "Newmarket", hi: "न्यूमार्केट" }, score: (c) => regionPct(c, "NMK") },
  { id: "region_RH", group: "region", label: { fa: "ریچموند هیل", en: "Richmond Hill", hi: "रिचमंड हिल" }, score: (c) => regionPct(c, "RH") },
  { id: "region_CLDR", group: "region", label: { fa: "کالداری", en: "Caldari", hi: "कैल्डारी" }, score: (c) => regionPct(c, "CLDR") },
  { id: "region_STF", group: "region", label: { fa: "استوفیل", en: "Stouffville", hi: "स्टफविल" }, score: (c) => regionPct(c, "STF") },
  { id: "region_MRG", group: "region", label: { fa: "مپل", en: "Maple", hi: "मेपल" }, score: (c) => regionPct(c, "MRG") },
  { id: "has_rpt", group: "special", label: { fa: "دارای شیفت آماده‌باش (RPT)", en: "Has standby (RPT) day", hi: "स्टैंडबाय (RPT) दिन है" }, score: (c) => regionPct(c, "RPT") },
  // Dynamic criterion: unlike the fixed ones above, this one doesn't score
  // anything on its own — the person picks WHICH weekdays it means (via the
  // day-picker shown once it's added to their priority list, up to 3 days),
  // and that choice is resolved at compute-time in computeResults(). It's a
  // single catalog entry (one slot in the priority chain) whose day list can
  // hold up to 3 weekdays, rather than three separate slots.
  { id: "day_off", group: "days", dynamic: true, multi: true, label: { fa: "روزهای تعطیل دلخواه", en: "Custom days off", hi: "पसंदीदा छुट्टियाँ" } },
];
const CRITERIA_MAX = 10;
const DAY_OFF_MAX_DAYS = 3;

function criterionColor(id) {
  if (id.startsWith("region_")) return REGION_COLORS[id.replace("region_", "")];
  if (id === "has_rpt") return REGION_COLORS.RPT;
  if (id === "day_off") return "#B3432A";
  return "var(--accent)";
}

const DAY_LIST_SEP = { fa: "، ", en: ", ", hi: ", " };

// The day-off criterion's label grows the chosen weekday names once any are
// picked (e.g. "Custom days off" -> "Custom days off — Friday, Saturday").
// Every other criterion's label passes through unchanged. Centralizing this
// here means every place that ever displayed `crit.label[lang]` (chips, the
// priority list, history, Excel/print export, the compare table) shows the
// same resolved text instead of six separate copies of this logic.
function resolveCriterionLabel(id, dayOffChoices) {
  const base = CRITERIA_CATALOG.find((c) => c.id === id);
  if (!base) return { fa: id, en: id, hi: id };
  if (!base.dynamic) return base.label;
  const dayIdxs = dayOffChoices ? dayOffChoices[id] : null;
  if (!Array.isArray(dayIdxs) || dayIdxs.length === 0) return base.label;
  return {
    fa: `${base.label.fa} — ${dayIdxs.map((i) => WEEKDAY_LABELS.fa[i]).join(DAY_LIST_SEP.fa)}`,
    en: `${base.label.en} — ${dayIdxs.map((i) => WEEKDAY_LABELS.en[i]).join(DAY_LIST_SEP.en)}`,
    hi: `${base.label.hi} — ${dayIdxs.map((i) => WEEKDAY_LABELS.hi[i]).join(DAY_LIST_SEP.hi)}`,
  };
}

// Score = what fraction of the chosen days this crew actually has off, as a
// 0-100 percentage (e.g. 2 of 3 chosen days off -> 67). No days chosen -> 0.
function dayOffScore(c, dayIdxs) {
  if (!Array.isArray(dayIdxs) || dayIdxs.length === 0) return 0;
  const offCount = dayIdxs.filter((dayIdx) => !c.days.some((d) => d.dayIdx === dayIdx)).length;
  return Math.round((offCount / dayIdxs.length) * 100);
}

function computeResults(crews, priorityList, dayOffChoices) {
  const criteriaObjs = priorityList.map((id) => {
    const base = CRITERIA_CATALOG.find((c) => c.id === id);
    if (!base) return null;
    const label = resolveCriterionLabel(id, dayOffChoices);
    if (base.dynamic) {
      const dayIdxs = dayOffChoices ? dayOffChoices[id] : null;
      return { id: base.id, label, score: (c) => dayOffScore(c, dayIdxs) };
    }
    return { id: base.id, label, score: base.score };
  }).filter(Boolean);
  const withFp = crews.map((c) => {
    const fingerprint = criteriaObjs.map((crit) => ({ id: crit.id, label: crit.label, score: crit.score(c) }));
    const regionSummary = {};
    c.days.forEach((d) => { (regionSummary[d.regionKey] = regionSummary[d.regionKey] || []).push(d.hours); });
    return { ...c, fingerprint, regionSummary };
  });
  withFp.sort((a, b) => {
    for (let i = 0; i < criteriaObjs.length; i++) {
      const diff = b.fingerprint[i].score - a.fingerprint[i].score;
      if (diff !== 0) return diff;
    }
    return 0;
  });
  return withFp;
}

// Displayed match % must never contradict the sort order: a lower-ranked
// crew must never show a higher percentage than one ranked above it. A
// simple weighted average can't guarantee that, so instead we encode the
// lexicographic comparison itself into one number: each criterion's tier
// completely dominates all lower tiers combined (using BigInt so it stays
// exact even with up to 10 criteria).
function displayScore(fingerprint) {
  const n = fingerprint.length;
  if (n === 0) return "0.00";
  // BASE just needs to be > 100 to mathematically guarantee a higher-priority
  // criterion always outweighs every lower criterion combined (so order can
  // never invert). Keeping it close to that minimum (101) — instead of a
  // huge value like 1000 — keeps each lower criterion's effect big enough
  // to actually show up in the displayed percentage instead of rounding away.
  const BASE = 101n;
  let weightedSum = 0n;
  let maxSum = 0n;
  for (let i = 0; i < n; i++) {
    const w = BASE ** BigInt(n - 1 - i);
    weightedSum += BigInt(Math.round(fingerprint[i].score)) * w;
    maxSum += 100n * w;
  }
  // With deep priority lists (6-7+ criteria), many crews legitimately tie on
  // the first 1-2 priorities and only differ starting at #3+ — their true
  // difference is real but tiny relative to how strongly the top priorities
  // dominate. Keep 4 decimal places so those legitimate (small) differences
  // are still visible instead of collapsing into "identical-looking" numbers.
  const scaled = (weightedSum * 100000000n) / maxSum;
  const pct = Math.round(Number(scaled) / 10000) / 10000;
  return pct;
}

// UI-only: show the priority score (0..1) as a whole-number percent of 100.
// Floor, not round, so only a perfect match can ever read "100%".
function scorePercent(ds) {
  return Math.floor(Number(ds) * 100 + 1e-9);
}

// UI-only colour for the overall priority %: >=75 gold, 40-74 green, <40 grey.
// The actual shades come from CSS vars set per light/dark mode (see rootVars).
function scoreColor(pct) {
  return pct >= 75 ? "var(--score-high)" : pct >= 40 ? "var(--score-mid)" : "var(--score-low)";
}

// ---------- i18n ----------

const REGION_LABELS = {
  NMK: { fa: "نیومارکت", en: "Newmarket", hi: "न्यूमार्केट" }, RH: { fa: "ریچموند هیل", en: "Richmond Hill", hi: "रिचमंड हिल" },
  CLDR: { fa: "کالداری", en: "Caldari", hi: "कैल्डारी" }, STF: { fa: "استوفیل", en: "Stouffville", hi: "स्टफविल" },
  MRG: { fa: "مپل", en: "Maple", hi: "मेपल" }, RPT: { fa: "آماده‌باش", en: "Standby", hi: "स्टैंडबाय" }, [CATCHALL]: { fa: "سایر", en: "Other", hi: "अन्य" },
};

const STRINGS = {
  title: { fa: "اولویت‌بندی شیفت", en: "Shift Prioritizer", hi: "शिफ्ट प्राथमिकता" },
  subtitle: { fa: "جدول شیفت رو آپلود کن، اولویت‌هاتو بچین", en: "Upload your shift sheet and build your priorities", hi: "अपनी शिफ्ट शीट अपलोड करें और प्राथमिकताएँ तय करें" },
  uploadStep: { fa: "۱. فایل اکسل شیفت‌ها", en: "1. Shift Excel File", hi: "१. शिफ्ट एक्सेल फ़ाइल" },
  uploadPlaceholder: { fa: "برای انتخاب فایل ضربه بزن (xlsx / xls)", en: "Tap to choose a file (xlsx / xls)", hi: "फ़ाइल चुनने के लिए टैप करें (xlsx / xls)" },
  changeFileLink: { fa: "تغییر", en: "Change", hi: "बदलें" },
  sheetLabel: { fa: "شیت:", en: "Sheet:", hi: "शीट:" },
  columnsDetected: { fa: "ستون‌های Crew #، Type، شیفت AM/PM، روزهای هفته و مجموع ساعات تشخیص داده شدند", en: "Crew #, Type, AM/PM shift, weekday and total-hours columns detected", hi: "Crew #, Type, AM/PM शिफ्ट, सप्ताह के दिन और कुल घंटे कॉलम पहचाने गए" },
  colorsDetectedYes: { fa: "منطقه از رنگ گراج + کدهای STF/MRG/RPT تشخیص داده شد", en: "Region detected from garage color + STF/MRG/RPT codes", hi: "गैराज रंग + STF/MRG/RPT कोड से क्षेत्र पहचाना गया" },
  colorsDetectedNo: { fa: "رنگ سلول‌ها خوانده نشد؛ فقط از کدهای STF/MRG/RPT استفاده شد", en: "Cell colors unreadable; only STF/MRG/RPT codes were used", hi: "सेल रंग नहीं पढ़े जा सके; केवल STF/MRG/RPT कोड उपयोग किए गए" },
  prefsStep: { fa: "۲. اولویت‌هاتو بچین (مهم‌ترین اول)", en: "2. Build your priorities (most important first)", hi: "२. प्राथमिकताएँ बनाएँ (सबसे ज़रूरी पहले)" },
  builderHint: { fa: "روی هر گزینه بزن تا به انتهای لیست اولویت اضافه بشه. با فلش‌ها ترتیب رو تغییر بده — رتبه‌بندی دقیقاً به همین ترتیب انجام می‌شه. حداکثر ۱۰ مورد.", en: "Tap an option to add it to your priority list. Reorder with the arrows — ranking follows this exact order. Up to 10 items.", hi: "किसी विकल्प को सूची में जोड़ने के लिए टैप करें। तीरों से क्रम बदलें। अधिकतम १० आइटम।" },
  catalogShift: { fa: "شیفت", en: "Shift", hi: "शिफ्ट" },
  catalogDays: { fa: "روزهای کاری", en: "Working days", hi: "कार्य दिवस" },
  catalogRegion: { fa: "منطقه", en: "Region", hi: "क्षेत्र" },
  catalogSpecial: { fa: "ویژه", en: "Special", hi: "विशेष" },
  yourPriorities: { fa: "لیست اولویت تو", en: "Your priority list", hi: "आपकी प्राथमिकता सूची" },
  resetPriorities: { fa: "پاک کردن اولویت‌ها", en: "Clear priorities", hi: "प्राथमिकताएँ साफ़ करें" },
  emptyPriorities: { fa: "هنوز چیزی انتخاب نکردی — از بالا یه گزینه بزن.", en: "Nothing selected yet — tap an option above.", hi: "अभी कुछ नहीं चुना गया।" },
  computeBtn: { fa: "محاسبه و رتبه‌بندی", en: "Calculate & Rank", hi: "गणना करें और रैंक करें" },
  resultsTitle: { fa: "نتیجه — به ترتیب اولویت", en: "Results — by priority", hi: "परिणाम — प्राथमिकता क्रम में" },
  best: { fa: "بهترین گزینه‌ها", en: "Top matches", hi: "सर्वश्रेष्ठ विकल्प" },
  rest: { fa: "بقیه نتایج", en: "Other results", hi: "अन्य परिणाम" },
  crewWord: { fa: "گروه", en: "Crew", hi: "क्रू" },
  type: { fa: "نوع", en: "Type", hi: "प्रकार" },
  shift: { fa: "شیفت", en: "Shift", hi: "शिफ्ट" },
  days: { fa: "روز", en: "days", hi: "दिन" },
  hours: { fa: "ساعت", en: "hrs", hi: "घंटे" },
  addCompare: { fa: "مقایسه", en: "Compare", hi: "तुलना" },
  compareTitle: { fa: "مقایسه گروه‌های انتخابی", en: "Compare selected crews", hi: "चयनित क्रू की तुलना" },
  printBtn: { fa: "پرینت / ذخیره PDF", en: "Print / Save as PDF", hi: "प्रिंट / PDF सेव करें" },
  printPopupBlocked: {
    fa: "مرورگر نتونست پنجره گزارش رو باز کنه (پاپ‌آپ مسدود شده). از تنظیمات مرورگر، پاپ‌آپ رو برای این سایت مجاز کن و دوباره امتحان کن.",
    en: "Your browser blocked the report window (pop-up blocked). Allow pop-ups for this site in your browser settings and try again.",
    hi: "ब्राउज़र रिपोर्ट विंडो नहीं खोल सका (पॉप-अप ब्लॉक)। इस साइट के लिए पॉप-अप की अनुमति दें और फिर से कोशिश करें.",
  },
  excelBtn: { fa: "خروجی اکسل", en: "Export Excel", hi: "एक्सेल में निर्यात करें" },
  resetBtn: { fa: "شروع دوباره با فایل جدید", en: "Start over with a new file", hi: "नई फ़ाइल से फिर शुरू करें" },
  errNoLayout: { fa: "ساختار جدول شیفت در این شیت پیدا نشد.", en: "Could not detect the schedule structure in this sheet.", hi: "इस शीट में शेड्यूल संरचना नहीं मिली।" },
  errMissingColumns: { fa: "ستون‌های زیر پیدا نشد:", en: "Missing column(s):", hi: "ये कॉलम नहीं मिले:" },
  errNoData: { fa: "ستون‌ها پیدا شدند ولی هیچ ردیف داده‌ای زیرشون نبود.", en: "Columns were found but no data rows were detected underneath them.", hi: "कॉलम मिले लेकिन उनके नीचे कोई डेटा पंक्ति नहीं मिली।" },
  fileError: { fa: "خواندن فایل با خطا مواجه شد.", en: "Failed to read the file.", hi: "फ़ाइल पढ़ने में त्रुटि हुई।" },
  region: { fa: "مناطق", en: "Regions", hi: "क्षेत्र" },
  reportDate: { fa: "تاریخ گزارش", en: "Report date", hi: "रिपोर्ट तिथि" },
  rank: { fa: "رتبه", en: "Rank", hi: "रैंक" },
  crewsFoundWord: { fa: "گروه پیدا شد", en: "crews found", hi: "क्रू मिले" },
  daysColumn: { fa: "روزها", en: "Days", hi: "दिन" },
  hoursWord: { fa: "ساعت", en: "hrs", hi: "घंटे" },
  menu: { fa: "منو", en: "Menu", hi: "मेनू" },
  settingsTitle: { fa: "ظاهر برنامه", en: "Appearance", hi: "थीम" },
  themeMode: { fa: "حالت", en: "Mode", hi: "मोड" },
  light: { fa: "روز", en: "Light", hi: "लाइट" },
  dark: { fa: "شب", en: "Dark", hi: "डार्क" },
  themeStyle: { fa: "سبک تم", en: "Theme style", hi: "थीम शैली" },
  themeWin11: { fa: "ویندوز ۱۱", en: "Windows 11", hi: "विंडोज़ ११" },
  themeMac: { fa: "مک‌اواس", en: "macOS", hi: "मैकओएस" },
  themeUniversal: { fa: "عمومی", en: "Universal", hi: "यूनिवर्सल" },
  helpTitle: { fa: "راهنمای استفاده", en: "How to use", hi: "उपयोग मार्गदर्शिका" },
  aboutTitle: { fa: "درباره برنامه", en: "About", hi: "ऐप के बारे में" },
  reportProblem: { fa: "گزارش مشکل / پیشنهاد", en: "Report a problem / feedback", hi: "समस्या रिपोर्ट करें" },
  exitApp: { fa: "خروج از برنامه", en: "Exit App", hi: "ऐप से बाहर निकलें" },
  exitWebNote: { fa: "چون این نسخه‌ی وب/PWA است، مرورگرها اجازه نمی‌دهند صفحه خودش را کاملاً ببندد. برای خروج، برگه یا برنامه را طبق روال معمول دستگاهت ببند.", en: "This is the web/PWA version, so browsers don't allow a page to close itself. To exit, close the tab or app the normal way for your device.", hi: "यह वेब/PWA संस्करण है, इसलिए ब्राउज़र पेज को खुद बंद करने की अनुमति नहीं देते। बाहर निकलने के लिए टैब या ऐप को सामान्य तरीके से बंद करें।" },
  compare2Title: { fa: "مقایسه گروه‌ها در کل هفته", en: "Compare crews for the whole week", hi: "क्रू की पूरी हफ़्ते तुलना" },
  crewA: { fa: "شماره گروه اول", en: "First crew number", hi: "पहला क्रू नंबर" },
  crewB: { fa: "شماره گروه دوم", en: "Second crew number", hi: "दूसरा क्रू नंबर" },
  compareBtn: { fa: "مقایسه کن", en: "Compare", hi: "तुलना करें" },
  crewNotFound: { fa: "این شماره گروه در فایل پیدا نشد.", en: "That crew number was not found in the file.", hi: "यह क्रू नंबर फ़ाइल में नहीं मिला।" },
  off: { fa: "تعطیل", en: "OFF", hi: "बंद" },
  close: { fa: "بستن", en: "Close", hi: "बंद करें" },

  // ---- day-off criteria ----
  pickDayHint: { fa: "کدوم روزها؟ (حداکثر ۳ تا)", en: "Which days? (up to 3)", hi: "कौन से दिन? (अधिकतम ३)" },
  errDayOffUnset: { fa: "برای «روزهای تعطیل دلخواه» که به لیست اضافه کردی، حداقل یه روز مشخص کن.", en: "Pick at least one weekday for the custom days-off criterion you added.", hi: "आपने जो \"पसंदीदा छुट्टियाँ\" जोड़ी है उसके लिए कम से कम एक दिन चुनें।" },

  // ---- loading splash ----
  loadingBoot: { fa: "در حال آماده‌سازی…", en: "Getting ready…", hi: "तैयार हो रहा है…" },
  loadingCompute: { fa: "در حال محاسبه و رتبه‌بندی…", en: "Calculating & ranking…", hi: "गणना और रैंकिंग हो रही है…" },

  // ---- profile ----
  profileTitle: { fa: "پروفایل من", en: "My Profile", hi: "मेरी प्रोफ़ाइल" },
  firstNameLabel: { fa: "نام کوچیک", en: "First name", hi: "पहला नाम" },
  myCrewLabel: { fa: "شمارهٔ گروه/شیفت من", en: "My crew number", hi: "मेरा क्रू नंबर" },
  saveProfileBtn: { fa: "ذخیره", en: "Save", hi: "सहेजें" },
  clearProfileBtn: { fa: "پاک کردن پروفایل", en: "Clear profile", hi: "प्रोफ़ाइल साफ़ करें" },
  profilePrivacyHint: {
    fa: "این اطلاعات فقط روی همین دستگاه/مرورگر ذخیره می‌شه — هیچ‌جا فرستاده نمی‌شه، و برای بقیه کسایی که همین برنامه رو دارن دیده نمی‌شه.",
    en: "This is saved only on this device/browser — it's never sent anywhere, and no one else using this app will see it.",
    hi: "यह केवल इस डिवाइस/ब्राउज़र पर सहेजा जाता है — कहीं नहीं भेजा जाता, और इस ऐप के अन्य उपयोगकर्ता इसे नहीं देख सकते।",
  },
  welcomeBack: { fa: "خوش‌آمدید،", en: "Welcome back,", hi: "वापसी पर स्वागत है," },
  myShiftCard: { fa: "شیفت من", en: "My shift", hi: "मेरी शिफ्ट" },
  viewMySchedule: { fa: "مشاهدهٔ برنامهٔ من", en: "View my schedule", hi: "मेरा शेड्यूल देखें" },
  crewNumberNotInFile: { fa: "این شمارهٔ گروه توی این فایل پیدا نشد.", en: "That crew number wasn't found in this file.", hi: "यह क्रू नंबर इस फ़ाइल में नहीं मिला।" },
  myShiftBadge: { fa: "شیفت من", en: "Mine", hi: "मेरा" },
  myScheduleTitle: { fa: "برنامهٔ هفتگی من", en: "My weekly schedule", hi: "मेरा साप्ताहिक कार्यक्रम" },
  todayLabel: { fa: "امروز", en: "Today", hi: "आज" },

  // ---- help submenu ----
  helpMenuLabel: { fa: "راهنما", en: "Help", hi: "सहायता" },

  // ---- crew lookup ----
  crewLookupTitle: { fa: "مشاهدهٔ گروه‌ها", en: "Browse crews", hi: "क्रू ब्राउज़ करें" },
  hubGroupCrews: { fa: "مقایسه و جست‌وجوی گروه‌ها", en: "Compare & look up crews", hi: "क्रू तुलना और खोज" },
  hubGroupMe: { fa: "من", en: "Me", hi: "मेरा" },
  hubComingSoon: { fa: "به‌زودی", en: "Coming soon", hi: "जल्द आ रहा है" },
  hubSwapFinder: { fa: "جایگزین برای مرخصی", en: "Find a leave replacement", hi: "छुट्टी बदली खोजें" },
  swapIntro: { fa: "روز و ساعت شیفتی که می‌خوای مرخصی بگیری رو بده. اول کسایی میان که اون روز شیفت ندارن و روزهای استراحت تو سر کارن — یعنی می‌تونی باهاشون جابه‌جا کنی.", en: "Enter the day and hours you want off. Crews who are off that day come up first if they work one of YOUR rest days — so you can offer a straight trade.", hi: "जिस दिन और समय की छुट्टी चाहिए, दर्ज करें। जो क्रू उस दिन फ्री हैं और आपकी छुट्टी वाले दिन काम करते हैं, वे पहले आएंगे — ताकि आप अदला-बदली कर सकें।" },
  swapMyCrew: { fa: "شمارهٔ گروه من", en: "My crew #", hi: "मेरा क्रू #" },
  swapDate: { fa: "تاریخ", en: "Date", hi: "तारीख" },
  swapStart: { fa: "شروع شیفت", en: "Shift start", hi: "शिफ्ट शुरू" },
  swapEnd: { fa: "پایان شیفت", en: "Shift end", hi: "शिफ्ट समाप्त" },
  swapFindBtn: { fa: "پیدا کن", en: "Find replacements", hi: "बदली खोजें" },
  swapNeedInputs: { fa: "تاریخ و ساعت شروع و پایان رو وارد کن.", en: "Enter the date, start and end time.", hi: "तारीख, शुरू और समाप्ति समय दर्ज करें।" },
  swapNoCrewHint: { fa: "بدون شمارهٔ گروهت فقط کسایی که اون روز آزادن نشون داده می‌شن (پیشنهاد جابه‌جایی نه).", en: "Without your crew #, only crews free that day are shown — no trade suggestions.", hi: "क्रू # के बिना केवल उस दिन फ्री क्रू दिखेंगे — अदला-बदली सुझाव नहीं।" },
  swapAlreadyOff: { fa: "طبق برنامه، تو این روز خودت شیفت نداری.", en: "Per the schedule, you're already off on this day.", hi: "शेड्यूल के अनुसार, इस दिन आप पहले से छुट्टी पर हैं।" },
  swapMyRestDays: { fa: "روزهای استراحت من", en: "My rest days", hi: "मेरे आराम के दिन" },
  swapResultsTitle: { fa: "بهترین گزینه‌ها", en: "Best options", hi: "सबसे अच्छे विकल्प" },
  swapCandidatesWord: { fa: "نفر", en: "candidates", hi: "उम्मीदवार" },
  swapNoResults: { fa: "هیچ گروهی اون روز آزاد نیست.", en: "No crew is off on that day.", hi: "उस दिन कोई क्रू फ्री नहीं है।" },
  swapFit: { fa: "تناسب", en: "Fit", hi: "मेल" },
  swapOffThatDay: { fa: "آزاد در روز", en: "Off on", hi: "छुट्टी:" },
  swapUsualStart: { fa: "شروع معمول:", en: "Usually starts", hi: "आमतौर पर शुरू:" },
  swapTradeFor: { fa: "جابه‌جایی: تو شیفتش رو بری در", en: "Trade: you take their shift on", hi: "अदला-बदली: आप उनकी शिफ्ट लें" },
  swapNoTrade: { fa: "روزهای استراحت تو سر کار نیست — فقط کاور بدون جابه‌جایی", en: "Doesn't work on your rest days — cover only, no trade", hi: "आपके आराम के दिनों में काम नहीं — केवल कवर" },
  swapMyRestWarn: { fa: "برای تو استراحت کمتر از ۸ ساعت بین شیفت‌ها می‌شه", en: "Would leave YOU under 8h rest between shifts", hi: "आपको शिफ्टों के बीच 8 घंटे से कम आराम मिलेगा" },
  swapTheirRestWarn: { fa: "استراحت اون کمتر از ۸ ساعت می‌شه", en: "Would leave them under 8h rest", hi: "उन्हें 8 घंटे से कम आराम मिलेगा" },
  swapBefore: { fa: "قبل:", en: "before:", hi: "पहले:" },
  swapAfter: { fa: "بعد:", en: "after:", hi: "बाद:" },
  swapViewSchedule: { fa: "برنامهٔ هفتگی", en: "Weekly schedule", hi: "साप्ताहिक शेड्यूल" },
  swapMatch: { fa: "تطابق با برنامهٔ تو", en: "match with your schedule", hi: "आपके शेड्यूल से मेल" },
  myShiftHomeTitle: { fa: "شیفت من", en: "My Shift", hi: "मेरी शिफ्ट" },
  myShiftHomeNoCrew: { fa: "شمارهٔ گروهت رو توی پروفایل وارد کن", en: "Add your crew # in Profile", hi: "प्रोफ़ाइल में क्रू # जोड़ें" },
  myShiftHomeNoFile: { fa: "اول فایل برنامه رو بارگذاری کن", en: "Load the schedule file first", hi: "पहले शेड्यूल फ़ाइल लोड करें" },
  swapShowMore: { fa: "نمایش بیشتر", en: "Show more", hi: "और दिखाएं" },
  crewLookupSearch: { fa: "جستجوی شماره یا اسم...", en: "Search number or name…", hi: "नंबर या नाम खोजें…" },

  // ---- admin ----
  adminMenuLabel: { fa: "حالت ادمین", en: "Admin", hi: "एडमिन" },
  adminLoginTitle: { fa: "ورود ادمین", en: "Admin login", hi: "एडमिन लॉगिन" },
  adminUsernameLabel: { fa: "نام کاربری", en: "Username", hi: "उपयोगकर्ता नाम" },
  adminPasswordLabel: { fa: "رمز عبور", en: "Password", hi: "पासवर्ड" },
  adminLoginBtn: { fa: "ورود", en: "Log in", hi: "लॉग इन" },
  adminLoginError: { fa: "نام کاربری یا رمز عبور اشتباهه.", en: "Incorrect username or password.", hi: "गलत उपयोगकर्ता नाम या पासवर्ड।" },
  adminPanelTitle: { fa: "پنل ادمین — فهرست اسامی گروه‌ها", en: "Admin panel — crew name directory", hi: "एडमिन पैनल — क्रू नाम सूची" },
  adminPanelHint: {
    fa: "اسم جلوی هر شماره گروه رو می‌تونی ویرایش کنی. اولویت: ویرایش دستی (دکمهٔ حذف داره) > ستون «Driver Name» فایل اکسل > فهرست پیش‌فرض توی خود برنامه. هر گروهی که راننده نداره می‌تونه خالی بمونه — با دکمهٔ پاک‌کردن (✕) کنار هر ردیف می‌تونی هر گروهی رو صریحاً خالی کنی. برنامه جلوی ثبت یه اسم رو برای دو شماره گروه مختلف می‌گیره. ویرایش‌های دستی فقط روی همین دستگاه/مرورگر ذخیره می‌شن.",
    en: "Edit the name next to each crew number. Priority: a manual edit (has a delete button) beats the file's \"Driver Name\" column, which beats the app's built-in default list. A crew with no driver can stay blank — use the ✕ button on any row to explicitly blank it. The app won't let the same name be saved for two different crew numbers. Manual edits are saved only on this device/browser.",
    hi: "हर क्रू नंबर के सामने नाम बदल सकते हैं। प्राथमिकता: मैन्युअल एडिट (डिलीट बटन वाला) > फ़ाइल के \"Driver Name\" कॉलम से > ऐप की बिल्ट-इन डिफ़ॉल्ट सूची। जिस क्रू में ड्राइवर नहीं है वह खाली रह सकता है — किसी भी पंक्ति को साफ़ तौर पर खाली करने के लिए ✕ बटन इस्तेमाल करें। एक ही नाम दो अलग-अलग क्रू नंबर के लिए सेव नहीं हो सकता। मैन्युअल बदलाव केवल इसी डिवाइस/ब्राउज़र पर सहेजे जाते हैं।",
  },
  adminNoCrews: { fa: "هنوز هیچ شماره گروهی ثبت نشده.", en: "No crew numbers yet.", hi: "अभी तक कोई क्रू नंबर नहीं है।" },
  adminManualBadge: { fa: "دستی", en: "Manual", hi: "मैनुअल" },
  adminAutoBadge: { fa: "از اکسل", en: "From Excel", hi: "एक्सेल से" }, adminDefaultBadge: { fa: "فهرست پیش‌فرض", en: "Default list", hi: "डिफ़ॉल्ट सूची" },
  adminBlankBadge: { fa: "خالی", en: "Blank", hi: "खाली" },
  adminClearTooltip: { fa: "خالی کردن این گروه", en: "Clear this crew", hi: "इस क्रू को खाली करें" },
  adminDupWarning: { fa: "این اسم قبلاً برای این گروه ثبت شده و ذخیره نشد:", en: "This name is already saved for another crew and was not saved:", hi: "यह नाम पहले से एक और क्रू के लिए सेव है, इसलिए सेव नहीं हुआ:" },
  adminLogoutBtn: { fa: "خروج از حالت ادمین", en: "Log out of Admin", hi: "एडमिन से लॉग आउट" },
  crewNumberLabel: { fa: "شماره گروه", en: "Crew number", hi: "क्रू नंबर" },
  driverNameLabel: { fa: "نام راننده", en: "Driver name", hi: "ड्राइवर का नाम" },
  dailyLogMenuLabel: { fa: "دفترچه شیفت روزانه", en: "Daily Shift Log", hi: "दैनिक शिफ्ट लॉग" },
  dailyLogTitle: { fa: "دفترچه شیفت روزانه", en: "Daily Shift Log", hi: "दैनिक शिफ्ट लॉग" },
  dailyLogHint: {
    fa: "هر روز ساعت و یارد شیفتت رو ثبت کن. این اطلاعات فقط روی همین دستگاه/مرورگر ذخیره می‌شه (مثل بقیه تنظیمات این اپ) و جایی ارسال نمی‌شه. برای دادنش به کسی، از بخش «خروجی گزارش» پایین پرینت یا PDF بگیر.",
    en: "Log your shift's times and yards each day. Saved only on this device/browser (like the rest of this app's settings) — nothing is sent anywhere. To share it, use the Export report section below to print or save as PDF.",
    hi: "हर दिन अपनी शिफ्ट का समय और यार्ड लॉग करें। यह केवल इसी डिवाइस/ब्राउज़र पर सहेजा जाता है। साझा करने के लिए नीचे रिपोर्ट निर्यात अनुभाग का उपयोग करें।",
  },
  dailyLogDateLabel: { fa: "تاریخ", en: "Date", hi: "तारीख़" },
  dailyLogStartTimeLabel: { fa: "ساعت شروع شیفت", en: "Shift start time", hi: "शिफ्ट शुरू होने का समय" },
  dailyLogEndTimeLabel: { fa: "ساعت پایان شیفت", en: "Shift end time", hi: "शिफ्ट समाप्ति समय" },
  dailyLogStartYardLabel: { fa: "یارد شروع شیفت", en: "Starting yard", hi: "शुरुआती यार्ड" },
  dailyLogEndYardLabel: { fa: "یارد پایان شیفت", en: "Ending yard", hi: "समाप्ति यार्ड" },
  dailyLogDescriptionLabel: { fa: "توضیحات", en: "Description", hi: "विवरण" },
  dailyLogChipYardChange: { fa: "تغییر یارد", en: "Yard change", hi: "यार्ड परिवर्तन" },
  dailyLogChipShuttleBus: { fa: "شاتل‌باس", en: "Shuttle bus", hi: "शटल बस" },
  dailyLogChipHoliday: { fa: "تعطیلی رسمی", en: "Official holiday", hi: "आधिकारिक अवकाश" },
  dailyLogTotalHoursLabel: { fa: "جمع ساعت", en: "Total hours", hi: "कुल घंटे" },
  dailyLogSaveBtn: { fa: "ذخیره", en: "Save", hi: "सहेजें" },
  dailyLogUpdateBtn: { fa: "بروزرسانی", en: "Update", hi: "अपडेट करें" },
  cancelBtn: { fa: "انصراف", en: "Cancel", hi: "रद्द करें" },
  dailyLogEntriesTitle: { fa: "رکوردهای ثبت‌شده", en: "Saved entries", hi: "सहेजी गई प्रविष्टियाँ" },
  dailyLogImportTitle: { fa: "درون‌ریزی چندتایی", en: "Import multiple entries", hi: "कई प्रविष्टियाँ आयात करें" },
  dailyLogImportHint: {
    fa: "هر شیفت رو تو یه خط بنویس، با Tab از هم جدا: تاریخ (YYYY-MM-DD)، ساعت شروع (HH:MM)، ساعت پایان (HH:MM)، یارد شروع، یارد پایان، توضیحات. دو تای آخر اختیاری‌ان. روزهایی که از قبل رکورد دارن رد می‌شن.",
    en: "One shift per line, Tab-separated: Date (YYYY-MM-DD), Start time (HH:MM), End time (HH:MM), Start yard, End yard, Description. The last two are optional. Days that already have an entry are skipped.",
    hi: "हर शिफ्ट एक पंक्ति में, Tab से अलग: तारीख़ (YYYY-MM-DD), शुरू (HH:MM), समाप्ति (HH:MM), शुरुआती यार्ड, समाप्ति यार्ड, विवरण। आख़िरी दो वैकल्पिक हैं।",
  },
  dailyLogImportBtn: { fa: "درون‌ریزی", en: "Import", hi: "आयात करें" },
  dailyLogBackupTitle: { fa: "پشتیبان‌گیری", en: "Backup", hi: "बैकअप" },
  dailyLogBackupHint: {
    fa: "این متن رو کپی کن و یه‌جای امن (مثلاً یادداشت گوشی یا ایمیل به خودت) نگه‌دار. اگه اطلاعات این دستگاه پاک شد، همین متن رو تو کادر «درون‌ریزی چندتایی» بالا پیست کن و دکمهٔ درون‌ریزی رو بزن تا همه برگردن.",
    en: "Copy this text and keep it somewhere safe (like Notes or an email to yourself). If this device's data ever gets cleared, paste it back into the \"Import multiple entries\" box above and tap Import to restore everything.",
    hi: "इस टेक्स्ट को कॉपी करके कहीं सुरक्षित रखें (जैसे नोट्स या खुद को भेजा गया ईमेल)। अगर इस डिवाइस का डेटा मिट जाए, तो इसे ऊपर वाले \"कई प्रविष्टियाँ आयात करें\" बॉक्स में पेस्ट करके आयात करें बटन दबाएं।",
  },
  dailyLogBackupBtn: { fa: "کپی متن پشتیبان", en: "Copy backup text", hi: "बैकअप टेक्स्ट कॉपी करें" },
  dailyLogBackupCopied: { fa: "کپی شد", en: "Copied", hi: "कॉपी हो गया" },
  dailyLogBackupEmpty: { fa: "هنوز چیزی برای پشتیبان‌گیری ثبت نشده", en: "Nothing saved yet to back up", hi: "बैकअप के लिए अभी कुछ नहीं है" },
  dailyLogEmpty: { fa: "هنوز هیچ رکوردی ثبت نشده.", en: "No entries yet.", hi: "अभी तक कोई प्रविष्टि नहीं।" },
  dailyLogExportTitle: { fa: "خروجی گزارش", en: "Export report", hi: "रिपोर्ट निर्यात करें" },
  dailyLogFromLabel: { fa: "از تاریخ", en: "From", hi: "से" },
  dailyLogToLabel: { fa: "تا تاریخ", en: "To", hi: "तक" },
  dailyLogThisMonthBtn: { fa: "این ماه", en: "This month", hi: "इस महीने" },
  dailyLogReportTitle: { fa: "گزارش شیفت روزانه", en: "Daily Shift Log Report", hi: "दैनिक शिफ्ट लॉग रिपोर्ट" },
  dailyLogNoEntriesInRange: { fa: "رکوردی در این بازه پیدا نشد.", en: "No entries in this range.", hi: "इस सीमा में कोई प्रविष्टि नहीं मिली।" },
  dailyLogTotalRowLabel: { fa: "جمع کل", en: "Total", hi: "कुल" },
  dailyLogColDay: { fa: "روز هفته", en: "Weekday", hi: "सप्ताह का दिन" },
  dailyLogColNo: { fa: "ردیف", en: "No.", hi: "क्रम" },
  dailyLogSelectYardPlaceholder: { fa: "انتخاب کنید", en: "Select...", hi: "चुनें..." },
  dailyLogWeekLabel: { fa: "هفته", en: "Week", hi: "सप्ताह" },
  dailyLogWeekTotalLabel: { fa: "جمع این هفته", en: "Total this week", hi: "इस सप्ताह का कुल" },
  dailyLogMonthTotalLabel: { fa: "جمع این ماه", en: "Total this month", hi: "इस महीने का कुल" },
  adminDailyLogAccessTitle: { fa: "دسترسی به «دفترچه شیفت روزانه»", en: 'Access to "Daily Shift Log"', hi: '"दैनिक शिफ्ट लॉग" तक पहुंच' },
  adminDailyLogAccessHint: {
    fa: "شماره‌گروه‌هایی که اینجا اضافه کنی، علاوه بر ادمین، می‌تونن از «دفترچه شیفت روزانه» استفاده کنن. این هم مثل رمز ادمین فقط سمت کلاینته، یه سیستم امنیتی واقعی نیست. توجه: اطلاعات ثبت‌شده هرکس فقط روی دستگاه خودش ذخیره می‌شه، نه اینجا — برای دیدنش باید خودش خروجی پرینت/PDF بگیره و بهت بده.",
    en: "Crew numbers added here can use \"Daily Shift Log\" in addition to Admin. Like the Admin password, this is client-side only, not real security. Note: each person's entries are saved only on their own device, not here — to see them, they need to export a print/PDF and send it to you.",
    hi: "यहां जोड़े गए क्रू नंबर, एडमिन के अलावा, \"दैनिक शिफ्ट लॉग\" का उपयोग कर सकते हैं। यह केवल क्लाइंट-साइड है। नोट: हर व्यक्ति की प्रविष्टियाँ केवल उसके अपने डिवाइस पर सहेजी जाती हैं।",
  },
  adminDailyLogAccessEmpty: { fa: "هنوز کسی اضافه نشده (فقط ادمین می‌بینه)", en: "No one added yet (only Admin sees it)", hi: "अभी तक कोई नहीं जोड़ा गया (केवल एडमिन देखता है)" },
};

function t(key, lang) { return STRINGS[key] ? (STRINGS[key][lang] || STRINGS[key].en) : key; }
function regionLabel(key, lang) { return REGION_LABELS[key] ? (REGION_LABELS[key][lang] || REGION_LABELS[key].en) : key; }
function fpColor(score) { return score >= 80 ? "var(--accent)" : score >= 40 ? "#C9A227" : "#D8D3C7"; }

function openFeedbackEmail(lang) {
  const subject = `Shift Priority Feedback v${APP_VERSION}`;
  const bodyLines = [
    "",
    "",
    "---",
    `App version: ${APP_VERSION}`,
    `Language: ${lang}`,
    `Browser: ${typeof navigator !== "undefined" ? navigator.userAgent : ""}`,
  ];
  const mailto = `mailto:33omid@gmail.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join("\n"))}`;
  window.location.href = mailto;
}

function exitApp(lang) {
  if (typeof window !== "undefined" && window.electronAPI && window.electronAPI.isElectron) {
    window.electronAPI.quit();
    return;
  }
  // Web/PWA: browsers only allow closing tabs that were opened by script.
  // Try it, and if it silently does nothing, at least explain why.
  try {
    window.close();
  } catch { /* ignore */ }
  setTimeout(() => {
    if (!document.hidden) alert(t("exitWebNote", lang));
  }, 150);
}

const WEEKDAY_LABELS = {
  fa: ["یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنج‌شنبه", "جمعه", "شنبه"],
  en: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  hi: ["रविवार", "सोमवार", "मंगलवार", "बुधवार", "गुरुवार", "शुक्रवार", "शनिवार"],
};

// ---------- Theme palettes ----------

const THEME_PALETTES = {
  win11: {
    light: { bg: "#F3F3F3", card: "#FFFFFF", border: "#E1E1E1", text: "#1B1B1B", muted: "#5F6368", accent: "#0067C0", accent2: "#B98A2E", radius: "8px" },
    dark: { bg: "#202020", card: "#2C2C2C", border: "#3A3A3A", text: "#F3F3F3", muted: "#B0B0B0", accent: "#4CC2FF", accent2: "#D8A94A", radius: "8px" },
  },
  macos: {
    light: { bg: "#F5F5F7", card: "#FFFFFF", border: "#E5E5EA", text: "#1D1D1F", muted: "#6E6E73", accent: "#007AFF", accent2: "#C79A3B", radius: "14px" },
    dark: { bg: "#1E1E1E", card: "#2C2C2E", border: "#3A3A3C", text: "#F5F5F7", muted: "#98989D", accent: "#0A84FF", accent2: "#D8A94A", radius: "14px" },
  },
  universal: {
    // Neutral surfaces + a vivid emerald accent; each section (hero/priority
    // flow, My Shift, hub tiles) carries its own hue so the app isn't one colour.
    light: { bg: "#F5F6F8", card: "#FFFFFF", border: "#E4E7EC", text: "#1B2230", muted: "#667085", accent: "#0B7D61", accent2: "#F59F00", radius: "10px" },
    dark: { bg: "#111418", card: "#1A1F26", border: "#2A313B", text: "#ECEFF3", muted: "#9AA4B2", accent: "#34D3A0", accent2: "#FBBF24", radius: "10px" },
  },
};

function getPalette(style, mode) {
  const s = THEME_PALETTES[style] ? style : "universal";
  return THEME_PALETTES[s][mode === "dark" ? "dark" : "light"];
}

// ---------- Small components ----------

function RegionChips({ regionSummary, lang }) {
  const sep = lang === "fa" ? "، " : ", ";
  return (
    <div style={styles.regionChipsRow}>
      {Object.entries(regionSummary).map(([key, hoursArr]) => (
        <span key={key} style={styles.regionChip}>
          <span style={{ ...styles.regionDot, background: REGION_COLORS[key] || "#9AA0A6" }} />
          {hoursArr.length} {regionLabel(key, lang)} ({hoursArr.join(sep)} {t("hoursWord", lang)})
        </span>
      ))}
    </div>
  );
}

function CheckLine({ text }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "var(--accent)", marginBottom: 4 }}>
      <span style={{ width: 5, height: 5, borderRadius: "50%", background: "var(--accent)", flexShrink: 0 }} />
      {text}
    </div>
  );
}

function FingerprintStrip({ fingerprint }) {
  return (
    <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginBottom: 8 }}>
      {fingerprint.map((f, i) => (
        <span key={f.id} title={`${i + 1}. ${f.score}%`} style={{ ...styles.fpSquare, background: fpColor(f.score) }}>
          {i + 1}
        </span>
      ))}
    </div>
  );
}

function FingerprintList({ fingerprint, lang }) {
  return (
    <div style={{ marginBottom: 10 }}>
      {fingerprint.map((f, i) => (
        <div key={f.id} style={styles.fpRow}>
          <span style={styles.fpRank}>{i + 1}</span>
          <span style={styles.fpLabel}>{f.label[lang]}</span>
          <div style={styles.fpBarBg}>
            <div style={{ ...styles.fpBarFill, width: `${f.score}%`, background: fpColor(f.score) }} />
          </div>
          <span style={styles.fpPct}>{f.score}%</span>
        </div>
      ))}
    </div>
  );
}

// ---------- Date/time ----------

const FA_WEEKDAYS = ["یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه"];
const FA_MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"];
const FA_DIGITS = ["۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹"];

function toFaDigits(n) {
  return String(n).replace(/[0-9]/g, (d) => FA_DIGITS[+d]);
}

function gregorianToJalali(gy, gm, gd) {
  const gDaysInMonth = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
  let jy;
  if (gy > 1600) { jy = 979; gy -= 1600; } else { jy = 0; gy -= 621; }
  const gy2 = gm > 2 ? gy + 1 : gy;
  let days = 365 * gy + Math.floor((gy2 + 3) / 4) - Math.floor((gy2 + 99) / 100) + Math.floor((gy2 + 399) / 400) - 80 + gd + gDaysInMonth[gm - 1];
  jy += 33 * Math.floor(days / 12053);
  days %= 12053;
  jy += 4 * Math.floor(days / 1461);
  days %= 1461;
  if (days > 365) {
    jy += Math.floor((days - 1) / 365);
    days = (days - 1) % 365;
  }
  let jm, jd;
  if (days < 186) {
    jm = 1 + Math.floor(days / 31);
    jd = 1 + (days % 31);
  } else {
    jm = 7 + Math.floor((days - 186) / 30);
    jd = 1 + ((days - 186) % 30);
  }
  return [jy, jm, jd];
}

function formatJalaliDate(date) {
  const [jy, jm, jd] = gregorianToJalali(date.getFullYear(), date.getMonth() + 1, date.getDate());
  const weekday = FA_WEEKDAYS[date.getDay()];
  return `${weekday} ${toFaDigits(jd)} ${FA_MONTHS[jm - 1]} ${toFaDigits(jy)}`;
}

function AnalogClock({ hourAngle, minuteAngle, secondAngle, size = 64 }) {
  const c = size / 2;
  const rad = (deg) => (deg * Math.PI) / 180;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <circle cx={c} cy={c} r={c - 2} fill="var(--card)" stroke="var(--border)" strokeWidth="2" />
      {[...Array(12)].map((_, i) => {
        const a = rad(i * 30);
        const x1 = c + Math.sin(a) * (c - 7), y1 = c - Math.cos(a) * (c - 7);
        const x2 = c + Math.sin(a) * (c - 3), y2 = c - Math.cos(a) * (c - 3);
        return <line key={i} x1={x1} y1={y1} x2={x2} y2={y2} stroke="var(--border)" strokeWidth="1.5" />;
      })}
      <line x1={c} y1={c} x2={c + Math.sin(rad(hourAngle)) * (c * 0.45)} y2={c - Math.cos(rad(hourAngle)) * (c * 0.45)} stroke="var(--text)" strokeWidth="3" strokeLinecap="round" />
      <line x1={c} y1={c} x2={c + Math.sin(rad(minuteAngle)) * (c * 0.68)} y2={c - Math.cos(rad(minuteAngle)) * (c * 0.68)} stroke="var(--accent)" strokeWidth="2.5" strokeLinecap="round" />
      <line x1={c} y1={c} x2={c + Math.sin(rad(secondAngle)) * (c * 0.75)} y2={c - Math.cos(rad(secondAngle)) * (c * 0.75)} stroke="var(--accent2)" strokeWidth="1" strokeLinecap="round" />
      <circle cx={c} cy={c} r="2.5" fill="var(--text)" />
    </svg>
  );
}

function DateTimeWidget({ lang }) {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);

  const enDate = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric", weekday: "long" }).format(now);
  const primaryDate =
    lang === "fa" ? formatJalaliDate(now)
    : lang === "hi" ? new Intl.DateTimeFormat("hi-IN", { year: "numeric", month: "long", day: "numeric", weekday: "long" }).format(now)
    : enDate;
  const digital = now.toLocaleTimeString("en-GB");

  const hours = now.getHours() % 12;
  const minutes = now.getMinutes();
  const seconds = now.getSeconds();
  const hourAngle = (hours + minutes / 60) * 30;
  const minuteAngle = (minutes + seconds / 60) * 6;
  const secondAngle = seconds * 6;

  return (
    <div style={{ ...styles.dtWidget, background: "linear-gradient(120deg, #6D5CE0 0%, #B15CE0 26%, #E0447F 52%, #E08A1E 76%, #0EA37E 100%)", border: "none", boxShadow: "0 6px 14px rgba(96,60,180,0.28)", "--card": "rgba(255,255,255,0.16)", "--border": "rgba(255,255,255,0.55)", "--text": "#fff", "--muted": "rgba(255,255,255,0.85)", "--accent": "#fff", "--accent2": "#FFE9A8" }}>
      <AnalogClock hourAngle={hourAngle} minuteAngle={minuteAngle} secondAngle={secondAngle} />
      <div style={styles.dtTextCol}>
        <div style={styles.dtDigital}>{digital}</div>
        <div style={styles.dtDateLine}>{primaryDate}</div>
        {lang !== "en" && <div style={styles.dtDateLine}>{enDate}</div>}
      </div>
    </div>
  );
}

// ---------- Modal shell ----------

function Modal({ title, onClose, onBack, children, headerGradient, accent }) {
  const headerStyle = headerGradient ? { ...styles.modalHeader, background: headerGradient, borderBottom: "none" } : styles.modalHeader;
  const titleStyle = headerGradient ? { ...styles.modalTitle, color: "#fff" } : styles.modalTitle;
  const closeBtnStyle = headerGradient ? { ...styles.modalCloseBtn, color: "#fff", background: "rgba(255,255,255,0.18)", borderRadius: 8, width: 28, height: 28, alignItems: "center", justifyContent: "center" } : styles.modalCloseBtn;
  const boxVars = accent ? { "--accent": accent } : {};
  return (
    <div className="no-print" style={styles.modalOverlay} onClick={onClose}>
      <div style={{ ...styles.modalBox, ...boxVars }} onClick={(e) => e.stopPropagation()}>
        <div style={headerStyle}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
            {onBack && (
              <button onClick={onBack} style={closeBtnStyle} title="Back"><ArrowLeft size={16} /></button>
            )}
            <span style={titleStyle}>{title}</span>
          </div>
          <button onClick={onClose} style={closeBtnStyle}><X size={16} /></button>
        </div>
        <div style={styles.modalBody}>{children}</div>
      </div>
    </div>
  );
}

// ---------- Crew name directory ----------
// crewNumber (string) -> driver name. A manual entry here always wins over
// a name auto-read from the uploaded file's "Driver Name" column — see
// resolveCrewName() below. That's how Admin fixes a typo, or adds a name
// for a crew number the current file has no name column for at all.
function loadCrewNames() {
  try { return JSON.parse(localStorage.getItem(CREW_NAMES_KEY) || "{}"); } catch { return {}; }
}
function saveCrewNames(map) {
  try { localStorage.setItem(CREW_NAMES_KEY, JSON.stringify(map)); } catch { /* ignore storage errors */ }
}
function resolveCrewName(crewNumber, crews, manualNames) {
  const key = String(crewNumber);
  if (manualNames && Object.prototype.hasOwnProperty.call(manualNames, key)) return manualNames[key];
  const c = crews?.find((cc) => String(cc.crew) === key);
  return c?.driverName || CREW_NAME_DEFAULTS[key] || "";
}

// ---------- Admin session ----------
// sessionStorage only — see ADMIN_SESSION_KEY above for why.
function loadAdminSession() {
  try { return sessionStorage.getItem(ADMIN_SESSION_KEY) === "1"; } catch { return false; }
}
function saveAdminSession(isAdmin) {
  try {
    if (isAdmin) sessionStorage.setItem(ADMIN_SESSION_KEY, "1");
    else sessionStorage.removeItem(ADMIN_SESSION_KEY);
  } catch { /* ignore storage errors */ }
}

// ---------- Profile (name + "my crew number") ----------
// Same storage mechanism as history above — plain localStorage, private to
// this browser/device. Kept as its own key so clearing history never
// touches the saved profile and vice versa.
function loadProfile() {
  try { return JSON.parse(localStorage.getItem(PROFILE_KEY) || "null"); } catch { return null; }
}
function saveProfile(profile) {
  try { localStorage.setItem(PROFILE_KEY, JSON.stringify(profile)); } catch { /* ignore storage errors */ }
}
function clearProfileStorage() {
  try { localStorage.removeItem(PROFILE_KEY); } catch { /* ignore */ }
}

// ---------- Last parsed schedule (avoids re-uploading every visit) ----------
function loadLastFile() {
  try { return JSON.parse(localStorage.getItem(LAST_FILE_KEY) || "null"); } catch { return null; }
}
function saveLastFile(entry) {
  try { localStorage.setItem(LAST_FILE_KEY, JSON.stringify(entry)); } catch { /* ignore storage errors, e.g. quota */ }
}
function clearLastFileStorage() {
  try { localStorage.removeItem(LAST_FILE_KEY); } catch { /* ignore */ }
}

// ---------- Daily Log (per-device only, not synced anywhere) ----------
const DAILY_LOG_ENTRIES_KEY = "shiftPriorityDailyLogEntries";
const DAILY_LOG_ACCESS_KEY = "shiftPriorityDailyLogAccess";
function loadDailyLogEntries() {
  try { const v = JSON.parse(localStorage.getItem(DAILY_LOG_ENTRIES_KEY) || "[]"); return Array.isArray(v) ? v : []; } catch { return []; }
}
function saveDailyLogEntries(list) {
  try { localStorage.setItem(DAILY_LOG_ENTRIES_KEY, JSON.stringify(list)); } catch { /* ignore storage errors */ }
}
function loadDailyLogAccess() {
  try { const v = JSON.parse(localStorage.getItem(DAILY_LOG_ACCESS_KEY) || "[]"); return Array.isArray(v) ? v : []; } catch { return []; }
}
function saveDailyLogAccess(list) {
  try { localStorage.setItem(DAILY_LOG_ACCESS_KEY, JSON.stringify(list)); } catch { /* ignore storage errors */ }
}
function computeLogHours(startTime, endTime) {
  if (!startTime || !endTime) return 0;
  const [sh, sm] = startTime.split(":").map(Number);
  const [eh, em] = endTime.split(":").map(Number);
  if ([sh, sm, eh, em].some((n) => Number.isNaN(n))) return 0;
  let mins = (eh * 60 + em) - (sh * 60 + sm);
  if (mins <= 0) mins += 24 * 60;
  return Math.round((mins / 60) * 100) / 100;
}

// Fixed set of TOK Transit yards for the Daily Log start/end-yard pickers
// (a plain select beats free text: no typos, easy to scan on the printed
// report). A legacy/free-typed value that isn't one of these three is
// still kept as an extra option so older entries never lose their data.
const DAILY_LOG_YARDS = ["Newmarket", "Richmond Hill", "Caldari"];

// Gregorian month names for the Daily Log's month-group headers. Entries
// are stored and shown with Gregorian dates (matching the <input type="date">
// picker), so this stays Gregorian too rather than switching to the Jalali
// calendar used elsewhere in the app for fa -- mixing the two would make the
// month header disagree with the day rows underneath it.
const GREGORIAN_MONTH_NAMES = {
  fa: ["ژانویه", "فوریه", "مارس", "آوریل", "مه", "ژوئن", "ژوئیه", "اوت", "سپتامبر", "اکتبر", "نوامبر", "دسامبر"],
  en: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
  hi: ["जनवरी", "फरवरी", "मार्च", "अप्रैल", "मई", "जून", "जुलाई", "अगस्त", "सितंबर", "अक्टूबर", "नवंबर", "दिसंबर"],
};
function formatMonthYear(dateStr, lang) {
  const [y, m] = dateStr.split("-").map(Number);
  const names = GREGORIAN_MONTH_NAMES[lang] || GREGORIAN_MONTH_NAMES.en;
  const monthName = names[m - 1] || "";
  const yearStr = lang === "fa" ? toFaDigits(y) : String(y);
  return `${monthName} ${yearStr}`;
}
// Sunday-Saturday week-of-month index, 1-based, matching how a wall
// calendar (and Omid's own paper log) breaks a month into weeks: the
// first week can be a short partial week if the month doesn't start on
// a Sunday.
function weekOfMonth(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  const firstWeekday = new Date(d.getFullYear(), d.getMonth(), 1).getDay();
  return Math.ceil((d.getDate() + firstWeekday) / 7);
}

function ProfilePanel({ lang, profile, onSave, onClear, onClose }) {
  const [firstName, setFirstName] = useState(profile?.firstName || "");
  const [crewNumber, setCrewNumber] = useState(profile?.crewNumber || "");

  const handleSave = () => {
    const name = firstName.trim();
    const crew = String(crewNumber).trim();
    if (!name && !crew) { onClear(); onClose(); return; }
    onSave({ firstName: name, crewNumber: crew });
    onClose();
  };

  return (
    <Modal title={t("profileTitle", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #E0447F, #F17CA6)" accent="#E0447F">
      <div style={styles.smallLabel}>{t("firstNameLabel", lang)}</div>
      <input
        type="text"
        value={firstName}
        onChange={(e) => setFirstName(e.target.value)}       
        style={styles.numInputWide}
      />
      <div style={{ ...styles.smallLabel, marginTop: 10 }}>{t("myCrewLabel", lang)}</div>
      <input
        type="number"
        value={crewNumber}
        onChange={(e) => setCrewNumber(e.target.value)}
        style={styles.numInputWide}
      />
      <p style={styles.hint}>{t("profilePrivacyHint", lang)}</p>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <button onClick={handleSave} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" }}>
          <Check size={14} /> {t("saveProfileBtn", lang)}
        </button>
        {(profile?.firstName || profile?.crewNumber) && (
          <button onClick={() => { onClear(); onClose(); }} style={{ ...styles.smallActionBtn, color: "#B3432A" }}>
            <Trash2 size={14} /> {t("clearProfileBtn", lang)}
          </button>
        )}
      </div>
    </Modal>
  );
}

// A read-only, single-crew weekly view for "my shift" — deliberately its
// own small component instead of reusing CompareTwoPanel's table (which is
// built around 2-5 typed-in crew numbers), so this stays a pure addition
// that can't regress the existing compare feature.
function MyScheduleModal({ crew, name, lang, themeMode, onClose, onBack }) {
  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const dayCell = (i) => crew.days.find((d) => d.dayIdx === i);
  // Today's weekday index (0=Sunday, matching dayIdx/Date.getDay()), so we can
  // highlight the current day's row below.
  const todayIdx = new Date().getDay();
  // Today's highlight color: neon lime pops on dark backgrounds but is
  // nearly invisible on light ones, so light mode gets a high-contrast
  // orange instead, same treatment (border + glow + badge), different hue.
  const todayTheme = themeMode === "dark" ? { accent: "#CCFF00", ring: "rgba(204,255,0,0.22)", glow: "rgba(204,255,0,0.6)", badgeGlow: "rgba(204,255,0,0.7)", badgeText: "#111" } : { accent: "#FF6A00", ring: "rgba(255,106,0,0.22)", glow: "rgba(255,106,0,0.5)", badgeGlow: "rgba(255,106,0,0.55)", badgeText: "#fff" };
  return (
        <Modal title={`${t("myScheduleTitle", lang)} — ${t("crewWord", lang)} ${String(crew.crew)}${name ? " · " + name : ""}`} onClose={onClose} onBack={onBack} headerGradient="linear-gradient(135deg, #0B8F87, #19B97A)" accent="#0B8F87">
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {weekdayNames.map((wd, i) => {
          const d = dayCell(i);
        const isToday = i === todayIdx;
          return (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, borderRadius: 8, padding: isToday ? "4px 6px" : "4px 0", border: isToday ? `2px solid ${todayTheme.accent}` : "2px solid transparent", boxShadow: isToday ? `0 0 0 3px ${todayTheme.ring}, 0 0 16px ${todayTheme.glow}` : "none", background: isToday ? "var(--card)" : "transparent" }}>
                                                        <span style={{ width: 88, flexShrink: 0, fontWeight: 700, fontSize: 12.5, display: "flex", alignItems: "center", gap: 5 }}>{wd}{isToday && (<span style={{ fontSize: 9.5, fontWeight: 800, color: todayTheme.badgeText, background: todayTheme.accent, borderRadius: 999, padding: "1.5px 7px", whiteSpace: "nowrap", boxShadow: `0 0 6px ${todayTheme.badgeGlow}` }}>{t("todayLabel", lang)}</span>)}</span>
              {d ? (
                <div style={{ flex: 1, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", background: "var(--bg)", border: "1px solid var(--border)", borderRadius: 8, padding: "7px 10px" }}>
                  <span style={{ width: 8, height: 8, borderRadius: "50%", background: REGION_COLORS[d.regionKey] || "#9AA0A6", flexShrink: 0 }} />
                  <span style={{ fontSize: 12 }}>{regionLabel(d.regionKey, lang)} · {d.code || "-"}</span>
                  <span style={{ marginInlineStart: "auto", fontWeight: 700, fontSize: 12.5 }}>{formatExcelTime(d.start)}–{formatExcelTime(d.end)}</span>
                  <span style={{ fontSize: 11, color: "var(--muted)" }}>{formatDuration(d.hours, lang)}</span>
                </div>
              ) : (
                <div style={{ flex: 1, border: "1.5px dashed #e0a0a0", background: "#fbeceb", color: "#b3432a", borderRadius: 8, padding: "7px 10px", fontWeight: 700, fontSize: 12 }}>
                  {t("off", lang)} ✕
                </div>
              )}
            </div>
          );
        })}
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 13, fontWeight: 700, borderTop: "1px solid var(--border)", paddingTop: 10 }}>
          <span>{crew.type} · {crew.shiftRaw}</span>
          <span>{crew.totalHours} {t("hours", lang)}</span>
        </div>
      </div>
    </Modal>
  );
}

// ---------- Daily Log ----------

function DailyLogPanel({ lang, onClose }) {
  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const todayStr = () => new Date().toISOString().slice(0, 10);
  const [entries, setEntries] = useState(() => loadDailyLogEntries());
  const [editingId, setEditingId] = useState(null);
  const lastEntry = entries.length ? entries[entries.length - 1] : null;
  const [date, setDate] = useState(todayStr());
  const [startTime, setStartTime] = useState(lastEntry?.startTime || "");
  const [endTime, setEndTime] = useState(lastEntry?.endTime || "");
  const [startYard, setStartYard] = useState(lastEntry?.startYard || "");
  const [endYard, setEndYard] = useState(lastEntry?.endYard || "");
  const [description, setDescription] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [importText, setImportText] = useState("");
  const [importResult, setImportResult] = useState(null);
  const [backupCopied, setBackupCopied] = useState(false);

  const dayIdx = useMemo(() => {
    const d = new Date(date + "T00:00:00");
    return isNaN(d.getTime()) ? null : d.getDay();
  }, [date]);

  const totalHours = useMemo(() => computeLogHours(startTime, endTime), [startTime, endTime]);

  // Entries grouped by calendar month, then by Sunday-Saturday week within
  // that month, each carrying its own subtotal -- mirrors the weekly-table
  // layout of Omid's own paper work log, so the on-screen list reads the
  // same way as the sheet he cross-checks against his paystub.
  const monthGroups = useMemo(() => {
    const byMonth = new Map();
    entries.forEach((e) => {
      const monthKey = e.date.slice(0, 7);
      if (!byMonth.has(monthKey)) byMonth.set(monthKey, []);
      byMonth.get(monthKey).push(e);
    });
    return Array.from(byMonth.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([monthKey, monthEntries]) => {
        const byWeek = new Map();
        monthEntries.forEach((e) => {
          const wk = weekOfMonth(e.date);
          if (!byWeek.has(wk)) byWeek.set(wk, []);
          byWeek.get(wk).push(e);
        });
        const weeks = Array.from(byWeek.entries())
          .sort((a, b) => a[0] - b[0])
          .map(([weekIdx, weekEntries]) => ({
            weekIdx,
            entries: weekEntries,
            totalHours: weekEntries.reduce((s, e) => s + (e.totalHours || 0), 0),
          }));
        return {
          monthKey,
          sampleDate: monthEntries[0].date,
          weeks,
          monthTotalHours: monthEntries.reduce((s, e) => s + (e.totalHours || 0), 0),
        };
      });
  }, [entries]);

  const resetFormAfterSave = () => {
    setDate(todayStr());
    setDescription("");
    setEditingId(null);
  };

  const handleSave = () => {
    if (!date || !startTime || !endTime) return;
    const entry = {
      id: editingId || String(Date.now()),
      date, startTime, endTime,
      startYard: startYard.trim(), endYard: endYard.trim(),
      description: description.trim(),
      totalHours,
    };
    const next = editingId
      ? entries.map((e) => (e.id === editingId ? entry : e))
      : [...entries, entry];
    next.sort((a, b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
    setEntries(next);
    saveDailyLogEntries(next);
    resetFormAfterSave();
  };

  const handleEdit = (entry) => {
    setEditingId(entry.id);
    setDate(entry.date);
    setStartTime(entry.startTime);
    setEndTime(entry.endTime);
    setStartYard(entry.startYard);
    setEndYard(entry.endYard);
    setDescription(entry.description);
  };

  const handleCancelEdit = () => {
    setEditingId(null);
    setDate(todayStr());
    setDescription("");
  };

  // Bulk-import: one shift per pasted line, Tab-separated (matches what a
  // spreadsheet gives you on copy). Skips a line whose date already has an
  // entry, rather than overwriting it -- the single-entry form + edit is
  // still the way to fix or replace one day.
  const handleImport = () => {
    const lines = importText.split("\n").map((l) => l.trim()).filter(Boolean);
    const existingDates = new Set(entries.map((e) => e.date));
    const imported = [];
    let skipped = 0;
    lines.forEach((line, i) => {
      const cols = line.split("\t").map((c) => c.trim());
      const [rDate, rStart, rEnd, rStartYard, rEndYard, rDescription] = cols;
      if (!rDate || !rStart || !rEnd || existingDates.has(rDate)) { skipped += 1; return; }
      imported.push({
        id: `${Date.now()}_${i}`,
        date: rDate, startTime: rStart, endTime: rEnd,
        startYard: rStartYard || "", endYard: rEndYard || "",
        description: rDescription || "",
        totalHours: computeLogHours(rStart, rEnd),
      });
      existingDates.add(rDate);
    });
    if (imported.length) {
      const next = [...entries, ...imported].sort((a, b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
      setEntries(next);
      saveDailyLogEntries(next);
    }
    setImportText("");
    setImportResult({ added: imported.length, skipped });
  };

  const handleDelete = (id) => {
    const next = entries.filter((e) => e.id !== id);
    setEntries(next);
    saveDailyLogEntries(next);
    if (editingId === id) handleCancelEdit();
  };

  const applyChip = (label) => {
    setDescription((prev) => {
      const parts = prev.split(",").map((p) => p.trim()).filter(Boolean);
      const idx = parts.indexOf(label);
      if (idx >= 0) parts.splice(idx, 1);
      else parts.push(label);
      return parts.join(", ");
    });
  };

  const setThisMonth = () => {
    const now = new Date();
    const first = new Date(now.getFullYear(), now.getMonth(), 1);
    const last = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    setFromDate(first.toISOString().slice(0, 10));
    setToDate(last.toISOString().slice(0, 10));
  };

  const handleExport = () => {
    const filtered = entries.filter((e) => (!fromDate || e.date >= fromDate) && (!toDate || e.date <= toDate));
    printDailyLogTable(filtered, lang);
  };

  // Plain-text, tab-separated backup of every saved entry -- same format
  // the bulk-import box above accepts, so copying this out and pasting it
  // back in is a full restore. This is the safety net for local storage
  // getting cleared (e.g. iOS Safari evicting site data after a period of
  // disuse) -- nothing here is sent anywhere, it just leaves the device
  // through the person's own copy/paste.
  const buildBackupText = () => entries.map((e) => [e.date, e.startTime, e.endTime, e.startYard || "", e.endYard || "", e.description || ""].join("\t")).join("\n");

  const handleCopyBackup = async () => {
    const text = buildBackupText();
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      setBackupCopied(true);
      setTimeout(() => setBackupCopied(false), 2000);
    } catch {
      // Clipboard API blocked/unavailable in this context (happens on some
      // iOS Safari setups) -- the textarea below is still visible and can
      // be selected and copied by hand.
    }
  };

  return (
    <Modal title={t("dailyLogTitle", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #2463EB, #4F8CFB)" accent="#2463EB">
      <p style={styles.hint}>{t("dailyLogHint", lang)}</p>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 140px" }}>
          <div style={styles.smallLabel}>{t("dailyLogDateLabel", lang)}{dayIdx !== null ? ` · ${weekdayNames[dayIdx]}` : ""}</div>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={styles.numInputWide} />
        </div>
        <div style={{ flex: "1 1 100px" }}>
          <div style={styles.smallLabel}>{t("dailyLogStartTimeLabel", lang)}</div>
          <input type="time" value={startTime} onChange={(e) => setStartTime(e.target.value)} style={styles.numInputWide} />
        </div>
        <div style={{ flex: "1 1 100px" }}>
          <div style={styles.smallLabel}>{t("dailyLogEndTimeLabel", lang)}</div>
          <input type="time" value={endTime} onChange={(e) => setEndTime(e.target.value)} style={styles.numInputWide} />
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <div style={{ flex: 1 }}>
          <div style={styles.smallLabel}>{t("dailyLogStartYardLabel", lang)}</div>
          <select value={startYard} onChange={(e) => setStartYard(e.target.value)} style={styles.numInputWide}>
            <option value="">{t("dailyLogSelectYardPlaceholder", lang)}</option>
            {DAILY_LOG_YARDS.map((y) => <option key={y} value={y}>{y}</option>)}
            {startYard && !DAILY_LOG_YARDS.includes(startYard) && <option value={startYard}>{startYard}</option>}
          </select>
        </div>
        <div style={{ flex: 1 }}>
          <div style={styles.smallLabel}>{t("dailyLogEndYardLabel", lang)}</div>
          <select value={endYard} onChange={(e) => setEndYard(e.target.value)} style={styles.numInputWide}>
            <option value="">{t("dailyLogSelectYardPlaceholder", lang)}</option>
            {DAILY_LOG_YARDS.map((y) => <option key={y} value={y}>{y}</option>)}
            {endYard && !DAILY_LOG_YARDS.includes(endYard) && <option value={endYard}>{endYard}</option>}
          </select>
        </div>
      </div>

      <div style={{ ...styles.smallLabel, marginTop: 10 }}>{t("dailyLogDescriptionLabel", lang)}</div>
      <div style={{ ...styles.chipRow, marginBottom: 6 }}>
        {[t("dailyLogChipYardChange", lang), t("dailyLogChipShuttleBus", lang), t("dailyLogChipHoliday", lang)].map((label) => {
          const active = description.split(",").map((p) => p.trim()).includes(label);
          return (
            <button
              key={label}
              type="button"
              style={active ? { ...styles.chip, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" } : styles.chip}
              onClick={() => applyChip(label)}
            >
              {label}
            </button>
          );
        })}
      </div>
      <textarea
        value={description}
        onChange={(e) => setDescription(e.target.value)}
        rows={2}
        style={{ ...styles.numInputWide, width: "100%", resize: "vertical", fontFamily: "inherit", boxSizing: "border-box" }}
      />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10, background: "var(--bg)", borderRadius: 8, padding: "8px 12px" }}>
        <span style={{ fontSize: 12.5, color: "var(--muted)" }}>{t("dailyLogTotalHoursLabel", lang)}</span>
        <span style={{ fontSize: 15, fontWeight: 700 }}>{formatDuration(totalHours, lang)}</span>
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <button
          onClick={handleSave}
          disabled={!date || !startTime || !endTime}
          style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)", opacity: (!date || !startTime || !endTime) ? 0.5 : 1 }}
        >
          <Check size={14} /> {editingId ? t("dailyLogUpdateBtn", lang) : t("dailyLogSaveBtn", lang)}
        </button>
        {editingId && (
          <button onClick={handleCancelEdit} style={styles.smallActionBtn}>
            <X size={14} /> {t("cancelBtn", lang)}
          </button>
        )}
      </div>

      <div style={{ borderTop: "1px solid var(--border)", marginTop: 16, paddingTop: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 4 }}>{t("dailyLogBackupTitle", lang)}</div>
        <p style={styles.hint}>{t("dailyLogBackupHint", lang)}</p>
        <textarea
          readOnly
          value={buildBackupText()}
          placeholder={t("dailyLogBackupEmpty", lang)}
          onFocus={(e) => e.target.select()}
          rows={3}
          style={{ ...styles.numInputWide, width: "100%", resize: "vertical", fontFamily: "monospace", fontSize: 11.5, boxSizing: "border-box" }}
        />
        <button
          onClick={handleCopyBackup}
          disabled={!entries.length}
          style={{ ...styles.smallActionBtn, marginTop: 8, opacity: entries.length ? 1 : 0.5 }}
        >
          <ClipboardList size={14} /> {backupCopied ? t("dailyLogBackupCopied", lang) : t("dailyLogBackupBtn", lang)}
        </button>
      </div>

      <div style={{ borderTop: "1px solid var(--border)", marginTop: 16, paddingTop: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 4 }}>{t("dailyLogImportTitle", lang)}</div>
        <p style={styles.hint}>{t("dailyLogImportHint", lang)}</p>
        <textarea
          value={importText}
          onChange={(e) => setImportText(e.target.value)}
          rows={4}
          placeholder="2026-09-05\t07:30\t15:50\tRichmond Hill\tRichmond Hill"
          style={{ ...styles.numInputWide, width: "100%", resize: "vertical", fontFamily: "monospace", fontSize: 11.5, boxSizing: "border-box" }}
        />
        <button
          onClick={handleImport}
          disabled={!importText.trim()}
          style={{ ...styles.smallActionBtn, marginTop: 8, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)", opacity: importText.trim() ? 1 : 0.5 }}
        >
          <Upload size={14} /> {t("dailyLogImportBtn", lang)}
        </button>
        {importResult && (
          <p style={styles.hint}>
            {lang === "fa"
              ? `${toFaDigits(importResult.added)} رکورد اضافه شد` + (importResult.skipped ? ` (${toFaDigits(importResult.skipped)} خط رد شد)` : "")
              : lang === "hi"
              ? `${importResult.added} प्रविष्टियाँ जोड़ी गईं` + (importResult.skipped ? ` (${importResult.skipped} पंक्तियाँ छोड़ी गईं)` : "")
              : `${importResult.added} entries added` + (importResult.skipped ? ` (${importResult.skipped} skipped)` : "")}
          </p>
        )}
      </div>

      <div style={{ borderTop: "1px solid var(--border)", marginTop: 16, paddingTop: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 8 }}>{t("dailyLogEntriesTitle", lang)}</div>
        {entries.length === 0 && <p style={styles.hint}>{t("dailyLogEmpty", lang)}</p>}
        {monthGroups.map((mg) => {
          let rowNo = 0;
          return (
            <div key={mg.monthKey} style={{ marginBottom: 16 }}>
              <div style={{ fontWeight: 800, fontSize: 12.5, marginBottom: 6, color: "var(--accent)" }}>
                {formatMonthYear(mg.sampleDate, lang)}
              </div>
              <div style={{ overflowX: "auto" }}>
                {mg.weeks.map((wk) => (
                  <div key={wk.weekIdx} style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 11.5, fontWeight: 700, color: "var(--muted)", margin: "6px 0 4px" }}>
                      {t("dailyLogWeekLabel", lang)} {lang === "fa" ? toFaDigits(wk.weekIdx) : wk.weekIdx}
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, minWidth: 480 }}>
                      <thead>
                        <tr>
                          <th style={styles.dailyLogTh}>{t("dailyLogColNo", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogColDay", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogDateLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogStartTimeLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogEndTimeLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogStartYardLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogEndYardLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogDescriptionLabel", lang)}</th>
                          <th style={styles.dailyLogTh}>{t("dailyLogTotalHoursLabel", lang)}</th>
                          <th style={styles.dailyLogTh}></th>
                        </tr>
                      </thead>
                      <tbody>
                        {wk.entries.map((e) => {
                          rowNo += 1;
                          return (
                            <tr key={e.id}>
                              <td style={styles.dailyLogTd}>{lang === "fa" ? toFaDigits(rowNo) : rowNo}</td>
                              <td style={styles.dailyLogTd}>{weekdayNames[new Date(e.date + "T00:00:00").getDay()]}</td>
                              <td style={styles.dailyLogTd}>{e.date}</td>
                              <td style={styles.dailyLogTd}>{e.startTime}</td>
                              <td style={styles.dailyLogTd}>{e.endTime}</td>
                              <td style={styles.dailyLogTd}>{e.startYard || "-"}</td>
                              <td style={styles.dailyLogTd}>{e.endYard || "-"}</td>
                              <td style={{ ...styles.dailyLogTd, maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis" }} title={e.description || undefined}>{e.description || "-"}</td>
                              <td style={styles.dailyLogTd}>{formatDuration(e.totalHours, lang)}</td>
                              <td style={{ ...styles.dailyLogTd, whiteSpace: "nowrap" }}>
                                <button onClick={() => handleEdit(e)} style={{ ...styles.smallActionBtn, padding: "3px 5px" }}><Pencil size={11} /></button>
                                <button onClick={() => handleDelete(e.id)} style={{ ...styles.smallActionBtn, padding: "3px 5px", color: "#B3432A" }}><Trash2 size={11} /></button>
                              </td>
                            </tr>
                          );
                        })}
                        <tr>
                          <td colSpan={8} style={{ ...styles.dailyLogTd, fontWeight: 700, background: "var(--bg)", textAlign: lang === "fa" ? "left" : "right" }}>{t("dailyLogWeekTotalLabel", lang)}</td>
                          <td colSpan={2} style={{ ...styles.dailyLogTd, fontWeight: 700, background: "var(--bg)" }}>{formatDuration(wk.totalHours, lang)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4, padding: "8px 10px", background: "var(--card)", border: "1px solid var(--border)", borderRadius: 8, fontWeight: 800, fontSize: 12.5 }}>
                <span>{t("dailyLogMonthTotalLabel", lang)}</span>
                <span>{formatDuration(mg.monthTotalHours, lang)}</span>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ borderTop: "1px solid var(--border)", marginTop: 16, paddingTop: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 4 }}>{t("dailyLogExportTitle", lang)}</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 130 }}>
            <div style={styles.smallLabel}>{t("dailyLogFromLabel", lang)}</div>
            <input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} style={styles.numInputWide} />
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <div style={styles.smallLabel}>{t("dailyLogToLabel", lang)}</div>
            <input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} style={styles.numInputWide} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={setThisMonth} style={styles.smallActionBtn}>{t("dailyLogThisMonthBtn", lang)}</button>
          <button onClick={handleExport} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" }}>
            <Printer size={14} /> {t("printBtn", lang)}
          </button>
        </div>
      </div>
    </Modal>
  );
}

// ---------- Help ----------

const HELP_CONTENT = {
  fa: [
    ["فایل اکسل رو آپلود کن", "شیت مربوطه رو انتخاب کن (اگه چند شیت داشت). برنامه خودش ستون‌های گروه، نوع، شیفت، روزها و ساعت رو پیدا می‌کنه."],
    ["اولویت‌هاتو بچین", "از فهرست معیارها (شیفت، روز کاری، منطقه، آماده‌باش) به ترتیب اهمیت کلیک کن. رتبه‌بندی دقیقاً طبق همین ترتیبه: اول معیار ۱، تساوی رو معیار ۲ می‌شکنه، و همین‌طور."],
    ["نتیجه رو ببین", "پنج گزینه برتر با مدال طلا/نقره/برنز نشون داده می‌شن. عدد سمت چپ هر کارت، درصد تطابق کلی اون گروه با اولویت‌های توئه."],
    ["مقایسه و خروجی", "چند گروه رو برای مقایسه انتخاب کن، یا از منو، دو گروه رو برای مقایسه کامل هفتگی انتخاب کن. گزارش رو می‌تونی پرینت/PDF یا اکسل بگیری."],
  ],
  en: [
    ["Upload the Excel file", "Pick the right sheet if there are several. The app auto-detects the crew, type, shift, weekday and hours columns."],
    ["Build your priorities", "Click criteria (shift, working days, region, standby) in order of importance. Ranking follows that exact order: criterion 1 first, ties broken by criterion 2, and so on."],
    ["Read the results", "The top five matches get gold/silver/bronze cards. The number on the left of each card is that crew's overall match score."],
    ["Compare & export", "Select crews to compare, or use the menu to compare two crews for the full week. Export the report to print/PDF or Excel."],
  ],
};

function HelpPanel({ lang, onClose }) {
  const content = HELP_CONTENT[lang] || HELP_CONTENT.en;
  return (
    <Modal title={t("helpTitle", lang)} onClose={onClose}>
      {content.map(([title, body], i) => (
        <div key={i} style={styles.helpStep}>
          <div style={styles.helpStepTitle}>{i + 1}. {title}</div>
          <div style={styles.helpStepBody}>{body}</div>
        </div>
      ))}
    </Modal>
  );
}

// ---------- About ----------

const ABOUT_CONTENT = {
  fa: {
    desc: "«اولویت‌بندی شیفت» ابزاری برای رانندگان TOK Transit است که جدول شیفت هفتگی (Crew Sheet) را از روی فایل اکسل می‌خواند و بر اساس اولویت‌های شخصی هر راننده (زمان شیفت، تعداد روز کاری، منطقه کاری و شیفت آماده‌باش) بهترین گروه‌ها را رتبه‌بندی می‌کند. تمام پردازش‌ها به‌صورت محلی و بدون ارسال هیچ داده‌ای به اینترنت انجام می‌شود.",
    authorLabel: "طراح و توسعه‌دهنده:",
  },
  en: {
    desc: "\u201cShift Priority\u201d is a tool for TOK Transit drivers that reads the weekly crew shift schedule from an Excel file and ranks crews according to each driver's own priorities (shift time, working days, work region, and standby shifts). All processing happens locally on your device; no data is ever sent over the internet.",
    authorLabel: "Designed and developed by:",
  },
  hi: {
    desc: "\u201cShift Priority\u201d TOK Transit ड्राइवरों के लिए एक टूल है जो साप्ताहिक क्रू शिफ्ट शेड्यूल को एक्सेल फ़ाइल से पढ़ता है और हर ड्राइवर की अपनी प्राथमिकताओं (शिफ्ट समय, कार्य दिवस, क्षेत्र, स्टैंडबाय शिफ्ट) के अनुसार क्रू को रैंक करता है। सारा प्रोसेसिंग स्थानीय रूप से होता है; कोई डेटा इंटरनेट पर नहीं भेजा जाता।",
    authorLabel: "डिज़ाइनर और डेवलपर:",
  },
};

function AboutPanel({ lang, onClose }) {
  const content = ABOUT_CONTENT[lang] || ABOUT_CONTENT.en;
  return (
    <Modal title={t("aboutTitle", lang)} onClose={onClose}>
      <div style={{ textAlign: "center", marginBottom: 16 }}>
        <div style={styles.aboutAppName}>Shift Priority</div>
        <div style={styles.aboutVersion}>v{APP_VERSION} · {new Date().toLocaleDateString("en-US", { month: "long", year: "numeric" })}</div>
      </div>
      <p style={styles.helpStepBody}>{content.desc}</p>
      <div style={styles.aboutAuthorBox}>
        <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 2 }}>{content.authorLabel}</div>
        <div style={styles.aboutAuthorName}>Omid Farhadnia</div>
      </div>
      <div style={{ fontSize: 11.5, color: "var(--muted)", textAlign: "center", marginTop: 10 }}>MIT License · Open Source</div>
    </Modal>
  );
}

// ---------- Help submenu (About / How to Use / Report a problem, grouped) ----------

function HelpMenuPanel({ lang, onPick, onClose }) {
  return (
    <Modal title={t("helpMenuLabel", lang)} onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
        <button style={styles.menuItem} onClick={() => onPick("help")}>
          <HelpCircle size={15} /> {t("helpTitle", lang)}
        </button>
        <button style={styles.menuItem} onClick={() => onPick("about")}>
          <Info size={15} /> {t("aboutTitle", lang)}
        </button>
        <button style={styles.menuItem} onClick={() => onPick("report")}>
          <Mail size={15} /> {t("reportProblem", lang)}
        </button>
      </div>
    </Modal>
  );
}

// ---------- Crew lookup (browse any crew's full weekly schedule) ----------
// Deliberately click-to-open, no typing required — the search box is just
// an accelerator for a long list, not the primary way to pick a crew.

function CrewLookupPanel({ lang, crews, crewNames, onPick, onClose }) {
  const [query, setQuery] = useState("");
  const list = useMemo(() => {
    const q = query.trim().toLowerCase();
    return (crews || [])
      .map((c) => ({ crew: c.crew, name: resolveCrewName(c.crew, crews, crewNames) }))
      .filter((c) => !q || String(c.crew).toLowerCase().includes(q) || (c.name || "").toLowerCase().includes(q))
      .sort((a, b) => Number(a.crew) - Number(b.crew) || String(a.crew).localeCompare(String(b.crew)));
  }, [crews, crewNames, query]);

  return (
    <Modal title={t("crewLookupTitle", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #0EA37E, #3DDC97)" accent="#0EA37E">
      <div style={{ position: "relative", marginBottom: 10 }}>
        <Search size={14} style={{ position: "absolute", insetInlineStart: 10, top: 10, color: "var(--muted)" }} />
        <input
          type="text"
          placeholder={t("crewLookupSearch", lang)}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={{ ...styles.numInputWide, paddingInlineStart: 30 }}
        />
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 360, overflowY: "auto" }}>
        {list.map((c) => (
          <button
            key={c.crew}
            onClick={() => onPick(c.crew)}
            style={{ ...styles.chip, flexDirection: "column", alignItems: "flex-start", gap: 2, minWidth: 88 }}
          >
            <span style={{ fontWeight: 700 }}>{t("crewWord", lang)} {String(c.crew)}</span>
            {c.name && <span style={{ fontSize: 11, color: "var(--muted)" }}>{c.name}</span>}
          </button>
        ))}
        {list.length === 0 && <p style={styles.hint}>{t("crewNotFound", lang)}</p>}
      </div>
    </Modal>
  );
}

// ---------- Leave replacement finder ----------
// Given a day the user wants off (date + shift hours), rank the crews that
// could cover it. Hard rule: the candidate must have NO shift on that
// weekday. Then, because the easiest ask is a straight trade, crews that
// WORK on one of the user's own rest days rank first — the user can offer
// to take that shift back in return. Ties break on rest gaps around the
// shift (a candidate finishing late the night before is a harder ask) and
// on how close the shift is to the candidate's usual hours and region.

const SWAP_MIN_REST_MIN = 8 * 60;

function timeToMinutes(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    const m = Math.round((v % 1) * 1440);
    return ((m % 1440) + 1440) % 1440;
  }
  const s = String(v).trim().toLowerCase();
  const m = s.match(/^(\d{1,2})(?::|\.|h)?(\d{2})?\s*(am|pm|a|p)?$/);
  if (!m) return null;
  let h = Number(m[1]);
  const mm = Number(m[2] || 0);
  const ap = m[3];
  if (ap && ap.startsWith("p") && h < 12) h += 12;
  if (ap && ap.startsWith("a") && h === 12) h = 0;
  if (h > 24 || mm > 59) return null;
  return (h * 60 + mm) % 1440;
}

function minutesToHHMM(m) {
  const x = ((Math.round(m) % 1440) + 1440) % 1440;
  return `${String(Math.floor(x / 60)).padStart(2, "0")}:${String(x % 60).padStart(2, "0")}`;
}

// A day's shift as [start, end] in minutes relative to that day's midnight;
// an end at or before the start means the shift runs past midnight.
function shiftSpan(startMin, endMin) {
  if (startMin === null) return null;
  let end = endMin === null ? startMin : endMin;
  if (end <= startMin) end += 1440;
  return [startMin, end];
}

function daySpan(crew, dayIdx) {
  const d = crew.days.find((x) => x.dayIdx === dayIdx);
  if (!d) return null;
  return shiftSpan(timeToMinutes(d.start), timeToMinutes(d.end));
}

// Rest (minutes) this crew would get before and after working [start,end]
// on `dayIdx`, given its own shifts on the neighbouring days. null = that
// neighbour day is off (plenty of rest).
function restAround(crew, dayIdx, span) {
  const prev = daySpan(crew, (dayIdx + 6) % 7);
  const next = daySpan(crew, (dayIdx + 1) % 7);
  const before = prev ? span[0] - (prev[1] - 1440) : null;
  const after = next ? (next[0] + 1440) - span[1] : null;
  return { before, after, ok: (before === null || before >= SWAP_MIN_REST_MIN) && (after === null || after >= SWAP_MIN_REST_MIN) };
}

function usualStart(crew) {
  const starts = crew.days.map((d) => timeToMinutes(d.start)).filter((x) => x !== null).sort((a, b) => a - b);
  if (!starts.length) return null;
  return starts[Math.floor(starts.length / 2)];
}

// Longest run of consecutive working days in a weekly pattern (wraps the week).
function maxConsecutiveDays(dayIdxs) {
  const set = new Set(dayIdxs);
  if (set.size === 7) return 7;
  let best = 0;
  for (let start = 0; start < 7; start++) {
    if (set.has((start + 6) % 7) || !set.has(start)) continue;
    let n = 0;
    while (set.has((start + n) % 7) && n < 7) n++;
    best = Math.max(best, n);
  }
  return best;
}

// How well one trade day fits MY schedule (0-100): the shift I'd take on my
// rest day vs the hours I'm giving up, same region, enough rest around it,
// and not stretching my week into a long run of days without a break.
function tradeMatchScore(o, { span, targetRegion, myCrew, weekday }) {
  const s = shiftSpan(timeToMinutes(o.start), timeToMinutes(o.end));
  let score = 0;
  if (s) {
    const circ = (a, b) => Math.min(Math.abs(a - b), 1440 - Math.abs(a - b));
    const diffH = (circ(s[0], span[0]) + circ(s[1] % 1440, span[1] % 1440)) / 2 / 60;
    score += Math.max(0, 50 - diffH * 8); // start/end times (max 50)
    const lenDiffH = Math.abs((s[1] - s[0]) - (span[1] - span[0])) / 60;
    score += Math.max(0, 10 - lenDiffH * 3); // similar shift length (max 10)
  }
  if (!targetRegion || o.regionKey === targetRegion) score += 10; // same region (max 10)
  if (o.myRest.ok) score += 20; // >= 8h rest for me around it (max 20)
  if (myCrew) {
    const newDays = myCrew.days.map((d) => d.dayIdx).filter((d) => d !== weekday).concat(o.dayIdx);
    const streak = maxConsecutiveDays(newDays);
    score += streak <= 5 ? 10 : Math.max(0, 10 - (streak - 5) * 5); // keeps a break in my week (max 10)
  }
  return Math.round(Math.min(100, score));
}

function findLeaveReplacements(crews, { weekday, startMin, endMin, myCrew, targetRegion }) {
  const span = shiftSpan(startMin, endMin);
  if (!span) return [];
  const myOffDays = myCrew ? [0, 1, 2, 3, 4, 5, 6].filter((d) => !myCrew.days.some((x) => x.dayIdx === d)) : [];
  const out = [];
  for (const c of crews || []) {
    if (myCrew && String(c.crew) === String(myCrew.crew)) continue;
    if (c.days.some((d) => d.dayIdx === weekday)) continue; // works that day -> can't cover
    if (!c.days.length) continue; // empty row / not an active crew

    const candRest = restAround(c, weekday, span);

    // Their shifts on my rest days = what I can offer to take in return.
    const swapOptions = c.days
      .filter((d) => myOffDays.includes(d.dayIdx))
      .map((d) => {
        const s = shiftSpan(timeToMinutes(d.start), timeToMinutes(d.end));
        const myRest = s ? restAround(myCrew, d.dayIdx, s) : { ok: true, before: null, after: null };
        // Days after the requested day first (the natural "I'll pay you back"), wrapping the week.
        const dist = (d.dayIdx - weekday + 7) % 7;
        return { ...d, myRest, dist };
      })
      .map((o) => ({ ...o, match: tradeMatchScore(o, { span, targetRegion, myCrew, weekday }) }))
      // Best fit with my own schedule first; ties -> the sooner "pay back" day.
      .sort((a, b) => (b.match - a.match) || (b.myRest.ok - a.myRest.ok) || (a.dist - b.dist));

    const us = usualStart(c);
    const diffH = us === null ? 6 : Math.min(Math.abs(us - startMin), 1440 - Math.abs(us - startMin)) / 60;
    const timeFit = Math.max(0, 100 - diffH * 15);
    const regionShare = targetRegion ? c.days.filter((d) => d.regionKey === targetRegion).length / c.days.length : null;
    const fit = Math.round(regionShare === null ? timeFit : timeFit * 0.7 + regionShare * 100 * 0.3);

    out.push({
      crew: c,
      swapOptions,
      canSwap: swapOptions.length > 0,
      swapRestOk: swapOptions.some((o) => o.myRest.ok),
      candRest,
      usualStart: us,
      fit,
    });
  }
  out.sort((a, b) =>
    (b.canSwap - a.canSwap) ||
    (b.candRest.ok - a.candRest.ok) ||
    (b.swapRestOk - a.swapRestOk) ||
    (b.fit - a.fit) ||
    (a.crew.totalHours - b.crew.totalHours) ||
    (Number(a.crew.crew) - Number(b.crew.crew))
  );
  return out;
}

function localDateStr(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function SwapFinderPanel({ lang, crews, crewNames, profile, onViewCrew, onClose }) {
  const [myCrewNum, setMyCrewNum] = useState(profile?.crewNumber ? String(profile.crewNumber) : "");
  const [date, setDate] = useState(localDateStr());
  const [startTime, setStartTime] = useState("");
  const [endTime, setEndTime] = useState("");
  const [results, setResults] = useState(null);
  const [showAll, setShowAll] = useState(false);
  const [error, setError] = useState("");

  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const weekday = date ? new Date(date + "T00:00:00").getDay() : null;
  const myCrew = useMemo(() => (crews || []).find((c) => String(c.crew) === myCrewNum.trim()) || null, [crews, myCrewNum]);
  const myShiftThatDay = myCrew && weekday !== null ? myCrew.days.find((d) => d.dayIdx === weekday) : null;

  // Pre-fill the hours from my own shift on that weekday; still editable.
  useEffect(() => {
    if (myShiftThatDay) {
      setStartTime(formatExcelTime(myShiftThatDay.start));
      setEndTime(formatExcelTime(myShiftThatDay.end));
    }
    setResults(null);
  }, [myCrew, weekday]); // eslint-disable-line react-hooks/exhaustive-deps

  const doFind = () => {
    const s = timeToMinutes(startTime);
    const e = timeToMinutes(endTime);
    if (weekday === null || s === null || e === null) { setError(t("swapNeedInputs", lang)); setResults(null); return; }
    setError("");
    setShowAll(false);
    setResults(findLeaveReplacements(crews, { weekday, startMin: s, endMin: e, myCrew, targetRegion: myShiftThatDay?.regionKey || null }));
  };

  const fmtRest = (m) => formatDuration(m / 60, lang);
  const shown = results ? results.slice(0, showAll ? 15 : 5) : [];

  return (
    <Modal title={t("hubSwapFinder", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #D9480F, #F59F00)" accent="#D9480F">
      <p style={styles.hint}>{t("swapIntro", lang)}</p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <label style={{ fontSize: 12, fontWeight: 700 }}>{t("swapMyCrew", lang)}
          <input type="number" value={myCrewNum} onChange={(e) => setMyCrewNum(e.target.value)} placeholder={t("crewWord", lang)} style={{ ...styles.numInputWide, width: "100%", minWidth: 0, marginTop: 4, boxSizing: "border-box" }} />
        </label>
        <label style={{ fontSize: 12, fontWeight: 700 }}>{t("swapDate", lang)}{weekday !== null ? ` · ${weekdayNames[weekday]}` : ""}
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...styles.numInputWide, width: "100%", minWidth: 0, marginTop: 4, boxSizing: "border-box" }} />
        </label>
        <label style={{ fontSize: 12, fontWeight: 700 }}>{t("swapStart", lang)}
          <input type="time" value={startTime} onChange={(e) => setStartTime(e.target.value)} style={{ ...styles.numInputWide, width: "100%", minWidth: 0, marginTop: 4, boxSizing: "border-box" }} />
        </label>
        <label style={{ fontSize: 12, fontWeight: 700 }}>{t("swapEnd", lang)}
          <input type="time" value={endTime} onChange={(e) => setEndTime(e.target.value)} style={{ ...styles.numInputWide, width: "100%", minWidth: 0, marginTop: 4, boxSizing: "border-box" }} />
        </label>
      </div>
      {myCrewNum.trim() && !myCrew && <div style={styles.errorBox}><AlertCircle size={15} /><span>{t("crewNumberNotInFile", lang)}</span></div>}
      {!myCrewNum.trim() && <p style={{ ...styles.hint, marginTop: 8 }}>{t("swapNoCrewHint", lang)}</p>}
      {myCrew && weekday !== null && !myShiftThatDay && <p style={{ ...styles.hint, marginTop: 8, color: "#B3432A" }}>{t("swapAlreadyOff", lang)}</p>}
      {myCrew && (
        <p style={{ ...styles.hint, marginTop: 8 }}>
          {t("swapMyRestDays", lang)}: <b>{[0, 1, 2, 3, 4, 5, 6].filter((d) => !myCrew.days.some((x) => x.dayIdx === d)).map((d) => weekdayNames[d]).join(DAY_LIST_SEP[lang] || ", ") || "—"}</b>
        </p>
      )}
      <button onClick={doFind} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)", marginTop: 8 }}>
        <Search size={14} /> {t("swapFindBtn", lang)}
      </button>
      {error && <div style={styles.errorBox}><AlertCircle size={15} /><span>{error}</span></div>}

      {results && (
        <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ fontSize: 12.5, fontWeight: 800 }}>
            {t("swapResultsTitle", lang)} — {weekdayNames[weekday]} <bdi dir="ltr">{startTime}–{endTime}</bdi> · {results.length} {t("swapCandidatesWord", lang)}
          </div>
          {results.length === 0 && <p style={styles.hint}>{t("swapNoResults", lang)}</p>}
          {shown.map((r, i) => {
            const name = resolveCrewName(r.crew.crew, crews, crewNames);
            const best = r.swapOptions[0];
            return (
              <div key={r.crew.crew} style={{ border: "1px solid var(--border)", borderInlineStart: `4px solid ${r.canSwap ? "#0EA37E" : "#C9A227"}`, borderRadius: 10, padding: "9px 11px", background: "var(--card)" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontWeight: 900, fontSize: 15, color: "var(--accent)" }}>{i + 1}</span>
                  <span style={{ fontWeight: 800, fontSize: 13.5 }}>{t("crewWord", lang)} {String(r.crew.crew)}{name ? ` · ${name}` : ""}</span>
                  <span style={{ marginInlineStart: "auto", fontSize: 11, color: "var(--muted)" }}>{t("swapFit", lang)} {r.fit}%</span>
                </div>
                <div style={{ fontSize: 12, marginTop: 5, display: "flex", flexDirection: "column", gap: 3 }}>
                  <span style={{ color: "#0EA37E", fontWeight: 700 }}>✓ {t("swapOffThatDay", lang)} {weekdayNames[weekday]}</span>
                  {r.usualStart !== null && <span style={{ color: "var(--muted)" }}>{t("swapUsualStart", lang)} {minutesToHHMM(r.usualStart)} · {r.crew.type} {r.crew.shiftRaw}</span>}
                  {r.canSwap ? (
                    <span style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                      <span style={{ fontWeight: 700 }}>⇄ {t("swapTradeFor", lang)}:</span>
                      {r.swapOptions.map((o, k) => (
                        <span key={o.dayIdx} style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", background: k === 0 ? "rgba(14,163,126,0.10)" : "var(--bg)", border: `1px solid ${k === 0 ? "#0EA37E" : "var(--border)"}`, borderRadius: 7, padding: "4px 8px" }}>
                          <b style={{ color: k === 0 ? "#0EA37E" : "var(--muted)" }}>{k + 1}</b>
                          <b>{weekdayNames[o.dayIdx]}</b>
                          <bdi dir="ltr">{formatExcelTime(o.start)}–{formatExcelTime(o.end)}</bdi>
                          {o.myRest.ok ? null : <span style={{ color: "#B3432A" }}>⚠</span>}
                          <span style={{ marginInlineStart: "auto", fontSize: 11, color: "var(--muted)" }}>{t("swapMatch", lang)} {o.match}%</span>
                        </span>
                      ))}
                    </span>
                  ) : (
                    <span style={{ color: "#A07C00" }}>{t("swapNoTrade", lang)}</span>
                  )}
                  {best && !best.myRest.ok && <span style={{ color: "#B3432A" }}>⚠ {t("swapMyRestWarn", lang)}</span>}
                  {!r.candRest.ok && (
                    <span style={{ color: "#B3432A" }}>
                      ⚠ {t("swapTheirRestWarn", lang)}
                      {r.candRest.before !== null && r.candRest.before < SWAP_MIN_REST_MIN ? ` (${t("swapBefore", lang)} ${fmtRest(r.candRest.before)})` : ""}
                      {r.candRest.after !== null && r.candRest.after < SWAP_MIN_REST_MIN ? ` (${t("swapAfter", lang)} ${fmtRest(r.candRest.after)})` : ""}
                    </span>
                  )}
                </div>
                <button onClick={() => onViewCrew(r.crew)} style={{ ...styles.smallActionBtn, marginTop: 7, padding: "5px 9px", fontSize: 11.5 }}>
                  <CalendarOff size={13} /> {t("swapViewSchedule", lang)}
                </button>
              </div>
            );
          })}
          {results.length > 5 && !showAll && (
            <button onClick={() => setShowAll(true)} style={styles.smallActionBtn}>{t("swapShowMore", lang)}</button>
          )}
        </div>
      )}
    </Modal>
  );
}

// ---------- Admin ----------
// See the ADMIN_USERNAME/ADMIN_PASSWORD note near the top of this file: this
// login only hides the panel below from casual users, it is not real
// security — there is no server, so nothing here can be kept truly secret.

function AdminLoginModal({ lang, onSuccess, onClose }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(false);

  const submit = () => {
    if (username.trim() === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
      saveAdminSession(true);
      onSuccess();
    } else {
      setError(true);
    }
  };

  return (
    <Modal title={t("adminLoginTitle", lang)} onClose={onClose}>
      <div style={styles.smallLabel}>{t("adminUsernameLabel", lang)}</div>
      <input
        type="text"
        value={username}
        onChange={(e) => { setUsername(e.target.value); setError(false); }}
        style={styles.numInputWide}
        autoComplete="off"
      />
      <div style={{ ...styles.smallLabel, marginTop: 10 }}>{t("adminPasswordLabel", lang)}</div>
      <input
        type="password"
        value={password}
        onChange={(e) => { setPassword(e.target.value); setError(false); }}
        onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
        style={styles.numInputWide}
        autoComplete="off"
      />
      {error && (
        <div style={styles.errorBox}><AlertCircle size={15} /><span>{t("adminLoginError", lang)}</span></div>
      )}
      <button onClick={submit} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)", marginTop: 10 }}>
        <Lock size={14} /> {t("adminLoginBtn", lang)}
      </button>
    </Modal>
  );
}

function AdminPanel({ lang, crews, crewNames, setCrewNames, dailyLogAccess, setDailyLogAccess, onLogout, onClose }) {
  const [drafts, setDrafts] = useState({});
  const [newCrew, setNewCrew] = useState("");
  const [newName, setNewName] = useState("");
  const [dupWarning, setDupWarning] = useState({}); // { [crewNum]: theOtherCrewNumItClashesWith }
  const [newDup, setNewDup] = useState(null);
  const [newDailyLogCrew, setNewDailyLogCrew] = useState("");

  const addDailyLogAccess = () => {
    const num = newDailyLogCrew.trim();
    if (!num) return;
    if (!dailyLogAccess.includes(num)) {
      const next = [...dailyLogAccess, num];
      setDailyLogAccess(next);
      saveDailyLogAccess(next);
    }
    setNewDailyLogCrew("");
  };
  const removeDailyLogAccess = (num) => {
    const next = dailyLogAccess.filter((n) => n !== num);
    setDailyLogAccess(next);
    saveDailyLogAccess(next);
  };

  // Union of every crew number in the currently loaded file, every crew
  // number in the built-in default list, plus every crew number that
  // already has a manually-saved name (so manual entries survive even once
  // the file that had them is gone, or a different file is loaded).
  const allNumbers = useMemo(() => {
    const set = new Set();
    (crews || []).forEach((c) => set.add(String(c.crew)));
    Object.keys(CREW_NAME_DEFAULTS).forEach((k) => set.add(k));
    Object.keys(crewNames || {}).forEach((k) => set.add(k));
    return Array.from(set).sort((a, b) => (Number(a) - Number(b)) || a.localeCompare(b));
  }, [crews, crewNames]);

  const autoNameFor = (num) => (crews || []).find((c) => String(c.crew) === num)?.driverName || "";
  const defaultNameFor = (num) => CREW_NAME_DEFAULTS[num] || "";
  // What a crew currently resolves to, same priority as resolveCrewName()
  // (manual override > auto-read > built-in default) — used to check a new
  // name against every OTHER crew before it's allowed to save.
  const resolvedNameFor = (num) => {
    if (crewNames && Object.prototype.hasOwnProperty.call(crewNames, num)) return crewNames[num];
    return autoNameFor(num) || defaultNameFor(num);
  };
  // A blank name never counts as a clash (every open crew is blank on
  // purpose); a real name can't be saved under two different crew numbers.
  const findDuplicate = (num, name) => {
    const clean = name.trim().toLowerCase();
    if (!clean) return null;
    const hit = allNumbers.find((n) => n !== num && resolvedNameFor(n).trim().toLowerCase() === clean);
    return hit || null;
  };

  const commit = (num, value) => {
    const clean = value.trim();
    const dup = findDuplicate(num, clean);
    if (dup) {
      setDupWarning((prev) => ({ ...prev, [num]: dup }));
      return false;
    }
    setDupWarning((prev) => { if (!(num in prev)) return prev; const n = { ...prev }; delete n[num]; return n; });
    const next = { ...crewNames, [String(num)]: clean };
    setCrewNames(next);
    saveCrewNames(next);
    return true;
  };
  // Explicitly blanks a crew (an open/unassigned run) rather than relying on
  // clearing the text field and hoping blur saves it.
  const clearRow = (num) => {
    commit(num, "");
    setDrafts((prev) => { const n = { ...prev }; delete n[num]; return n; });
  };
  const removeOverride = (num) => {
    const next = { ...crewNames };
    delete next[String(num)];
    setCrewNames(next);
    saveCrewNames(next);
    setDrafts((prev) => { const n = { ...prev }; delete n[num]; return n; });
    setDupWarning((prev) => { if (!(num in prev)) return prev; const n = { ...prev }; delete n[num]; return n; });
  };
  const addNew = () => {
    const num = newCrew.trim();
    if (!num) return;
    const dup = findDuplicate(num, newName.trim());
    if (dup) { setNewDup(dup); return; }
    setNewDup(null);
    commit(num, newName.trim());
    setNewCrew("");
    setNewName("");
  };

  return (
    <Modal title={t("adminPanelTitle", lang)} onClose={onClose}>
      <p style={styles.hint}>{t("adminPanelHint", lang)}</p>

      <div style={{ display: "flex", gap: 6, marginBottom: 4 }}>
        <input
          type="number"
          placeholder={t("crewNumberLabel", lang)}
          value={newCrew}
          onChange={(e) => { setNewCrew(e.target.value); setNewDup(null); }}
          style={{ ...styles.numInputWide, minWidth: 90, width: 90, flex: "none" }}
        />
        <input
          type="text"
          placeholder={t("driverNameLabel", lang)}
          value={newName}
          onChange={(e) => { setNewName(e.target.value); setNewDup(null); }}
          style={styles.numInputWide}
        />
        <button onClick={addNew} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" }}>
          <Plus size={14} />
        </button>
      </div>
      {newDup && (
        <p style={{ fontSize: 11, color: "#B3432A", margin: "0 0 8px" }}>
          {t("adminDupWarning", lang)} {t("crewWord", lang)} {newDup}
        </p>
      )}

      <div style={{ maxHeight: 340, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
        {allNumbers.length === 0 && <p style={styles.hint}>{t("adminNoCrews", lang)}</p>}
        {allNumbers.map((num) => {
          const hasOverride = Object.prototype.hasOwnProperty.call(crewNames || {}, num);
          const auto = autoNameFor(num);
          const def = defaultNameFor(num);
          const value = drafts[num] !== undefined ? drafts[num] : (hasOverride ? crewNames[num] : (auto || def));
          const badge = hasOverride
            ? (crewNames[num] === "" ? t("adminBlankBadge", lang) : t("adminManualBadge", lang))
            : auto ? t("adminAutoBadge", lang) : def ? t("adminDefaultBadge", lang) : t("adminBlankBadge", lang);
          const dup = dupWarning[num];
          return (
            <div key={num} style={{ display: "flex", flexDirection: "column", gap: 3 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid var(--border)", borderRadius: 8, padding: "6px 8px" }}>
                <span style={{ fontWeight: 700, fontSize: 12.5, minWidth: 60 }}>{t("crewWord", lang)} {num}</span>
                <input
                  type="text"
                  value={value}
                  onChange={(e) => setDrafts((prev) => ({ ...prev, [num]: e.target.value }))}
                  onBlur={() => { if (drafts[num] !== undefined) commit(num, drafts[num]); }}
                  style={{ ...styles.numInputWide, padding: "5px 8px", fontSize: 12.5 }}
                />
                <span style={{ fontSize: 10.5, color: "var(--muted)", whiteSpace: "nowrap" }}>
                  {badge}
                </span>
                <button onClick={() => clearRow(num)} title={t("adminClearTooltip", lang)} style={{ ...styles.smallActionBtn, padding: "5px 7px" }}>
                  <X size={13} />
                </button>
                {hasOverride && (
                  <button onClick={() => removeOverride(num)} style={{ ...styles.smallActionBtn, padding: "5px 7px", color: "#B3432A" }}>
                    <Trash2 size={13} />
                  </button>
                )}
              </div>
              {dup && (
                <p style={{ fontSize: 11, color: "#B3432A", margin: "0 4px" }}>
                  {t("adminDupWarning", lang)} {t("crewWord", lang)} {dup}
                </p>
              )}
            </div>
          );
        })}
      </div>

      <div style={{ borderTop: "1px solid var(--border)", marginTop: 16, paddingTop: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 4 }}>{t("adminDailyLogAccessTitle", lang)}</div>
        <p style={styles.hint}>{t("adminDailyLogAccessHint", lang)}</p>
        <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
          <input
            type="number"
            placeholder={t("crewNumberLabel", lang)}
            value={newDailyLogCrew}
            onChange={(e) => setNewDailyLogCrew(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") addDailyLogAccess(); }}
            style={{ ...styles.numInputWide, minWidth: 90, width: 90, flex: "none" }}
          />
          <button onClick={addDailyLogAccess} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" }}>
            <Plus size={14} />
          </button>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {dailyLogAccess.length === 0 && <p style={styles.hint}>{t("adminDailyLogAccessEmpty", lang)}</p>}
          {dailyLogAccess.map((num) => (
            <span key={num} style={{ display: "flex", alignItems: "center", gap: 4, border: "1px solid var(--border)", borderRadius: 20, padding: "3px 4px 3px 10px", fontSize: 12 }}>
              {t("crewWord", lang)} {num}
              <button onClick={() => removeDailyLogAccess(num)} style={{ ...styles.smallActionBtn, padding: "3px 5px", border: "none" }}>
                <X size={11} />
              </button>
            </span>
          ))}
        </div>
      </div>

      <button onClick={onLogout} style={{ ...styles.smallActionBtn, color: "#B3432A", marginTop: 12 }}>
        <LogOut size={14} /> {t("adminLogoutBtn", lang)}
      </button>
    </Modal>
  );
}


// ---------- Settings (theme) ----------

function SettingsPanel({ lang, themeStyle, setThemeStyle, themeMode, setThemeMode, onClose }) {
  const styleOptions = [
    { v: "win11", l: t("themeWin11", lang) },
    { v: "macos", l: t("themeMac", lang) },
    { v: "universal", l: t("themeUniversal", lang) },
  ];
  return (
    <Modal title={t("settingsTitle", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #E08A1E, #F6B93B)" accent="#E08A1E">
      <div style={styles.prefGroup}>
        <div style={styles.prefTitle}>{t("themeMode", lang)}</div>
        <div style={styles.chipRow}>
          <button onClick={() => setThemeMode("light")} style={{ ...styles.chip, ...(themeMode === "light" ? styles.chipActive : {}) }}>
            <Sun size={14} /> {t("light", lang)}
          </button>
          <button onClick={() => setThemeMode("dark")} style={{ ...styles.chip, ...(themeMode === "dark" ? styles.chipActive : {}) }}>
            <Moon size={14} /> {t("dark", lang)}
          </button>
        </div>
      </div>
      <div style={styles.prefGroup}>
        <div style={styles.prefTitle}>{t("themeStyle", lang)}</div>
        <div style={styles.chipRow}>
          {styleOptions.map((o) => (
            <button key={o.v} onClick={() => setThemeStyle(o.v)} style={{ ...styles.chip, ...(themeStyle === o.v ? styles.chipActive : {}) }}>
              {themeStyle === o.v && <Check size={13} />} {o.l}
            </button>
          ))}
        </div>
      </div>
    </Modal>
  );
}

// Excel stores time-of-day cells as a fraction of a day (0–1); convert to HH:MM.
function formatExcelTime(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    let totalMinutes = Math.round(v * 24 * 60);
    totalMinutes = ((totalMinutes % 1440) + 1440) % 1440;
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
    const mm = String(totalMinutes % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  }
  return String(v).trim();
}

function formatDuration(hoursDecimal, lang) {
  const totalMinutes = Math.round(hoursDecimal * 60);
  const hh = Math.floor(totalMinutes / 60);
  const mm = totalMinutes % 60;
  if (lang === "fa") {
    let s = `${toFaDigits(hh)} ساعت`;
    if (mm > 0) s += ` و ${toFaDigits(mm)} دقیقه`;
    return s;
  }
  return mm > 0 ? `${hh}h ${mm}m` : `${hh}h`;
}

const PERSON_PALETTE = [
  { bg: "#FCE4EC", border: "#F48FB1", text: "#AD1457" },
  { bg: "#E3F2FD", border: "#90CAF9", text: "#1565C0" },
  { bg: "#E8F5E9", border: "#A5D6A7", text: "#2E7D32" },
  { bg: "#EDE7F6", border: "#B39DDB", text: "#5E35B1" },
  { bg: "#FFF3E0", border: "#FFCC80", text: "#E65100" },
];

function buildCompareHtml(matched, lang, timestamp) {
  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const dir = lang === "fa" ? "rtl" : "ltr";
  const off = t("off", lang);
  const dayCell = (crew, i) => crew.days.find((d) => d.dayIdx === i);

  const headerCells = matched.map((m, i) => `<th style="background:${PERSON_PALETTE[i % 5].bg};color:${PERSON_PALETTE[i % 5].text};border:1px solid #ddd;padding:10px;">${t("crewWord", lang)} ${m.crew}</th>`).join("");

  const rows = weekdayNames.map((wd, di) => {
    const cells = matched.map((m, i) => {
      const d = dayCell(m, di);
      const pal = PERSON_PALETTE[i % 5];
      if (!d) {
        return `<td style="border:1px solid #ddd;padding:6px;"><div style="border:1.5px dashed #e0a0a0;background:#fbeceb;color:#b3432a;border-radius:8px;padding:14px;text-align:center;font-weight:700;">${off} ✕</div></td>`;
      }
      const regionDot = REGION_COLORS[d.regionKey] || "#9AA0A6";
      return `<td style="border:1px solid #ddd;padding:6px;">
        <div style="background:${pal.bg};border:1px solid ${pal.border};border-radius:8px;padding:10px;">
          <div style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:6px;flex-wrap:wrap;">
            <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${regionDot};"></span>
            <span style="background:#fff;border-radius:8px;padding:2px 6px;">${regionLabel(d.regionKey, lang)}</span>
            <span style="background:#fff;border-radius:8px;padding:2px 6px;">${d.code || "-"}</span>
          </div>
          <div style="font-weight:700;font-size:14px;">${formatExcelTime(d.start)}–${formatExcelTime(d.end)}</div>
          <div style="font-size:11.5px;color:#555;">${formatDuration(d.hours, lang)}</div>
        </div>
      </td>`;
    }).join("");
    return `<tr><td style="border:1px solid #ddd;padding:8px;font-weight:700;background:#faf8f3;white-space:nowrap;">${wd}</td>${cells}</tr>`;
  }).join("");

  const totalsRow = `<tr><td style="border:1px solid #ddd;padding:8px;font-weight:800;background:#faf8f3;">${t("hours", lang)}</td>${matched.map((m) => `<td style="border:1px solid #ddd;padding:8px;text-align:center;font-weight:800;">${m.totalHours}</td>`).join("")}</tr>`;
  const metaRows = [
    [t("type", lang), matched.map((m) => m.type)],
    [t("shift", lang), matched.map((m) => m.shiftRaw)],
    [t("daysColumn", lang), matched.map((m) => m.workedCount)],
  ].map(([label, vals]) => `<tr><td style="border:1px solid #ddd;padding:8px;font-weight:600;background:#faf8f3;">${label}</td>${vals.map((v) => `<td style="border:1px solid #ddd;padding:8px;text-align:center;">${v}</td>`).join("")}</tr>`).join("");

  return `<!doctype html><html lang="${lang}" dir="${dir}"><head><meta charset="UTF-8" />
  <title>${t("compare2Title", lang)}</title>
  <style>
    body { font-family: Tahoma, 'Vazirmatn', sans-serif; margin: 24px; color:#20242B; }
    table { border-collapse: collapse; width: 100%; }
    .hdr { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px; }
    .hdr h1 { font-size:20px; margin:0 0 4px; }
    .ts { text-align:left; font-size:12px; color:#666; }
    .toolbar { position: sticky; top: 0; background: #fff; padding: 10px 0 16px; display:flex; gap:8px; justify-content:flex-end; border-bottom: 1px solid #eee; margin-bottom: 16px; z-index: 10; }
    .toolbar button { font-size:14px; padding:10px 16px; border-radius:8px; border:1px solid #ccc; background:#fff; cursor:pointer; }
    .toolbar .close-btn { background:#B3432A; color:#fff; border-color:#B3432A; font-weight:700; }
    .toolbar .print-btn { background:#2F5D62; color:#fff; border-color:#2F5D62; font-weight:700; }
    @media print { .toolbar { display: none !important; } }
  </style></head>
  <body>
    <div class="toolbar">
      <button class="print-btn" onclick="window.print()">🖨️ ${t("printBtn", lang)}</button>
      <button class="close-btn" onclick="window.close()">✕ ${t("close", lang)}</button>
    </div>
    <div class="hdr">
      <div><h1>${t("compare2Title", lang)}</h1></div>
      <div class="ts">${timestamp}</div>
    </div>
    <table>
      <thead><tr><th style="border:1px solid #ddd;padding:10px;"></th>${headerCells}</tr></thead>
      <tbody>${rows}${totalsRow}${metaRows}</tbody>
    </table>
  </body></html>`;
}

function buildResultsHtml(results, priorityList, lang, timestamp) {
  const dir = lang === "fa" ? "rtl" : "ltr";
  const critLabels = priorityList.map((id, i) => `${i + 1}. ${results[0]?.fingerprint?.[i]?.label?.[lang] || CRITERIA_CATALOG.find((c) => c.id === id)?.label[lang]}`);
  const headers = [t("rank", lang), t("crewWord", lang), t("type", lang), t("shift", lang), t("daysColumn", lang), t("hours", lang), t("region", lang), ...critLabels];
  const headCells = headers.map((h) => `<th style="border:1px solid #ddd;padding:8px;background:#EAF1EF;">${h}</th>`).join("");
  const medalBg = { 0: "#FBF1DA", 1: "#F2F2F2", 2: "#F6E9DC" };
  const rows = results.map((r, idx) => {
    const regionText = Object.entries(r.regionSummary).map(([k, arr]) => `${arr.length} ${regionLabel(k, lang)} (${arr.join(", ")} ${t("hoursWord", lang)})`).join("; ");
    const cells = [idx + 1, r.crew, r.type, r.shiftRaw, r.workedCount, r.totalHours, regionText, ...r.fingerprint.map((f) => `${f.score}%`)];
    const bg = medalBg[idx] || (idx % 2 ? "#FAFAF8" : "#FFFFFF");
    return `<tr>${cells.map((c) => `<td style="border:1px solid #ddd;padding:7px;background:${bg};">${c}</td>`).join("")}</tr>`;
  }).join("");

  return `<!doctype html><html lang="${lang}" dir="${dir}"><head><meta charset="UTF-8" />
  <title>${t("resultsTitle", lang)}</title>
  <style>
    body { font-family: Tahoma, 'Vazirmatn', sans-serif; margin: 24px; color:#20242B; }
    table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
    .hdr { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px; }
    .hdr h1 { font-size:20px; margin:0 0 4px; }
    .ts { text-align:left; font-size:12px; color:#666; }
    .toolbar { position: sticky; top: 0; background: #fff; padding: 10px 0 16px; display:flex; gap:8px; justify-content:flex-end; border-bottom: 1px solid #eee; margin-bottom: 16px; z-index: 10; }
    .toolbar button { font-size:14px; padding:10px 16px; border-radius:8px; border:1px solid #ccc; background:#fff; cursor:pointer; }
    .toolbar .close-btn { background:#B3432A; color:#fff; border-color:#B3432A; font-weight:700; }
    .toolbar .print-btn { background:#2F5D62; color:#fff; border-color:#2F5D62; font-weight:700; }
    @media print { .toolbar { display: none !important; } }
  </style></head>
  <body>
    <div class="toolbar">
      <button class="print-btn" onclick="window.print()">🖨️ ${t("printBtn", lang)}</button>
      <button class="close-btn" onclick="window.close()">✕ ${t("close", lang)}</button>
    </div>
    <div class="hdr">
      <div><h1>${t("resultsTitle", lang)}</h1></div>
      <div class="ts">${timestamp}</div>
    </div>
    <table><thead><tr>${headCells}</tr></thead><tbody>${rows}</tbody></table>
  </body></html>`;
}

// Opens `html` as a real new-tab navigation (a blob: URL) instead of the
// classic window.open("") + document.write() combo -- mobile Safari (in
// particular the installed home-screen PWA) frequently blocks or silently
// fails that combo, so nothing visibly happens when the person taps Print.
// A direct window.open(url, "_blank") is a real link-like navigation the
// browser is far less likely to block. Returns false if the popup was
// blocked so the caller can tell the person what to do.
function openPrintableReport(html, lang) {
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, "_blank");
  if (!win) {
    URL.revokeObjectURL(url);
    alert(t("printPopupBlocked", lang));
    return false;
  }
  // Give the new tab time to actually load the blob before releasing it --
  // revoking too early can blank the page on some mobile browsers.
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  return true;
}

function printResultsTable(results, priorityList, lang) {
  const now = new Date();
  const timestamp = `${formatJalaliDate(now)} — ${new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" }).format(now)} — ${now.toLocaleTimeString("en-GB")}`;
  openPrintableReport(buildResultsHtml(results, priorityList, lang, timestamp), lang);
}

function buildDailyLogHtml(entries, lang, timestamp) {
  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const dir = lang === "fa" ? "rtl" : "ltr";
  const headers = [
    t("dailyLogColDay", lang), t("dailyLogDateLabel", lang), t("dailyLogStartTimeLabel", lang),
    t("dailyLogEndTimeLabel", lang), t("dailyLogStartYardLabel", lang), t("dailyLogEndYardLabel", lang),
    t("dailyLogDescriptionLabel", lang), t("dailyLogTotalHoursLabel", lang),
  ];
  const headCells = headers.map((h) => `<th style="border:1px solid #ddd;padding:8px;background:#EAF1EF;">${h}</th>`).join("");
  const rows = entries.map((e, idx) => {
    const wd = weekdayNames[new Date(e.date + "T00:00:00").getDay()];
    const bg = idx % 2 ? "#FAFAF8" : "#FFFFFF";
    const cells = [wd, e.date, e.startTime, e.endTime, e.startYard || "-", e.endYard || "-", e.description || "-", formatDuration(e.totalHours, lang)];
    return `<tr>${cells.map((c) => `<td style="border:1px solid #ddd;padding:7px;background:${bg};">${c}</td>`).join("")}</tr>`;
  }).join("");
  const totalHours = entries.reduce((sum, e) => sum + (e.totalHours || 0), 0);
  const totalRow = `<tr><td colspan="7" style="border:1px solid #ddd;padding:8px;font-weight:800;background:#faf8f3;text-align:${dir === "rtl" ? "left" : "right"};">${t("dailyLogTotalRowLabel", lang)}</td><td style="border:1px solid #ddd;padding:8px;font-weight:800;background:#faf8f3;">${formatDuration(totalHours, lang)}</td></tr>`;
  const body = entries.length
    ? `<table><thead><tr>${headCells}</tr></thead><tbody>${rows}${totalRow}</tbody></table>`
    : `<p>${t("dailyLogNoEntriesInRange", lang)}</p>`;

  return `<!doctype html><html lang="${lang}" dir="${dir}"><head><meta charset="UTF-8" />
  <title>${t("dailyLogReportTitle", lang)}</title>
  <style>
    body { font-family: Tahoma, 'Vazirmatn', sans-serif; margin: 24px; color:#20242B; }
    table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
    .hdr { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px; }
    .hdr h1 { font-size:20px; margin:0 0 4px; }
    .ts { text-align:left; font-size:12px; color:#666; }
    .toolbar { position: sticky; top: 0; background: #fff; padding: 10px 0 16px; display:flex; gap:8px; justify-content:flex-end; border-bottom: 1px solid #eee; margin-bottom: 16px; z-index: 10; }
    .toolbar button { font-size:14px; padding:10px 16px; border-radius:8px; border:1px solid #ccc; background:#fff; cursor:pointer; }
    .toolbar .close-btn { background:#B3432A; color:#fff; border-color:#B3432A; font-weight:700; }
    .toolbar .print-btn { background:#2F5D62; color:#fff; border-color:#2F5D62; font-weight:700; }
    @media print { .toolbar { display: none !important; } }
  </style></head>
  <body>
    <div class="toolbar">
      <button class="print-btn" onclick="window.print()">🖨️ ${t("printBtn", lang)}</button>
      <button class="close-btn" onclick="window.close()">✕ ${t("close", lang)}</button>
    </div>
    <div class="hdr">
      <div><h1>${t("dailyLogReportTitle", lang)}</h1></div>
      <div class="ts">${timestamp}</div>
    </div>
    ${body}
  </body></html>`;
}

function printDailyLogTable(entries, lang) {
  const now = new Date();
  const timestamp = `${formatJalaliDate(now)} — ${new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" }).format(now)} — ${now.toLocaleTimeString("en-GB")}`;
  openPrintableReport(buildDailyLogHtml(entries, lang, timestamp), lang);
}

function printCompareTable(matched, lang) {
  const now = new Date();
  const timestamp = `${formatJalaliDate(now)} — ${new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" }).format(now)} — ${now.toLocaleTimeString("en-GB")}`;
  openPrintableReport(buildCompareHtml(matched, lang, timestamp), lang);
}

function hexNoHash(h) { return h.replace("#", "").toUpperCase(); }

function applyBorderCenterToRange(ws) {
  if (!ws["!ref"]) return;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const thin = { style: "thin", color: { rgb: "D9D9D9" } };
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const ref = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      const prevAlign = (ws[ref].s && ws[ref].s.alignment) || {};
      ws[ref].s = {
        ...(ws[ref].s || {}),
        border: { top: thin, bottom: thin, left: thin, right: thin },
        alignment: { horizontal: "center", vertical: "center", wrapText: true, ...prevAlign },
      };
    }
  }
}

function compareToExcel(matched, lang) {
  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const offLabel = t("off", lang);
  const dayColLabel = lang === "fa" ? "روز" : lang === "hi" ? "दिन" : "Day";
  const regionColLabel = lang === "fa" ? "منطقه / ران" : lang === "hi" ? "क्षेत्र / रन" : "Region / Run";
  const timeColLabel = lang === "fa" ? "ساعت" : lang === "hi" ? "समय" : "Time";
  const durColLabel = lang === "fa" ? "مدت" : lang === "hi" ? "अवधि" : "Duration";
  const totalCols = 1 + matched.length * 3;

  const aoa = [];
  aoa.push([t("compare2Title", lang), ...Array(totalCols - 1).fill("")]);
  aoa.push([`Shift Priority v${APP_VERSION} · Omid Farhadnia · ${new Date().toLocaleDateString(lang === "fa" ? "fa-IR" : "en-CA")}`, ...Array(totalCols - 1).fill("")]);
  const row1 = [dayColLabel];
  matched.forEach((m) => row1.push(`${t("crewWord", lang)} ${m.crew}`, "", ""));
  aoa.push(row1);
  const row2 = [""];
  matched.forEach(() => row2.push(regionColLabel, timeColLabel, durColLabel));
  aoa.push(row2);
  weekdayNames.forEach((wd, di) => {
    const row = [wd];
    matched.forEach((m) => {
      const d = m.days.find((dd) => dd.dayIdx === di);
      if (!d) row.push(offLabel, "", "");
      else row.push(`${regionLabel(d.regionKey, lang)} ${d.code || ""}`.trim(), `${formatExcelTime(d.start)}–${formatExcelTime(d.end)}`, formatDuration(d.hours, lang));
    });
    aoa.push(row);
  });
  const totalsRowIdx = aoa.length;
  const totalsRow = [t("hours", lang)];
  matched.forEach((m) => totalsRow.push(m.totalHours, "", ""));
  aoa.push(totalsRow);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const headerRow1Idx = 2, headerRow2Idx = 3;
  const merges = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: totalCols - 1 } },
    { s: { r: headerRow1Idx, c: 0 }, e: { r: headerRow2Idx, c: 0 } },
  ];
  matched.forEach((_, i) => {
    const startCol = 1 + i * 3;
    merges.push({ s: { r: headerRow1Idx, c: startCol }, e: { r: headerRow1Idx, c: startCol + 2 } });
    merges.push({ s: { r: totalsRowIdx, c: startCol }, e: { r: totalsRowIdx, c: startCol + 2 } });
  });
  ws["!merges"] = merges;
  ws["!cols"] = row1.map((_, ci) => ({ wch: ci === 0 ? 13 : 16 }));
  ws["!rows"] = [{ hpt: 24 }, { hpt: 16 }];

  const setStyle = (r, c, style) => {
    const ref = XLSX.utils.encode_cell({ r, c });
    if (!ws[ref]) ws[ref] = { t: "s", v: "" };
    ws[ref].s = { ...(ws[ref].s || {}), ...style, alignment: { horizontal: "center", vertical: "center", ...(style.alignment || {}) } };
  };

  setStyle(0, 0, { fill: { fgColor: { rgb: "2F5D62" } }, font: { bold: true, sz: 14, color: { rgb: "FFFFFF" } } });
  setStyle(1, 0, { fill: { fgColor: { rgb: "3E7C82" } }, font: { italic: true, sz: 10, color: { rgb: "FFFFFF" } } });
  setStyle(headerRow1Idx, 0, { fill: { fgColor: { rgb: "EAF1EF" } }, font: { bold: true, color: { rgb: "20242B" } } });
  weekdayNames.forEach((wd, di) => setStyle(headerRow2Idx + 1 + di, 0, { fill: { fgColor: { rgb: "FAF8F3" } }, font: { bold: true, color: { rgb: "20242B" } } }));
  setStyle(totalsRowIdx, 0, { fill: { fgColor: { rgb: "EAF1EF" } }, font: { bold: true, color: { rgb: "20242B" } } });

  matched.forEach((m, i) => {
    const pal = PERSON_PALETTE[i % 5];
    const startCol = 1 + i * 3;
    setStyle(headerRow1Idx, startCol, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { bold: true, color: { rgb: hexNoHash(pal.text) } } });
    setStyle(headerRow2Idx, startCol, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { bold: true, sz: 9, color: { rgb: hexNoHash(pal.text) } } });
    setStyle(headerRow2Idx, startCol + 1, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { bold: true, sz: 9, color: { rgb: hexNoHash(pal.text) } } });
    setStyle(headerRow2Idx, startCol + 2, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { bold: true, sz: 9, color: { rgb: hexNoHash(pal.text) } } });
    setStyle(totalsRowIdx, startCol, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { bold: true, color: { rgb: hexNoHash(pal.text) } } });
    weekdayNames.forEach((wd, di) => {
      const rIdx = headerRow2Idx + 1 + di;
      const d = m.days.find((dd) => dd.dayIdx === di);
      if (!d) {
        for (let c = startCol; c < startCol + 3; c++) setStyle(rIdx, c, { fill: { fgColor: { rgb: "FBEAEA" } }, font: { color: { rgb: "B3432A" }, bold: true } });
      } else {
        for (let c = startCol; c < startCol + 3; c++) setStyle(rIdx, c, { fill: { fgColor: { rgb: hexNoHash(pal.bg) } }, font: { color: { rgb: hexNoHash(pal.text) } } });
      }
    });
  });

  applyBorderCenterToRange(ws);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Compare");
  XLSX.writeFile(wb, "crew-comparison.xlsx", { cellStyles: true });
}

// ---------- Compare crews (full week, up to 5) ----------

function CompareTwoPanel({ lang, crews, onClose }) {
  const [crewNums, setCrewNums] = useState(["", ""]);
  const [notFound, setNotFound] = useState(false);
  const [matched, setMatched] = useState([]);

  const updateNum = (idx, v) => setCrewNums((prev) => prev.map((x, i) => (i === idx ? v : x)));
  const addSlot = () => setCrewNums((prev) => (prev.length < 5 ? [...prev, ""] : prev));
  const removeSlot = (idx) => setCrewNums((prev) => (prev.length > 2 ? prev.filter((_, i) => i !== idx) : prev));

  const doCompare = () => {
    const entered = crewNums.map((n) => n.trim()).filter(Boolean);
    const found = entered.map((n) => crews.find((c) => String(c.crew) === n)).filter(Boolean);
    if (found.length !== entered.length || found.length < 2) {
      setNotFound(true);
      setMatched([]);
      return;
    }
    setNotFound(false);
    setMatched(found);
  };

  const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
  const dayCell = (crew, i) => crew.days.find((d) => d.dayIdx === i);

  return (
    <Modal title={t("compare2Title", lang)} onClose={onClose} headerGradient="linear-gradient(135deg, #6D5CE0, #9C8CFB)" accent="#6D5CE0">
      <div style={styles.chipRow}>
        {crewNums.map((n, i) => (
          <div key={i} style={{ display: "flex", gap: 4, alignItems: "center" }}>
            <input
              type="number"
              placeholder={`${t("crewWord", lang)} ${i + 1}`}
              value={n}
              onChange={(e) => updateNum(i, e.target.value)}
              style={{ ...styles.numInputWide, minWidth: 80, width: 80, flex: "none" }}
            />
            {crewNums.length > 2 && (
              <button onClick={() => removeSlot(i)} style={styles.arrowBtn}><X size={13} /></button>
            )}
          </div>
        ))}
        {crewNums.length < 5 && (
          <button onClick={addSlot} style={styles.arrowBtn}><Plus size={14} /></button>
        )}
      </div>
      <button onClick={doCompare} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)", marginTop: 8 }}>
        <Users size={14} /> {t("compareBtn", lang)}
      </button>

      {notFound && (
        <div style={styles.errorBox}><AlertCircle size={15} /><span>{t("crewNotFound", lang)}</span></div>
      )}

      {matched.length >= 2 && (
        <>
          <div style={{ display: "flex", gap: 8, margin: "12px 0" }}>
            <button onClick={() => printCompareTable(matched, lang)} style={styles.smallActionBtn}><Printer size={14} /> {t("printBtn", lang)}</button>
            <button onClick={() => compareToExcel(matched, lang)} style={styles.smallActionBtn}><FileSpreadsheet size={14} /> {t("excelBtn", lang)}</button>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={styles.compareTable}>
              <thead>
                <tr>
                  <th style={styles.compareLabelCell}></th>
                  {matched.map((m, i) => (
                    <th key={m.crew} style={{ ...styles.compareHeadCell, background: PERSON_PALETTE[i % 5].bg, color: PERSON_PALETTE[i % 5].text }}>
                      {t("crewWord", lang)} {String(m.crew)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {weekdayNames.map((wd, di) => (
                  <tr key={di}>
                    <td style={styles.compareLabelCell}>{wd}</td>
                    {matched.map((m, i) => {
                      const d = dayCell(m, di);
                      const pal = PERSON_PALETTE[i % 5];
                      return (
                        <td key={m.crew} style={{ ...styles.compareCell, padding: 4 }}>
                          {d ? (
                            <div style={{ background: pal.bg, border: `1px solid ${pal.border}`, borderRadius: 8, padding: 8, color: pal.text }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10.5, marginBottom: 4, flexWrap: "wrap", justifyContent: "center", color: pal.text }}>
                                <span style={{ width: 7, height: 7, borderRadius: "50%", background: REGION_COLORS[d.regionKey] || "#9AA0A6" }} />
                                {regionLabel(d.regionKey, lang)} · {d.code || "-"}
                              </div>
                              <div style={{ fontWeight: 700, fontSize: 12.5, color: pal.text }}>{formatExcelTime(d.start)}–{formatExcelTime(d.end)}</div>
                              <div style={{ fontSize: 10.5, color: "#5A5A5A" }}>{formatDuration(d.hours, lang)}</div>
                            </div>
                          ) : (
                            <div style={{ border: "1.5px dashed #e0a0a0", background: "#fbeceb", color: "#b3432a", borderRadius: 8, padding: 8, fontWeight: 700, fontSize: 11.5 }}>
                              {t("off", lang)} ✕
                            </div>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
                <tr>
                  <td style={styles.compareLabelCell}>{t("type", lang)}</td>
                  {matched.map((m) => <td key={m.crew} style={styles.compareCell}>{m.type}</td>)}
                </tr>
                <tr>
                  <td style={styles.compareLabelCell}>{t("shift", lang)}</td>
                  {matched.map((m) => <td key={m.crew} style={styles.compareCell}>{m.shiftRaw}</td>)}
                </tr>
                <tr>
                  <td style={styles.compareLabelCell}>{t("daysColumn", lang)}</td>
                  {matched.map((m) => <td key={m.crew} style={styles.compareCell}>{m.workedCount}</td>)}
                </tr>
                <tr>
                  <td style={{ ...styles.compareLabelCell, fontWeight: 700 }}>{t("hours", lang)}</td>
                  {matched.map((m) => <td key={m.crew} style={{ ...styles.compareCell, fontWeight: 700 }}>{m.totalHours}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </>
      )}
    </Modal>
  );
}

// ---------- Result cards ----------

const MEDAL_STYLE = {
  1: { icon: Trophy, color: "#B98A2E", border: "#B98A2E" },
  2: { icon: Medal, color: "#8A8F96", border: "#B8BCC2" },
  3: { icon: Award, color: "#A9713F", border: "#C08A55" },
};

function TopCard({ r, rank, lang, compareSet, toggleCompare, profile }) {
  const inCompare = compareSet.includes(r.crew);
  const medal = MEDAL_STYLE[rank];
  const Icon = medal ? medal.icon : null;
  const ds = displayScore(r.fingerprint);
  const isMine = profile?.crewNumber && String(r.crew) === String(profile.crewNumber);
  return (
    <div style={{ ...styles.topCard, borderColor: isMine ? "var(--accent)" : medal ? medal.border : "var(--border)", ...(isMine ? { borderWidth: 2 } : {}) }}>
      <div style={{ ...styles.topBadge, background: medal ? medal.color : "var(--accent)" }}>
        {Icon ? <Icon size={13} /> : rank}
        {Icon && <span>#{rank}</span>}
      </div>
      <div style={styles.heroTop}>
        <div style={styles.heroCrew}>
          {t("crewWord", lang)} {String(r.crew)}
          {isMine && <span style={styles.mineBadge}><Star size={10} /> {t("myShiftBadge", lang)}</span>}
        </div>
        <div style={{ ...styles.scoreBadge, color: scoreColor(scorePercent(ds)) }}>{scorePercent(ds)}%</div>
      </div>
      <div style={styles.heroMeta}>{r.type} · {r.shiftRaw} · {r.workedCount} {t("days", lang)} · {r.totalHours} {t("hours", lang)}</div>
      <FingerprintList fingerprint={r.fingerprint} lang={lang} />
      <RegionChips regionSummary={r.regionSummary} lang={lang} />
      <button onClick={() => toggleCompare(r.crew)} style={{ ...styles.compareToggle, ...(inCompare ? styles.compareToggleActive : {}) }}>
        {inCompare && <Check size={12} />} {t("addCompare", lang)}
      </button>
    </div>
  );
}

function ResultCard({ r, rank, lang, compareSet, toggleCompare, profile }) {
  const inCompare = compareSet.includes(r.crew);
  const ds = displayScore(r.fingerprint);
  const isMine = profile?.crewNumber && String(r.crew) === String(profile.crewNumber);
  return (
    <div style={{ ...styles.resultCard, ...(isMine ? { borderColor: "var(--accent)", borderWidth: 2 } : {}) }}>
      <div style={styles.resultCardTop}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={styles.rankBadge}>{rank}</span>
          <span style={styles.resultCrew}>{t("crewWord", lang)} {String(r.crew)}</span>
          {isMine && <span style={styles.mineBadge}><Star size={10} /> {t("myShiftBadge", lang)}</span>}
        </div>
        <span style={{ ...styles.scoreBadgeSm, color: scoreColor(scorePercent(ds)) }}>{scorePercent(ds)}%</span>
      </div>
      <div style={styles.resultMeta}>{r.type} · {r.shiftRaw} · {r.workedCount} {t("days", lang)} · {r.totalHours} {t("hours", lang)}</div>
      <FingerprintStrip fingerprint={r.fingerprint} />
      <RegionChips regionSummary={r.regionSummary} lang={lang} />
      <button onClick={() => toggleCompare(r.crew)} style={{ ...styles.compareToggle, ...(inCompare ? styles.compareToggleActive : {}) }}>
        {inCompare && <Check size={12} />} {t("addCompare", lang)}
      </button>
    </div>
  );
}

// ---------- Main component ----------

export default function ShiftPriorityRanker() {
  const [lang, setLang] = useState("en");
  const [themeStyle, setThemeStyle] = useState("universal");
  const [themeMode, setThemeMode] = useState("light");
  const [showPriorityFlow, setShowPriorityFlow] = useState(false);
  // 'settings' | 'help' | 'about' | 'helpMenu' | 'compare2' | 'profile'
  // | 'crewLookup' | 'adminLogin' | 'admin'
  const [activePanel, setActivePanel] = useState(null);

  // Personal profile (name + "my crew number") — loaded once from this
  // browser's own localStorage. Nothing here ever leaves the device: two
  // different people opening the same build each only ever see what THEY
  // saved, never each other's.
  const [profile, setProfile] = useState(() => loadProfile());
  const [showMySchedule, setShowMySchedule] = useState(false);

  // Crew-number -> driver-name directory (auto-read "Driver Name" column,
  // overlaid with manual Admin edits) and whether this tab is unlocked as
  // Admin — see the loadCrewNames/loadAdminSession helpers above.
  const [crewNames, setCrewNames] = useState(() => loadCrewNames());
  const [isAdmin, setIsAdmin] = useState(() => loadAdminSession());
  // The crew picked from the "Browse crews" lookup panel, or null.
  const [lookupCrew, setLookupCrew] = useState(null);
  const [swapViewCrew, setSwapViewCrew] = useState(null);

  // Daily Log -- see loadDailyLogAccess/saveDailyLogAccess above.
  const [dailyLogAccess, setDailyLogAccess] = useState(() => loadDailyLogAccess());
  // Open to everyone now (Omid asked to drop the per-crew Admin gate so
  // nobody has to log in as Admin just to use their own Daily Shift Log).
  // dailyLogAccess/the Admin panel list are kept around, unused, in case
  // per-crew gating is wanted again later.
  const dailyLogVisible = true;

  // Restore the last successfully-parsed schedule (if any) so the app opens
  // straight to it instead of forcing a re-upload every time — see
  // LAST_FILE_KEY above. `workbook`/`sheetNames` are NOT restorable (the raw
  // parsed XLSX workbook isn't stored, only the already-extracted crew
  // data), so switching sheets on a restored file isn't available until a
  // fresh upload — a fine trade-off since almost nobody needs that.
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [lastFile] = useState(() => loadLastFile());
  const [selectedSheet, setSelectedSheet] = useState(() => lastFile?.sheetName || "");
  const [fileName, setFileName] = useState(() => lastFile?.fileName || "");
  const [parsed, setParsed] = useState(() => lastFile?.parsed || null);
  const [error, setError] = useState("");

  const [priorityList, setPriorityList] = useState([]);
  // Which weekdays the "day_off" dynamic criterion currently means, e.g.
  // { day_off: [5, 6] } for Friday + Saturday (up to DAY_OFF_MAX_DAYS).
  // Kept separate from priorityList because it is per-criterion
  // configuration, not membership/order.
  const [dayOffChoices, setDayOffChoices] = useState({});
  const [showResults, setShowResults] = useState(false);
  const [compareSet, setCompareSet] = useState([]);
  const [results, setResults] = useState([]);

  // Brief branded loading screen: once right after the app mounts, and
  // again (briefly) each time "Calculate & Rank" runs — see runCompute().
  const [booting, setBooting] = useState(true);
  const [computing, setComputing] = useState(false);
  useEffect(() => {
    const timer = setTimeout(() => setBooting(false), 700);
    return () => clearTimeout(timer);
  }, []);

  // Ask the browser not to evict this site's localStorage (profile, saved
  // schedule, history) under storage pressure or an inactivity policy —
  // mainly relevant to Safari/iOS, where a PWA's local data can otherwise be
  // cleared after a period of disuse. Best-effort only: not every browser
  // implements this, and even where it does, it's a request, not a
  // guarantee, so this can never fully replace re-uploading if data really
  // is gone — it just makes that less likely.
  useEffect(() => {
    try { navigator.storage?.persist?.(); } catch { /* not supported here, ignore */ }
  }, []);

  // The results panel only ever reflects a snapshot from the last time
  // "Calculate & Rank" was pressed. If the person then adds/removes/reorders
  // a priority, or changes which weekdays a day-off criterion means, that
  // old snapshot must not keep showing as if it were still current — most
  // visibly when every priority is removed: the button becomes disabled
  // (nothing to compute), so without this the stale top results would stay
  // on screen forever with no way to refresh them. Hiding results here means
  // "Calculate & Rank" must be pressed again after any such change.
  useEffect(() => {
    setShowResults(false);
  }, [priorityList, dayOffChoices]);

  const dir = lang === "fa" ? "rtl" : "ltr";
  const palette = getPalette(themeStyle, themeMode);
  // Keep the installed app's window/title bar in step with the active theme.
  useEffect(() => {
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", themeMode === "dark" ? palette.card : palette.accent);
  }, [palette, themeMode]);

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    setParsed(null);
    setShowResults(false);
    setCompareSet([]);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array", cellStyles: true });
        setWorkbook(wb);
        const names = visibleSheetNames(wb);
        setSheetNames(names);
        const last = names[names.length - 1];
        setSelectedSheet(last);
        runParse(wb, last, file.name);
      } catch (err) {
        setError(t("fileError", lang));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const runParse = (wb, sheetName, fileNameOverride) => {
    setError("");
    setShowResults(false);
    setCompareSet([]);
    const sheet = wb.Sheets[sheetName];
    const result = parseSchedule(sheet);
    if (!result.ok) {
      setParsed(null);
      clearLastFileStorage();
      if (result.reason === "no_layout" && result.missing) {
        setError(`${t("errMissingColumns", lang)} ${result.missing.join(", ")}`);
      } else if (result.reason === "no_data") {
        setError(t("errNoData", lang));
      } else {
        setError(t("errNoLayout", lang));
      }
      return;
    }
    setParsed(result);
    // fileNameOverride covers the exact moment a brand-new file is uploaded,
    // where the `fileName` state variable in this closure may not have
    // caught up to the setFileName() call yet — see handleFile().
    saveLastFile({ fileName: fileNameOverride ?? fileName, sheetName, parsed: result });
  };

  const changeSheet = (name) => {
    setSelectedSheet(name);
    if (workbook) runParse(workbook, name);
  };

  const addCriterion = (id) => {
    setPriorityList((prev) => (prev.includes(id) || prev.length >= CRITERIA_MAX ? prev : [...prev, id]));
  };
  const removeCriterion = (id) => setPriorityList((prev) => prev.filter((x) => x !== id));
  const moveCriterion = (idx, d) => {
    setPriorityList((prev) => {
      const arr = [...prev];
      const target = idx + d;
      if (target < 0 || target >= arr.length) return arr;
      [arr[idx], arr[target]] = [arr[target], arr[idx]];
      return arr;
    });
  };

  const runCompute = () => {
    if (!parsed || priorityList.length === 0) return;
    const hasUnsetDayOff = priorityList.some((id) => {
      const base = CRITERIA_CATALOG.find((c) => c.id === id);
      return base?.dynamic && (!Array.isArray(dayOffChoices[id]) || dayOffChoices[id].length === 0);
    });
    if (hasUnsetDayOff) { setError(t("errDayOffUnset", lang)); return; }
    setError("");
    // The ranking itself is near-instant even for a big crew sheet, but a
    // result swap with zero visual feedback reads as if the button did
    // nothing. Showing the loading animation for a small fixed minimum
    // (rather than exactly as long as the math takes) gives a consistent,
    // deliberate feel instead of an inconsistent flash.
    setComputing(true);
    setTimeout(() => {
      const computed = computeResults(parsed.crews, priorityList, dayOffChoices);
      setResults(computed);
      setShowResults(true);
      setComputing(false);
    }, 500);
  };

  const toggleCompare = (crew) => {
    setCompareSet((prev) => (prev.includes(crew) ? prev.filter((c) => c !== crew) : prev.length < 5 ? [...prev, crew] : prev));
  };
  const compareResults = useMemo(() => results.filter((r) => compareSet.includes(r.crew)), [results, compareSet]);

  const exportExcel = () => {
    const critLabels = priorityList.map((id, i) => `${i + 1}. ${results[0]?.fingerprint?.[i]?.label?.[lang] || CRITERIA_CATALOG.find((c) => c.id === id)?.label[lang]}`);
    const headers = [t("rank", lang), t("crewWord", lang), t("type", lang), t("shift", lang), t("days", lang), t("hours", lang), t("region", lang), ...critLabels];
    const dataRows = results.map((r, idx) => [
      idx + 1,
      r.crew,
      r.type,
      r.shiftRaw,
      r.workedCount,
      r.totalHours,
      Object.entries(r.regionSummary).map(([k, arr]) => `${arr.length} ${regionLabel(k, lang)} (${arr.join(", ")} ${t("hoursWord", lang)})`).join("; "),
      ...r.fingerprint.map((f) => `${f.score}%`),
    ]);

    const totalCols = headers.length;
    const aoa = [
      [t("resultsTitle", lang), ...Array(totalCols - 1).fill("")],
      [`Shift Priority v${APP_VERSION} · Omid Farhadnia · ${new Date().toLocaleDateString(lang === "fa" ? "fa-IR" : "en-CA")}`, ...Array(totalCols - 1).fill("")],
      headers,
      ...dataRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: totalCols - 1 } },
    ];
    ws["!cols"] = headers.map((h, i) => ({ wch: i === 0 ? 8 : i === 6 ? 30 : 16 }));
    ws["!rows"] = [{ hpt: 24 }, { hpt: 16 }];

    const setStyle = (r, c, style) => {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      ws[ref].s = { ...(ws[ref].s || {}), ...style };
    };
    setStyle(0, 0, { fill: { fgColor: { rgb: "2F5D62" } }, font: { bold: true, sz: 14, color: { rgb: "FFFFFF" } } });
    setStyle(1, 0, { fill: { fgColor: { rgb: "3E7C82" } }, font: { italic: true, sz: 10, color: { rgb: "FFFFFF" } } });
    headers.forEach((_, c) => setStyle(2, c, { fill: { fgColor: { rgb: "EAF1EF" } }, font: { bold: true, color: { rgb: "20242B" } } }));

    const medalFill = { 0: "FBF1DA", 1: "F2F2F2", 2: "F6E9DC" };
    dataRows.forEach((_, i) => {
      if (medalFill[i]) {
        headers.forEach((_, c) => setStyle(3 + i, c, { fill: { fgColor: { rgb: medalFill[i] } } }));
      }
    });

    applyBorderCenterToRange(ws);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, "shift-priority-report.xlsx", { cellStyles: true });
  };

  const groupedCatalog = ["shift", "days", "region", "special"].map((g) => ({ group: g, items: CRITERIA_CATALOG.filter((c) => c.group === g) }));
  const groupLabelKey = { shift: "catalogShift", days: "catalogDays", region: "catalogRegion", special: "catalogSpecial" };

  const rootVars = {
    "--bg": palette.bg, "--card": palette.card, "--border": palette.border, "--text": palette.text,
    "--muted": palette.muted, "--accent": palette.accent, "--accent2": palette.accent2, "--radius": palette.radius,
    ...(themeMode === "dark"
      ? { "--score-high": "#FACC15", "--score-mid": "#4ADE80", "--score-low": "#7D8796" }
      : { "--score-high": "#A16207", "--score-mid": "#15803D", "--score-low": "#7B8494" }),
  };

  return (
    <div dir={dir} style={{ ...styles.page, ...rootVars }}>
      <style>{`
        .print-report { display: none; }
        @media print {
          body { background: #fff; }
          .no-print { display: none !important; }
          .print-report { display: block !important; }
        }
        .spp-ring {
          position: absolute;
          inset: 0;
          border-radius: 50%;
          border: 3px solid var(--border);
          border-top-color: var(--accent);
          animation: spp-spin 0.85s linear infinite;
        }
        @keyframes spp-spin { to { transform: rotate(360deg); } }
      `}</style>

      {(booting || computing) && (
        <div className="no-print" style={styles.splashOverlay}>
          <div style={styles.splashInner}>
            <div style={styles.splashRing}>
              <div className="spp-ring" />
              <img src="./logo.png" alt="" style={styles.splashLogo} />
            </div>
            <p style={styles.splashText}>{t(booting ? "loadingBoot" : "loadingCompute", lang)}</p>
          </div>
        </div>
      )}

      <div className="no-print" style={{ ...styles.topBar, flexDirection: dir === "rtl" ? "row" : "row-reverse" }}>
        <div style={styles.langBtn}>
          {["fa", "en", "hi"].map((l) => (
            <button key={l} onClick={() => setLang(l)} style={{ ...styles.langPill, ...(lang === l ? styles.langPillActive : {}) }}>
              {l === "fa" ? "فا" : l === "en" ? "EN" : "हि"}
            </button>
          ))}
        </div>
        <div style={{ position: "relative" }}>
          {showPriorityFlow && (
            <button onClick={() => setShowPriorityFlow(false)} style={styles.menuBtn}><ArrowLeft size={16} /></button>
          )}
        </div>
      </div>

      {activePanel === "profile" && (
        <ProfilePanel
          lang={lang}
          profile={profile}
          onSave={(p) => { setProfile(p); saveProfile(p); }}
          onClear={() => { setProfile(null); clearProfileStorage(); }}
          onClose={() => setActivePanel(null)}
        />
      )}
      {activePanel === "settings" && (
        <SettingsPanel lang={lang} themeStyle={themeStyle} setThemeStyle={setThemeStyle} themeMode={themeMode} setThemeMode={setThemeMode} onClose={() => setActivePanel(null)} />
      )}
      {activePanel === "help" && <HelpPanel lang={lang} onClose={() => setActivePanel(null)} />}
      {activePanel === "about" && <AboutPanel lang={lang} onClose={() => setActivePanel(null)} />}
      {activePanel === "helpMenu" && (
        <HelpMenuPanel
          lang={lang}
          onPick={(which) => {
            if (which === "report") { openFeedbackEmail(lang); setActivePanel(null); }
            else setActivePanel(which);
          }}
          onClose={() => setActivePanel(null)}
        />
      )}
      {activePanel === "compare2" && parsed && (
        <CompareTwoPanel lang={lang} crews={parsed.crews} onClose={() => setActivePanel(null)} />
      )}
      {activePanel === "swapFinder" && parsed && (
        <SwapFinderPanel
          lang={lang}
          crews={parsed.crews}
          crewNames={crewNames}
          profile={profile}
          onViewCrew={(c) => setSwapViewCrew(c)}
          onClose={() => setActivePanel(null)}
        />
      )}
      {swapViewCrew && (
        <MyScheduleModal
          crew={swapViewCrew}
          name={resolveCrewName(swapViewCrew.crew, parsed?.crews, crewNames)}
          lang={lang}
          themeMode={themeMode}
          onClose={() => setSwapViewCrew(null)}
          onBack={() => setSwapViewCrew(null)}
        />
      )}
      {activePanel === "crewLookup" && parsed && (
        <CrewLookupPanel
          lang={lang}
          crews={parsed.crews}
          crewNames={crewNames}
          onPick={(crewNumber) => {
            const c = parsed.crews.find((cc) => String(cc.crew) === String(crewNumber));
            if (c) { setLookupCrew(c); setActivePanel(null); }
          }}
          onClose={() => setActivePanel(null)}
        />
      )}
      {activePanel === "dailyLog" && dailyLogVisible && (
        <DailyLogPanel lang={lang} onClose={() => setActivePanel(null)} />
      )}
      {lookupCrew && (
        <MyScheduleModal
          crew={lookupCrew}
          name={resolveCrewName(lookupCrew.crew, parsed?.crews, crewNames)}
          lang={lang}
          themeMode={themeMode}
          onClose={() => setLookupCrew(null)}
          onBack={() => { setLookupCrew(null); setActivePanel("crewLookup"); }}
        />
      )}
      {activePanel === "adminLogin" && (
        <AdminLoginModal
          lang={lang}
          onSuccess={() => { setIsAdmin(true); setActivePanel("admin"); }}
          onClose={() => setActivePanel(null)}
        />
      )}
      {activePanel === "admin" && isAdmin && (
        <AdminPanel
          lang={lang}
          crews={parsed?.crews}
          crewNames={crewNames}
          setCrewNames={setCrewNames}
          dailyLogAccess={dailyLogAccess}
          setDailyLogAccess={setDailyLogAccess}
          onLogout={() => { setIsAdmin(false); saveAdminSession(false); setActivePanel(null); }}
          onClose={() => setActivePanel(null)}
        />
      )}
      {showMySchedule && parsed && profile?.crewNumber && (() => {
        const myCrew = parsed.crews.find((c) => String(c.crew) === String(profile.crewNumber));
        return myCrew ? (
          <MyScheduleModal
            crew={myCrew}
            name={resolveCrewName(myCrew.crew, parsed.crews, crewNames)}
            lang={lang}
            themeMode={themeMode}
            onClose={() => setShowMySchedule(false)}
          />
        ) : null;
      })()}

      <div>
        <header className="no-print" style={styles.header}>
          <img src="./logo.png" alt="" style={styles.headerLogo} />
          <div style={styles.routeDots}>
            <span style={styles.dot} /><span style={styles.routeLine} /><span style={styles.dot} /><span style={styles.routeLine} /><span style={{ ...styles.dot, background: "var(--accent2)" }} />
          </div>
          <h1 style={styles.title}>{t("title", lang)}</h1>
          <p style={styles.subtitle}>{t("subtitle", lang)}</p>
          {profile?.firstName && (
            <p style={styles.welcomeLine}>{t("welcomeBack", lang)} <b>{profile.firstName}</b></p>
          )}
          {showPriorityFlow ? (
            <DateTimeWidget lang={lang} />
          ) : (
            // Home only: "My Shift" block physically LEFT of the clock in every
            // language (the row is forced LTR; each block keeps the page direction).
            <div style={{ display: "flex", justifyContent: "center", alignItems: "stretch", gap: 8, flexWrap: "nowrap", direction: "ltr" }}>
              {(() => {
                const myCrew = parsed && profile?.crewNumber ? parsed.crews.find((c) => String(c.crew) === String(profile.crewNumber)) : null;
                const today = myCrew ? myCrew.days.find((d) => d.dayIdx === new Date().getDay()) : null;
                const onClick = () => {
                  if (!profile?.crewNumber) setActivePanel("profile");
                  else if (!parsed) setShowPriorityFlow(true);
                  else if (myCrew) setShowMySchedule(true);
                  else setActivePanel("profile");
                };
                const sub = !profile?.crewNumber ? t("myShiftHomeNoCrew", lang)
                  : !parsed ? t("myShiftHomeNoFile", lang)
                  : !myCrew ? t("crewNumberNotInFile", lang)
                  : null;
                return (
                  <button onClick={onClick} style={{ direction: dir, marginTop: 12, flex: "0 0 auto", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 2, width: 104, padding: "8px 8px", border: "none", borderRadius: "var(--radius)", cursor: "pointer", color: "#fff", background: "linear-gradient(135deg, #0B8F87, #19B97A)", boxShadow: "0 6px 14px rgba(11,143,135,0.28)", fontFamily: "inherit" }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 4, fontWeight: 800, fontSize: 13.5, whiteSpace: "nowrap" }}><Star size={13} color="#fff" /> {t("myShiftHomeTitle", lang)}</span>
                    {myCrew && <span style={{ fontSize: 11.5, opacity: 0.9 }}>{t("crewWord", lang)} {String(myCrew.crew)}</span>}
                    {myCrew && (today ? (
                      <><span style={{ fontSize: 11, opacity: 0.9 }}>{t("todayLabel", lang)}</span><bdi dir="ltr" style={{ fontSize: 12.5, fontWeight: 800, whiteSpace: "nowrap" }}>{formatExcelTime(today.start)}–{formatExcelTime(today.end)}</bdi></>
                    ) : (
                      <span style={{ fontSize: 12.5, fontWeight: 700, textAlign: "center" }}>{t("todayLabel", lang)}: {t("off", lang)}</span>
                    ))}
                    {sub && <span style={{ fontSize: 10.5, opacity: 0.9, lineHeight: 1.35, textAlign: "center" }}>{sub}</span>}
                  </button>
                );
              })()}
              <div style={{ direction: dir, display: "flex", flex: "0 1 auto", minWidth: 0 }}><DateTimeWidget lang={lang} /></div>
            </div>
          )}
        </header>

        {!showPriorityFlow && (
          <div className="no-print">
            <button onClick={() => setShowPriorityFlow(true)} style={styles.heroTile}>
              <span style={styles.heroTileIcon}><Star size={22} color="#fff" /></span>
              <span style={{ flex: 1 }}>
                <div style={styles.heroTileTitle}>{t("title", lang)}</div>
                <div style={styles.heroTileSubtitle}>{t("subtitle", lang)}</div>
              </span>
              <ArrowLeft size={16} color="#fff" />
            </button>

            <div style={styles.hubGroupLabel}>{t("hubGroupCrews", lang)}</div>
            <div style={styles.hubGrid}>
              <button onClick={() => setActivePanel("compare2")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #6D5CE0, #9C8CFB)" }}>
                <GitCompare size={18} color="#fff" />
                <span style={styles.hubTileLabel}>{t("compare2Title", lang)}</span>
              </button>
              <button onClick={() => setActivePanel("crewLookup")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #0EA37E, #3DDC97)" }}>
                <Search size={18} color="#fff" />
                <span style={styles.hubTileLabel}>{t("crewLookupTitle", lang)}</span>
              </button>
              <button onClick={() => setActivePanel("swapFinder")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #D9480F, #F59F00)" }}>
                <CalendarOff size={18} color="#fff" />
                <span style={styles.hubTileLabel}>{t("hubSwapFinder", lang)}</span>
              </button>
            </div>

            <div style={styles.hubGroupLabel}>{t("hubGroupMe", lang)}</div>
            <div style={styles.hubGrid}>
              <button onClick={() => setActivePanel("profile")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #E0447F, #F17CA6)" }}>
                <User size={18} color="#fff" />
                <span style={styles.hubTileLabel}>{t("profileTitle", lang)}</span>
              </button>
              {dailyLogVisible && (
                <button onClick={() => setActivePanel("dailyLog")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #2463EB, #4F8CFB)" }}>
                  <ClipboardList size={18} color="#fff" />
                  <span style={styles.hubTileLabel}>{t("dailyLogMenuLabel", lang)}</span>
                </button>
              )}
              <button onClick={() => setActivePanel("settings")} style={{ ...styles.hubTile, background: "linear-gradient(135deg, #E08A1E, #F6B93B)" }}>
                <Sun size={18} color="#fff" />
                <span style={styles.hubTileLabel}>{t("settingsTitle", lang)}</span>
              </button>
            </div>

            <div style={styles.hubListCard}>
              <button style={styles.menuItem} onClick={() => setActivePanel("helpMenu")}>
                <HelpCircle size={15} /> {t("helpMenuLabel", lang)}
              </button>
              <button style={styles.menuItem} onClick={() => setActivePanel(isAdmin ? "admin" : "adminLogin")}>
                <Shield size={15} /> {t("adminMenuLabel", lang)}
              </button>
              <button style={{ ...styles.menuItem, color: "#B3432A" }} onClick={() => exitApp(lang)}>
                <LogOut size={15} /> {t("exitApp", lang)}
              </button>
            </div>
          </div>
        )}

        {showPriorityFlow && (
        <div style={{ "--accent": "#D12F48" }}>
        {profile?.crewNumber && parsed && (() => {
          const myCrew = parsed.crews.find((c) => String(c.crew) === String(profile.crewNumber));
          return (
            <div className="no-print" style={styles.myShiftCard}>
              <div>
                <div style={{ fontSize: 11.5, fontWeight: 700, color: "var(--muted)" }}>{t("myShiftCard", lang)}</div>
                <div style={{ fontSize: 16, fontWeight: 800 }}>
                  {t("crewWord", lang)} {profile.crewNumber}
                  {(() => {
                    const myName = resolveCrewName(profile.crewNumber, parsed.crews, crewNames);
                    return myName ? <span style={{ fontWeight: 700 }}> · {myName}</span> : null;
                  })()}
                  {myCrew && <span style={{ fontWeight: 600, fontSize: 12.5, color: "var(--muted)" }}> · {myCrew.type} · {myCrew.shiftRaw}</span>}
                </div>
                {!myCrew && <div style={{ fontSize: 11.5, color: "#B3432A", marginTop: 2 }}>{t("crewNumberNotInFile", lang)}</div>}
              </div>
              {myCrew && (
                <button onClick={() => setShowMySchedule(true)} style={{ ...styles.smallActionBtn, background: "var(--accent)", color: "#fff", borderColor: "var(--accent)" }}>
                  <Star size={13} /> {t("viewMySchedule", lang)}
                </button>
              )}
            </div>
          );
        })()}

        <section className="no-print" style={styles.card}>
          <div style={styles.stepLabel}>{t("uploadStep", lang)}</div>
          {!parsed && (
            <label style={styles.uploadZone}>
              <Upload size={22} color="var(--accent)" />
              <span style={{ marginTop: 8, fontSize: 14, color: "var(--text)" }}>{fileName || t("uploadPlaceholder", lang)}</span>
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          )}

          {parsed && (
            <div style={styles.fileSummaryRow}>
              <FileSpreadsheet size={18} color="var(--accent)" style={{ flexShrink: 0 }} />
              <span style={styles.fileSummaryText}>
                {fileName} · {parsed.crews.length} {t("crewsFoundWord", lang)}{sheetNames.length > 1 ? ` · ${selectedSheet}` : ""}
              </span>
              <label style={styles.fileChangeLink}>
                {t("changeFileLink", lang)}
                <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
              </label>
            </div>
          )}

          {sheetNames.length > 1 && (
            <div style={{ marginTop: 12 }}>
              <div style={styles.smallLabel}>{t("sheetLabel", lang)}</div>
              <div style={styles.radioRow}>
                {sheetNames.map((s) => (
                  <label key={s} style={styles.radioLabel}>
                    <input type="radio" name="sheet" checked={selectedSheet === s} onChange={() => changeSheet(s)} />
                    {s}
                  </label>
                ))}
              </div>
            </div>
          )}

          {error && (
            <div style={styles.errorBox}>
              <AlertCircle size={15} style={{ flexShrink: 0, marginTop: 1 }} />
              <span>{error}</span>
            </div>
          )}

          {parsed && !parsed.colorsDetected && (
            <div style={styles.warnLine}>
              <AlertCircle size={13} />
              <span>{t("colorsDetectedNo", lang)}</span>
            </div>
          )}
        </section>

        {parsed && (
          <section className="no-print" style={styles.card}>
            <div style={styles.stepLabel}>{t("prefsStep", lang)}</div>
            <p style={styles.hint}>{t("builderHint", lang)}</p>

            {groupedCatalog.map(({ group, items }) => (
              <div key={group} style={styles.prefGroup}>
                <div style={styles.prefTitle}>{t(groupLabelKey[group], lang)}</div>
                <div style={styles.prefChipRow}>
                  {items.map((item, i) => {
                    const rankIdx = priorityList.indexOf(item.id);
                    const selected = rankIdx !== -1;
                    const isLastOdd = items.length % 2 === 1 && i === items.length - 1;
                    return (
                      <button
                        key={item.id}
                        onClick={() => (selected ? removeCriterion(item.id) : addCriterion(item.id))}
                        style={{ ...styles.prefChip, ...(selected ? { ...styles.prefChipActive, borderColor: criterionColor(item.id) } : {}), ...(isLastOdd ? { gridColumn: "1 / span 2" } : {}) }}
                      >
                        {selected ? <span style={styles.chipRankBadge}>{rankIdx + 1}</span> : <Plus size={12} />}
                        {item.group === "region" && <span style={{ ...styles.regionDot, background: criterionColor(item.id) }} />}
                        {item.dynamic && <CalendarOff size={12} />}
                        {resolveCriterionLabel(item.id, dayOffChoices)[lang]}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}

            <div style={styles.priorityPanel}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={styles.prefTitle}>{t("yourPriorities", lang)}</div>
                {priorityList.length > 0 && (
                  <button onClick={() => setPriorityList([])} style={styles.resetPriorityBtn}>
                    <RotateCcw size={12} /> {t("resetPriorities", lang)}
                  </button>
                )}
              </div>
              {priorityList.length === 0 && <p style={styles.hint}>{t("emptyPriorities", lang)}</p>}
              {priorityList.length > 0 && (
                <div style={styles.priorityChipRow}>
                  {priorityList.map((id, idx) => (
                    <span key={id} style={{ ...styles.priorityChip, borderColor: criterionColor(id) }}>
                      <span style={{ ...styles.priorityChipRank, background: criterionColor(id) }}>{idx + 1}</span>
                      <span style={styles.priorityChipLabel}>{resolveCriterionLabel(id, dayOffChoices)[lang]}</span>
                      <button onClick={() => moveCriterion(idx, -1)} style={styles.priorityChipBtn} disabled={idx === 0}><ArrowUp size={12} /></button>
                      <button onClick={() => moveCriterion(idx, 1)} style={styles.priorityChipBtn} disabled={idx === priorityList.length - 1}><ArrowDown size={12} /></button>
                      <button onClick={() => removeCriterion(id)} style={{ ...styles.priorityChipBtn, color: "#B3432A" }}><X size={12} /></button>
                    </span>
                  ))}
                </div>
              )}
              {priorityList.map((id) => {
                const crit = CRITERIA_CATALOG.find((c) => c.id === id);
                if (!crit?.dynamic) return null;
                const weekdayNames = WEEKDAY_LABELS[lang] || WEEKDAY_LABELS.en;
                const chosenDays = Array.isArray(dayOffChoices[id]) ? dayOffChoices[id] : [];
                return (
                  <div key={id} style={styles.dayPickBlock}>
                    <span style={styles.dayPickHint}>{resolveCriterionLabel(id, dayOffChoices)[lang]} — {t("pickDayHint", lang)}</span>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 5 }}>
                      {weekdayNames.map((wd, di) => {
                        const isChosen = chosenDays.includes(di);
                        const atMax = chosenDays.length >= DAY_OFF_MAX_DAYS;
                        return (
                          <button
                            key={di}
                            disabled={!isChosen && atMax}
                            onClick={() => setDayOffChoices((prev) => {
                              const cur = Array.isArray(prev[id]) ? prev[id] : [];
                              const next = isChosen
                                ? cur.filter((d) => d !== di)
                                : (cur.length >= DAY_OFF_MAX_DAYS ? cur : [...cur, di]);
                              return { ...prev, [id]: next };
                            })}
                            style={{
                              ...styles.dayPickBtn,
                              ...(isChosen ? styles.dayPickBtnActive : {}),
                              ...(!isChosen && atMax ? { opacity: 0.45, cursor: "not-allowed" } : {}),
                            }}
                          >
                            {wd}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>

            <button onClick={runCompute} disabled={priorityList.length === 0 || computing} style={{ ...styles.computeBtn, opacity: priorityList.length === 0 || computing ? 0.5 : 1 }}>
              <ListChecks size={17} />
              {t("computeBtn", lang)}
            </button>
          </section>
        )}

        {showResults && results.length > 0 && (
          <section className="no-print" style={styles.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
              <div style={styles.stepLabel}>{t("resultsTitle", lang)}</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => printResultsTable(results, priorityList, lang)} style={styles.smallActionBtn}><Printer size={14} /> {t("printBtn", lang)}</button>
                <button onClick={exportExcel} style={styles.smallActionBtn}><FileSpreadsheet size={14} /> {t("excelBtn", lang)}</button>
              </div>
            </div>

            <div style={styles.topLabel}>{t("best", lang)}</div>
            <div style={styles.topGrid}>
              {results.slice(0, 5).map((r, idx) => (
                <TopCard key={r.crew + "-top-" + idx} r={r} rank={idx + 1} lang={lang} compareSet={compareSet} toggleCompare={toggleCompare} profile={profile} />
              ))}
            </div>

            {results.length > 5 && (
              <>
                <div style={{ ...styles.topLabel, marginTop: 18 }}>{t("rest", lang)}</div>
                <div style={styles.grid}>
                  {results.slice(5).map((r, idx) => (
                    <ResultCard key={r.crew + "-" + idx} r={r} rank={idx + 6} lang={lang} compareSet={compareSet} toggleCompare={toggleCompare} profile={profile} />
                  ))}
                </div>
              </>
            )}
          </section>
        )}

        {compareResults.length > 0 && (
          <section className="no-print" style={styles.card}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
              <GitCompare size={17} color="var(--accent)" />
              <div style={styles.stepLabel}>{t("compareTitle", lang)}</div>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={styles.compareTable}>
                <thead>
                  <tr>
                    <th style={styles.compareLabelCell}></th>
                    {compareResults.map((r) => (
                      <th key={r.crew} style={styles.compareHeadCell}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
                          {t("crewWord", lang)} {String(r.crew)}
                          <button onClick={() => toggleCompare(r.crew)} style={styles.compareRemoveBtn}><X size={12} /></button>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  <tr><td style={styles.compareLabelCell}>{t("type", lang)}</td>{compareResults.map((r) => <td key={r.crew} style={styles.compareCell}>{r.type}</td>)}</tr>
                  <tr><td style={styles.compareLabelCell}>{t("shift", lang)}</td>{compareResults.map((r) => <td key={r.crew} style={styles.compareCell}>{r.shiftRaw}</td>)}</tr>
                  <tr><td style={styles.compareLabelCell}>{t("daysColumn", lang)}</td>{compareResults.map((r) => <td key={r.crew} style={styles.compareCell}>{r.workedCount}</td>)}</tr>
                  <tr><td style={styles.compareLabelCell}>{t("hours", lang)}</td>{compareResults.map((r) => <td key={r.crew} style={styles.compareCell}>{r.totalHours}</td>)}</tr>
                  <tr>
                    <td style={styles.compareLabelCell}>{t("region", lang)}</td>
                    {compareResults.map((r) => <td key={r.crew} style={styles.compareCell}><RegionChips regionSummary={r.regionSummary} lang={lang} /></td>)}
                  </tr>
                  {priorityList.map((id, i) => {
                    const critLabel = compareResults[0]?.fingerprint?.[i]?.label?.[lang] ?? resolveCriterionLabel(id, dayOffChoices)[lang];
                    return (
                      <tr key={id}>
                        <td style={styles.compareLabelCell}>{i + 1}. {critLabel}</td>
                        {compareResults.map((r) => (
                          <td key={r.crew} style={{ ...styles.compareCell, fontWeight: 700 }}>{r.fingerprint[i]?.score ?? 0}%</td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {results.length > 0 && (
          <div className="print-report">
            <h2 style={{ fontFamily: "Tahoma, sans-serif" }}>{t("resultsTitle", lang)}</h2>
            <p>{t("reportDate", lang)}: {new Date().toLocaleDateString(lang === "fa" ? "fa-IR" : lang === "hi" ? "hi-IN" : "en-CA")}</p>
            <table style={styles.printTable}>
              <thead>
                <tr>
                  {[t("rank", lang), t("crewWord", lang), t("type", lang), t("shift", lang), t("daysColumn", lang), t("hours", lang), t("region", lang)].map((h) => (
                    <th key={h} style={styles.printTh}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {results.map((r, idx) => (
                  <tr key={idx}>
                    <td style={styles.printTd}>{idx + 1}</td>
                    <td style={styles.printTd}>{String(r.crew)}</td>
                    <td style={styles.printTd}>{r.type}</td>
                    <td style={styles.printTd}>{r.shiftRaw}</td>
                    <td style={styles.printTd}>{r.workedCount}</td>
                    <td style={styles.printTd}>{r.totalHours}</td>
                    <td style={styles.printTd}>{Object.entries(r.regionSummary).map(([k, arr]) => `${arr.length} ${regionLabel(k, lang)} (${arr.join(", ")} ${t("hoursWord", lang)})`).join("; ")}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {parsed && (
          <button
            className="no-print"
            onClick={() => { setWorkbook(null); setSheetNames([]); setParsed(null); setShowResults(false); setFileName(""); setError(""); setCompareSet([]); setPriorityList([]); setResults([]); clearLastFileStorage(); }}
            style={styles.resetBtn}
          >
            <RotateCcw size={14} /> {t("resetBtn", lang)}
          </button>
        )}

        </div>
        )}

        <footer className="no-print" style={styles.footer}>
          Shift Priority v{APP_VERSION} · © Omid Farhadnia · MIT License (Open Source)
        </footer>
      </div>
    </div>
  );
}

const styles = {
  page: { background: "var(--bg)", minHeight: "100vh", padding: "20px 14px 40px", fontFamily: "'Vazirmatn', Tahoma, system-ui, sans-serif", color: "var(--text)", position: "relative" },
  topBar: { display: "flex", justifyContent: "space-between", alignItems: "center", position: "relative", zIndex: 20 },
  langBtn: { display: "flex", gap: 4 },
  langPill: { fontSize: 11.5, padding: "6px 9px", borderRadius: 20, border: "1px solid var(--border)", background: "var(--card)", cursor: "pointer", color: "var(--text)" },
  langPillActive: { background: "var(--accent)", borderColor: "var(--accent)", color: "#fff" },
  menuBtn: { display: "flex", alignItems: "center", justifyContent: "center", width: 30, height: 30, borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", cursor: "pointer", color: "var(--text)" },
  menuDropdown: { position: "absolute", top: 36, background: "var(--card)", border: "1px solid var(--border)", borderRadius: 10, boxShadow: "0 6px 20px rgba(0,0,0,0.12)", padding: 6, display: "flex", flexDirection: "column", gap: 2, width: "max-content", minWidth: 220, maxWidth: "min(280px, calc(100vw - 24px))", zIndex: 200 },
  menuItem: { display: "flex", alignItems: "center", gap: 8, fontSize: 13, padding: "8px 10px", borderRadius: 7, border: "none", background: "transparent", color: "var(--text)", cursor: "pointer", textAlign: "start", whiteSpace: "nowrap" },
  heroTile: { display: "flex", alignItems: "center", gap: 12, width: "100%", textAlign: "start", border: "none", borderRadius: "var(--radius)", padding: 16, marginBottom: 18, background: "linear-gradient(135deg, #E8435A, #F7934C)", cursor: "pointer", boxShadow: "0 6px 16px rgba(232,67,90,0.25)" },
  heroTileIcon: { width: 40, height: 40, borderRadius: 12, background: "rgba(255,255,255,0.18)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  heroTileTitle: { fontSize: 15, fontWeight: 800, color: "#fff" },
  heroTileSubtitle: { fontSize: 11.5, color: "rgba(255,255,255,0.85)", marginTop: 2 },
  hubGroupLabel: { fontSize: 12, fontWeight: 700, color: "var(--muted)", marginBottom: 8 },
  hubGrid: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 8, marginBottom: 18 },
  hubTile: { display: "flex", flexDirection: "column", justifyContent: "space-between", padding: 10, borderRadius: "var(--radius)", minHeight: 84, border: "none", cursor: "pointer", textAlign: "start", position: "relative", boxShadow: "0 4px 10px rgba(0,0,0,0.14)" },
  hubTileLabel: { fontSize: 10.5, fontWeight: 700, color: "#fff", lineHeight: 1.3, marginTop: 8 },
  hubComingSoonBadge: { position: "absolute", top: 6, insetInlineStart: 6, fontSize: 8, fontWeight: 700, padding: "1px 5px", borderRadius: 20, background: "rgba(255,255,255,0.9)", color: "#4B5563" },
  hubListCard: { border: "1px solid var(--border)", borderRadius: "var(--radius)", background: "var(--card)", padding: 6, display: "flex", flexDirection: "column", gap: 1, marginBottom: 14 },
  header: { textAlign: "center", marginBottom: 18, paddingTop: 10 },
  headerLogo: { width: 56, height: 56, borderRadius: 14, objectFit: "contain", marginBottom: 6 },
  routeDots: { display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 10 },
  splashOverlay: { position: "fixed", inset: 0, zIndex: 999, background: "var(--bg)", display: "flex", alignItems: "center", justifyContent: "center" },
  splashInner: { display: "flex", flexDirection: "column", alignItems: "center", gap: 14 },
  splashRing: { position: "relative", width: 88, height: 88, display: "flex", alignItems: "center", justifyContent: "center" },
  splashLogo: { width: 64, height: 64, borderRadius: 16, objectFit: "contain", position: "relative", zIndex: 1 },
  splashText: { fontSize: 13, fontWeight: 600, color: "var(--muted)", margin: 0 },
  dot: { width: 8, height: 8, borderRadius: "50%", background: "var(--accent)", display: "inline-block" },
  routeLine: { width: 28, height: 2, background: "var(--border)", display: "inline-block" },
  title: { fontSize: 22, fontWeight: 700, margin: "0 0 4px" },
  subtitle: { fontSize: 13.5, color: "var(--muted)", margin: 0 },
  dtWidget: { display: "inline-flex", alignItems: "center", gap: 10, marginTop: 12, background: "var(--card)", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: "8px 14px" },
  dtTextCol: { textAlign: "start" },
  dtDigital: { fontSize: 16, fontWeight: 700, fontVariantNumeric: "tabular-nums", color: "var(--text)" },
  dtDateLine: { fontSize: 11, color: "var(--muted)" },
  card: { background: "var(--card)", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: 16, marginBottom: 14 },
  stepLabel: { fontSize: 14, fontWeight: 700, color: "var(--accent)" },
  hint: { fontSize: 12.5, color: "var(--muted)", margin: "0 0 10px" },
  uploadZone: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", border: "1.5px dashed var(--border)", borderRadius: "var(--radius)", padding: "22px 10px", cursor: "pointer", background: "var(--bg)" },
  errorBox: { display: "flex", gap: 6, alignItems: "flex-start", color: "#B3432A", fontSize: 12.5, marginTop: 10, background: "#F7E9E4", borderRadius: 6, padding: 8 },
  detectBox: { marginTop: 10, borderTop: "1px solid var(--border)", paddingTop: 10 },
  fileSummaryRow: { display: "flex", alignItems: "center", gap: 8, background: "var(--bg)", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: "10px 12px" },
  fileSummaryText: { flex: 1, fontSize: 12.5, color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  fileChangeLink: { fontSize: 12, fontWeight: 600, color: "var(--accent)", cursor: "pointer", flexShrink: 0, textDecoration: "underline" },
  warnLine: { display: "flex", gap: 6, alignItems: "flex-start", color: "var(--muted)", fontSize: 11.5, marginTop: 10 },
  prefGroup: { marginBottom: 14 },
  prefTitle: { fontSize: 13.5, fontWeight: 600, marginBottom: 8, color: "var(--text)" },
  radioRow: { display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" },
  radioLabel: { display: "flex", alignItems: "center", gap: 5, fontSize: 13, color: "var(--text)" },
  chipRow: { display: "flex", flexWrap: "wrap", gap: 8 },
  chip: { display: "flex", alignItems: "center", gap: 5, fontSize: 13, padding: "7px 12px", borderRadius: 20, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: "pointer" },
  chipActive: { background: "var(--bg)", color: "var(--text)", borderWidth: 2 },
  prefChipRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 7 },
  prefChip: { display: "flex", alignItems: "center", justifyContent: "center", gap: 5, fontSize: 13, padding: "8px 10px", borderRadius: 12, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: "pointer" },
  prefChipActive: { background: "var(--bg)", color: "var(--text)", borderWidth: 2 },
  chipRankBadge: { width: 16, height: 16, borderRadius: "50%", background: "var(--accent)", color: "#fff", fontSize: 10, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" },
  smallLabel: { fontSize: 12.5, color: "var(--muted)" },
  regionDot: { width: 9, height: 9, borderRadius: "50%", display: "inline-block", flexShrink: 0 },
  priorityPanel: { background: "var(--bg)", borderRadius: "var(--radius)", padding: 12, marginBottom: 14, border: "1px solid var(--border)" },
  resetPriorityBtn: { display: "flex", alignItems: "center", gap: 4, fontSize: 11, padding: "4px 8px", borderRadius: 6, border: "1px solid var(--border)", background: "var(--card)", color: "#B3432A", cursor: "pointer" },
  priorityRow: { display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid var(--border)" },
  priorityRank: { width: 20, height: 20, borderRadius: "50%", color: "#fff", fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  priorityLabel: { flex: 1, fontSize: 13 },
  regionArrows: { display: "flex", gap: 3 },
  dayPickRow: { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 5, padding: "0 0 10px 28px" },
  dayPickHint: { fontSize: 11, color: "var(--muted)", marginInlineEnd: 4 },
  dayPickBtn: { fontSize: 11, fontWeight: 600, padding: "4px 9px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: "pointer" },
  dayPickBtnActive: { background: "#B3432A", color: "#fff", borderColor: "#B3432A" },
  dayPickBlock: { marginTop: 8, padding: "8px 10px", borderRadius: 10, border: "1px solid var(--border)", background: "var(--card)" },
  priorityChipRow: { display: "flex", flexWrap: "wrap", gap: 6, marginTop: 4 },
  priorityChip: { display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12.5, borderRadius: 20, paddingInlineStart: 10, paddingInlineEnd: 6, paddingTop: 4, paddingBottom: 4, background: "var(--card)", border: "1.5px solid var(--border)" },
  priorityChipRank: { width: 16, height: 16, borderRadius: "50%", color: "#fff", fontSize: 9.5, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  priorityChipLabel: { fontSize: 12.5, color: "var(--text)" },
  priorityChipBtn: { width: 16, height: 16, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "var(--muted)", background: "transparent", border: "none", padding: 0 },
  welcomeLine: { fontSize: 13.5, fontWeight: 600, color: "var(--text)", marginTop: 4 },
  myShiftCard: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", background: "var(--bg)", border: "1.5px solid var(--accent)", borderRadius: "var(--radius)", padding: "12px 16px", marginBottom: 14 },
  arrowBtn: { border: "1px solid var(--border)", background: "var(--card)", borderRadius: 5, width: 24, height: 24, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "var(--text)" },
  computeBtn: { width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "var(--accent)", color: "#fff", border: "none", borderRadius: "var(--radius)", padding: "12px 0", fontSize: 14.5, fontWeight: 700, cursor: "pointer", marginTop: 4 },
  smallActionBtn: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, padding: "7px 11px", borderRadius: 7, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: "pointer" },
  dailyLogTh: { border: "1px solid var(--border)", padding: "4px 3px", background: "var(--bg)", fontWeight: 700, whiteSpace: "normal", wordBreak: "normal", textAlign: "center", verticalAlign: "bottom", lineHeight: 1.25, fontSize: 10 },
  dailyLogTd: { border: "1px solid var(--border)", padding: "5px 6px", whiteSpace: "nowrap" },
  grid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 12 },
  topLabel: { fontSize: 12.5, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 10 },
  topGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 12, marginBottom: 6 },
  topCard: { position: "relative", background: "var(--card)", border: "1.5px solid var(--border)", borderRadius: "var(--radius)", padding: 16 },
  topBadge: { position: "absolute", top: -10, insetInlineStart: 14, color: "#fff", fontSize: 11, fontWeight: 700, padding: "3px 10px", borderRadius: 12, display: "flex", alignItems: "center", gap: 4 },
  heroTop: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 },
  heroCrew: { fontSize: 18, fontWeight: 800, display: "flex", alignItems: "center", gap: 6 },
  mineBadge: { display: "inline-flex", alignItems: "center", gap: 3, fontSize: 10.5, fontWeight: 800, color: "#fff", background: "var(--accent)", borderRadius: 999, padding: "2px 8px" },
  scoreBadge: { fontSize: 17, fontWeight: 800, whiteSpace: "nowrap" },
  scoreBadgeSm: { fontSize: 12, fontWeight: 800, whiteSpace: "nowrap" },
  heroMeta: { fontSize: 12.5, color: "var(--muted)", marginBottom: 8 },
  fpRow: { display: "flex", alignItems: "center", gap: 8, marginBottom: 5 },
  fpRank: { width: 16, height: 16, borderRadius: 4, background: "var(--accent)", color: "#fff", fontSize: 9.5, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  fpLabel: { fontSize: 11.5, color: "var(--text)", width: 130, flexShrink: 0 },
  fpBarBg: { flex: 1, height: 6, borderRadius: 3, background: "var(--bg)", overflow: "hidden" },
  fpBarFill: { height: "100%", borderRadius: 3 },
  fpPct: { fontSize: 11, color: "var(--muted)", width: 32, textAlign: "end", flexShrink: 0 },
  fpSquare: { width: 20, height: 20, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9.5, fontWeight: 700, color: "#fff" },
  resultCard: { background: "var(--card)", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: 12 },
  resultCardTop: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 },
  rankBadge: { width: 22, height: 22, borderRadius: "50%", background: "var(--accent)", color: "#fff", fontSize: 11.5, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" },
  resultCrew: { fontSize: 14.5, fontWeight: 700 },
  resultMeta: { fontSize: 11.5, color: "var(--muted)", margin: "2px 0 8px" },
  regionChipsRow: { display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 },
  regionChip: { display: "flex", alignItems: "center", gap: 4, fontSize: 11, background: "var(--bg)", borderRadius: 10, padding: "3px 8px", color: "var(--text)" },
  compareToggle: { display: "flex", alignItems: "center", justifyContent: "center", gap: 5, width: "100%", fontSize: 12, padding: "6px 0", borderRadius: 7, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: "pointer" },
  compareToggleActive: { background: "var(--accent)", borderColor: "var(--accent)", color: "#fff" },
  compareTable: { borderCollapse: "collapse", width: "100%", minWidth: 420 },
  compareLabelCell: { fontSize: 12.5, fontWeight: 600, color: "var(--muted)", padding: "8px 10px", textAlign: "start", background: "var(--bg)", border: "1px solid var(--border)", whiteSpace: "nowrap" },
  compareHeadCell: { fontSize: 13, fontWeight: 700, padding: "8px 10px", border: "1px solid var(--border)", background: "var(--bg)" },
  compareCell: { fontSize: 12.5, padding: "8px 10px", border: "1px solid var(--border)", textAlign: "center" },
  compareRemoveBtn: { border: "none", background: "transparent", cursor: "pointer", color: "#B3432A", display: "flex" },
  numInputWide: { flex: 1, minWidth: 120, padding: "8px 10px", borderRadius: 8, border: "1px solid var(--border)", fontSize: 13, background: "var(--card)", color: "var(--text)" },
  resetBtn: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, width: "100%", background: "transparent", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: "10px 0", fontSize: 13, color: "var(--muted)", cursor: "pointer" },
  footer: { textAlign: "center", fontSize: 11, color: "var(--muted)", marginTop: 18 },
  printTable: { width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "Tahoma, sans-serif" },
  printTh: { border: "1px solid #999", padding: "6px 8px", background: "#eee", textAlign: "center" },
  printTd: { border: "1px solid #999", padding: "6px 8px", textAlign: "center" },
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 },
  modalBox: { background: "var(--card)", borderRadius: "var(--radius)", width: "100%", maxWidth: 480, maxHeight: "82vh", display: "flex", flexDirection: "column", overflow: "hidden", border: "1px solid var(--border)" },
  modalHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 16px", borderBottom: "1px solid var(--border)" },
  modalTitle: { fontSize: 15, fontWeight: 700, color: "var(--text)" },
  modalCloseBtn: { border: "none", background: "transparent", cursor: "pointer", color: "var(--muted)", display: "flex" },
  modalBody: { padding: 16, overflowY: "auto", color: "var(--text)" },
  helpStep: { marginBottom: 14 },
  helpStepTitle: { fontSize: 13.5, fontWeight: 700, marginBottom: 4, color: "var(--accent)" },
  helpStepBody: { fontSize: 13, color: "var(--text)", lineHeight: 1.6 },
  aboutAppName: { fontSize: 20, fontWeight: 800, color: "var(--accent)" },
  aboutVersion: { fontSize: 12, color: "var(--muted)" },
  aboutAuthorBox: { background: "var(--bg)", border: "1px solid var(--border)", borderRadius: "var(--radius)", padding: 12, textAlign: "center", marginTop: 12 },
  aboutAuthorName: { fontSize: 16, fontWeight: 700, color: "var(--text)" },
};
