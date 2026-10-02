import { describe, it, expect } from "vitest";
import {
  parseDispatchMessage, normalizeTime, shiftHours, validateShiftRow, applyExtraShifts,
  pruneExpiredShifts, weekBounds, FLAGS,
} from "../src/extraShiftParser.js";
import { seedDriversFromCrews, parseBulkDrivers, findDriverForCrew } from "../src/drivers.js";

// The sample messages are from 2020 (Sept 29 2020 was a Tuesday), so most
// tests pin "today" to a 2020 date to check the weekday lines up.
const TODAY = "2020-07-28"; // a Tuesday

describe("parseDispatchMessage: the sample dispatch messages", () => {
  it("two shifts in one message, each with its own date/run/yard", () => {
    const rows = parseDispatchMessage("shift switches for next week, Tuesday Sept.29 Pro8A BRT 5:22-9:37 and Thursday Oct.1 23STF BRT 4:52-15:07", TODAY);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ date: "2020-09-29", run: "PRO8A", location: "Richmond Hill", regionKey: "RH", start: "05:22", end: "09:37", flags: [] });
    expect(rows[1]).toMatchObject({ date: "2020-10-01", run: "23STF", location: "Richmond Hill", start: "04:52", end: "15:07", flags: [] });
  });

  it('"ru15" is a run code and BRT is Richmond Hill', () => {
    const [row] = parseDispatchMessage("would you available for a shift on Friday Sept.25, ru15 BRT 13:52-20:07?", TODAY);
    expect(row).toMatchObject({ date: "2020-09-25", run: "RUN 15", location: "Richmond Hill", start: "13:52", end: "20:07", flags: [] });
  });

  it('"tomorrow Wednesday" is relative to today; a 4-digit start with no end is flagged; a bus number is ignored', () => {
    const rows = parseDispatchMessage("Please confirm you will come to Newmarket tomorrow Wednesday for your run V7 1322 report Bus 25005", TODAY);
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ date: "2020-07-29", run: "V7", location: "Newmarket", regionKey: "NMK", start: "13:22", end: "" });
    expect(rows[0].flags).toEqual(expect.arrayContaining([FLAGS.relativeDate, FLAGS.missingEnd]));
  });

  it('flags an impossible time like 5:62; "run 14" becomes RUN 14; date after the times', () => {
    const [row] = parseDispatchMessage("do run 14 BRT 5:62-15:07 Friday August 28.?", TODAY);
    expect(row).toMatchObject({ date: "2020-08-28", run: "RUN 14", location: "Richmond Hill", start: "5:62", end: "15:07" });
    expect(row.flags).toContain(FLAGS.invalidStart);
    expect(row.flags).not.toContain(FLAGS.invalidEnd);
  });

  it('"13:52 to 22:07" is a range; no run code is fine', () => {
    const [row] = parseDispatchMessage("Gm Thursday July 30 Newmarket 13:52 to 22:07", TODAY);
    expect(row).toMatchObject({ date: "2020-07-30", run: "", location: "Newmarket", start: "13:52", end: "22:07", flags: [] });
  });
});

describe("dates", () => {
  it("no year -> the nearest upcoming date (next year once this year's has passed)", () => {
    const [row] = parseDispatchMessage("Thursday July 30 BRT 5:00-9:00", "2020-08-01");
    expect(row.date).toBe("2021-07-30");
    expect(row.flags).toContain(FLAGS.weekdayMismatch); // July 30 2021 is a Friday
  });

  it("today itself still counts as upcoming", () => {
    const [row] = parseDispatchMessage("Tuesday July 28 BRT 5:00-9:00", TODAY);
    expect(row.date).toBe("2020-07-28");
    expect(row.flags).toEqual([]);
  });

  it("accepts day-before-month, ordinals and an explicit year", () => {
    expect(parseDispatchMessage("Tuesday 29 Sept BRT 5:22-9:37", TODAY)[0].date).toBe("2020-09-29");
    expect(parseDispatchMessage("Sept 29th BRT 5:22-9:37", TODAY)[0].date).toBe("2020-09-29");
    expect(parseDispatchMessage("Sept 29, 2021 BRT 5:22-9:37", TODAY)[0].date).toBe("2021-09-29");
  });

  it("a bare weekday is the next such day and is flagged as relative", () => {
    const [row] = parseDispatchMessage("Friday Newmarket 13:52-20:07", TODAY);
    expect(row.date).toBe("2020-07-31");
    expect(row.flags).toContain(FLAGS.relativeDate);
  });

  it("flags a message with no date at all", () => {
    const [row] = parseDispatchMessage("ru15 BRT 13:52-20:07", TODAY);
    expect(row.date).toBe("");
    expect(row.flags).toContain(FLAGS.missingDate);
  });

  it("a date with no times still gives one row, with the times flagged", () => {
    const [row] = parseDispatchMessage("can you work Friday Sept.25 Pro8A BRT?", TODAY);
    expect(row).toMatchObject({ date: "2020-09-25", run: "PRO8A", start: "", end: "" });
    expect(row.flags).toEqual(expect.arrayContaining([FLAGS.missingStart, FLAGS.missingEnd]));
  });

  it("empty text gives no rows", () => {
    expect(parseDispatchMessage("", TODAY)).toEqual([]);
    expect(parseDispatchMessage("thanks!", TODAY)).toEqual([]);
  });
});

describe("times", () => {
  it("normalizes and validates", () => {
    expect(normalizeTime("5:22")).toMatchObject({ text: "05:22", valid: true });
    expect(normalizeTime("1322")).toMatchObject({ text: "13:22", valid: true });
    expect(normalizeTime("5:62").valid).toBe(false);
    expect(normalizeTime("25:00").valid).toBe(false);
    expect(normalizeTime("").empty).toBe(true);
  });

  it("computes hours, including past midnight", () => {
    expect(shiftHours("05:22", "09:37")).toBe(4.25);
    expect(shiftHours("22:00", "02:00")).toBe(4);
    expect(shiftHours("13:22", "")).toBeNull();
  });
});

describe("validateShiftRow (after hand edits in the preview)", () => {
  it("re-checks edited times and keeps the relative-date warning only while the date is unchanged", () => {
    const [row] = parseDispatchMessage("Newmarket tomorrow Wednesday run V7 1322", TODAY);
    expect(validateShiftRow(row, TODAY)).toEqual(expect.arrayContaining([FLAGS.relativeDate, FLAGS.missingEnd]));
    const fixed = { ...row, end: "21:00" };
    expect(validateShiftRow(fixed, TODAY)).toEqual([FLAGS.relativeDate]);
    expect(validateShiftRow({ ...fixed, date: "2020-07-30" }, TODAY)).toEqual([]);
    expect(validateShiftRow({ ...fixed, start: "5:62" }, TODAY)).toContain(FLAGS.invalidStart);
    expect(validateShiftRow({ ...fixed, date: "2020-07-01" }, TODAY)).toContain(FLAGS.pastDate);
  });
});

describe("applyExtraShifts / pruneExpiredShifts", () => {
  const crews = [{ crew: 40, type: "FT", shiftRaw: "AM", totalHours: 8, workedCount: 1, days: [{ dayIdx: 1, code: "A1", hours: 8, start: "05:00", end: "13:00", regionKey: "NMK" }], driverName: "" }];
  const drivers = [
    { id: "d1", name: "Omid", employeeId: "1599", type: "full", crewNumber: "40", active: true },
    { id: "d2", name: "Pat", employeeId: "2001", type: "part", crewNumber: "", active: true },
  ];
  const shifts = [
    { id: "s1", driverId: "d1", date: "2020-09-29", run: "PRO8A", regionKey: "RH", start: "05:22", end: "09:37" },
    { id: "s2", driverId: "d2", date: "2020-10-01", run: "23STF", regionKey: "RH", start: "04:52", end: "15:07" },
    { id: "s3", driverId: "d1", date: "2020-10-08", run: "X1", regionKey: "RH", start: "05:00", end: "06:00" },
  ];

  it("adds only the shifts in the reference date's Sun–Sat week, without touching the original rows", () => {
    expect(weekBounds("2020-09-30")).toEqual({ start: "2020-09-27", end: "2020-10-03" });
    const out = applyExtraShifts(crews, shifts, drivers, "2020-09-30");
    const mine = out.find((c) => String(c.crew) === "40");
    expect(mine.days.map((d) => d.dayIdx)).toEqual([1, 2]);
    expect(mine.days[1]).toMatchObject({ code: "PRO8A", extra: true, extraDate: "2020-09-29", hours: 4.25 });
    expect(mine.totalHours).toBe(12.25);
    expect(crews[0].days).toHaveLength(1);
    // A part-timer with no crew number gets their own row.
    const pt = out.find((c) => c.extraOnly);
    expect(pt).toMatchObject({ crew: "P2001", driverName: "Pat", employeeId: "2001" });
    expect(pt.days[0].dayIdx).toBe(4);
  });

  it("returns the board unchanged in a week with no extra shifts", () => {
    expect(applyExtraShifts(crews, shifts, drivers, "2020-09-20")).toBe(crews);
  });

  it("drops shifts whose date has passed", () => {
    expect(pruneExpiredShifts(shifts, "2020-10-01").map((s) => s.id)).toEqual(["s2", "s3"]);
  });
});

describe("drivers directory helpers", () => {
  it("seeds one driver per crew from names + employee IDs, manual entries winning over defaults", () => {
    const drivers = seedDriversFromCrews({
      crewNames: { "40": "Omid F.", "2": "" },
      crewEmployeeIds: { "10": "9999" },
      nameDefaults: { "40": "Omid Farhadnia", "10": "Elahe Alamdar" },
      employeeIdDefaults: { "40": "1599", "10": "1593" },
    });
    expect(drivers).toHaveLength(2); // blank crew 2 is an open run, not a driver
    expect(drivers.find((d) => d.crewNumber === "40")).toMatchObject({ name: "Omid F.", employeeId: "1599", type: "full", active: true });
    expect(drivers.find((d) => d.crewNumber === "10")).toMatchObject({ name: "Elahe Alamdar", employeeId: "9999" });
    expect(findDriverForCrew(drivers, 40).name).toBe("Omid F.");
  });

  it('parses bulk "Name, ID" lines and reports bad or duplicate ones', () => {
    const { rows, errors } = parseBulkDrivers("Jane Doe, 1234\n\nJohn Roe,5678\nno id here\nJane Again, 1234", [{ employeeId: "5678" }]);
    expect(rows).toEqual([{ name: "Jane Doe", employeeId: "1234" }]);
    expect(errors.map((e) => e.reason)).toEqual(["duplicate", "format", "duplicate"]);
  });
});
