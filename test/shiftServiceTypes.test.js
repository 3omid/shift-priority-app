import { describe, it, expect } from "vitest";
import {
  detectServiceType, parseServiceTypeText, serviceTypeForShift, normalizeRunCode, shiftKey, splitShiftKey,
} from "../src/shiftServiceTypes.js";
import { applyExtraShifts } from "../src/extraShiftParser.js";

describe("shift service types (OR / MOR), per shift not per crew", () => {
  it("detects the type, MOR winning over the 'On Request' inside 'Mobility On Request'", () => {
    expect(detectServiceType("MOR")).toBe("MOR");
    expect(detectServiceType("Mobility On Request")).toBe("MOR");
    expect(detectServiceType("OR")).toBe("OR");
    expect(detectServiceType("on-request")).toBe("OR");
    expect(detectServiceType("Newmarket")).toBeNull();
    expect(detectServiceType("for Morning")).toBeNull();
  });

  it("mixed shifts: OR + MOR together, either order, or a word like mixed/both", () => {
    for (const txt of ["OR+MOR", "MOR + OR", "OR/MOR", "MOR plus On Request", "On Request and Mobility", "mixed", "both", "ترکیبی"]) {
      expect(detectServiceType(txt), txt).toBe("OR+MOR");
    }
    expect(detectServiceType("Mobility On Request")).toBe("MOR"); // one type, not mixed
    const { updates, errors } = parseServiceTypeText("44 Newmarket OR+MOR\nMOR + OR, BRT, 9\n17 mixed");
    expect(updates).toEqual({ "44@NMK": "OR+MOR", "9@RH": "OR+MOR", "17@*": "OR+MOR" });
    expect(errors).toEqual([]);
    expect(serviceTypeForShift(updates, "44", "NMK")).toBe("OR+MOR");
  });

  it("normalizes run codes and builds run@yard keys", () => {
    expect(normalizeRunCode("PRO 9 MRC")).toBe("PRO9MRC");
    expect(normalizeRunCode("run 14")).toBe("14");
    expect(normalizeRunCode("ru15")).toBe("15");
    expect(shiftKey("21", "NMK")).toBe("21@NMK");
    expect(shiftKey("V7", null)).toBe("V7@*");
    expect(shiftKey("V7", "OTHER")).toBe("V7@*");
    expect(splitShiftKey("PRO9MRC@RH")).toEqual({ code: "PRO9MRC", regionKey: "RH" });
  });

  it("parses run + yard + type lines in any order, lists, ranges and spreadsheet tabs", () => {
    const { updates, errors } = parseServiceTypeText([
      "21 Newmarket MOR",
      "Crew 12 Newmarket On Request",
      "PRO 9 MRC, BRT, MOR",
      "30-32 Newmarket MOR",
      "V7\tNewmarket\tOR",
      "23STF mobility",
      "",
      "Newmarket MOR",
      "40 Newmarket",
    ].join("\n"));
    expect(updates).toEqual({
      "21@NMK": "MOR", "12@NMK": "OR", "PRO9MRC@RH": "MOR",
      "30@NMK": "MOR", "31@NMK": "MOR", "32@NMK": "MOR", "V7@NMK": "OR", "23STF@*": "MOR",
    });
    expect(errors).toEqual([
      { line: 8, text: "Newmarket MOR", reason: "noRun" },
      { line: 9, text: "40 Newmarket", reason: "noType" },
    ]);
  });

  it("the same crew can have MOR on one day and OR on another", () => {
    const map = { "21@NMK": "MOR", "12@NMK": "OR", "V7@*": "MOR" };
    const crew41 = { crew: 41, days: [
      { dayIdx: 1, code: "21", regionKey: "NMK" },
      { dayIdx: 4, code: "12", regionKey: "NMK" },
      { dayIdx: 5, code: "21", regionKey: "RH" }, // run 21 at another yard: not set
    ] };
    expect(crew41.days.map((d) => serviceTypeForShift(map, d.code, d.regionKey))).toEqual(["MOR", "OR", null]);
    // A run saved without a yard matches it at any yard; exact yard wins.
    expect(serviceTypeForShift(map, "v 7", "RH")).toBe("MOR");
    expect(serviceTypeForShift({ "V7@*": "MOR", "V7@NMK": "OR" }, "V7", "NMK")).toBe("OR");
    expect(serviceTypeForShift(map, "", "NMK")).toBeNull();
    expect(serviceTypeForShift(null, "21", "NMK")).toBeNull();
  });

  it("works for dispatch extra shifts too (their run code + yard)", () => {
    const out = applyExtraShifts([], [{ id: "s", driverId: "d", date: "2026-10-06", run: "RUN 21", regionKey: "NMK", start: "05:00", end: "09:00" }],
      [{ id: "d", name: "X", employeeId: "1", crewNumber: "41", active: true }], "2026-10-06");
    const day = out[0].days[0];
    expect(serviceTypeForShift({ "21@NMK": "MOR" }, day.code, day.regionKey)).toBe("MOR");
  });
});
