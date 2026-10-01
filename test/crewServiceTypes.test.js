import { describe, it, expect } from "vitest";
import { detectServiceType, parseServiceTypeText, serviceTypeOf } from "../src/crewServiceTypes.js";

describe("crew service types (OR / MOR)", () => {
  it("detects the type, MOR winning over the 'On Request' inside 'Mobility On Request'", () => {
    expect(detectServiceType("MOR")).toBe("MOR");
    expect(detectServiceType("Mobility On Request")).toBe("MOR");
    expect(detectServiceType("mobility")).toBe("MOR");
    expect(detectServiceType("OR")).toBe("OR");
    expect(detectServiceType("on-request")).toBe("OR");
    expect(detectServiceType("Newmarket")).toBeNull();
    expect(detectServiceType("for Morning")).toBeNull();
  });

  it("parses lines with crew numbers, ranges, lists and spreadsheet tabs", () => {
    const { updates, errors } = parseServiceTypeText([
      "21 Newmarket MOR",
      "Crew 5 - On Request",
      "30-32 MOR",
      "12\tRichmond Hill\tOR",
      "7, 8 mobility",
      "",
      "Newmarket MOR",
      "40 Newmarket",
    ].join("\n"));
    expect(updates).toEqual({ "21": "MOR", "5": "OR", "30": "MOR", "31": "MOR", "32": "MOR", "12": "OR", "7": "MOR", "8": "MOR" });
    expect(errors).toEqual([
      { line: 7, text: "Newmarket MOR", reason: "noCrew" },
      { line: 8, text: "40 Newmarket", reason: "noType" },
    ]);
  });

  it("a later line wins for the same crew; times are not crew numbers", () => {
    const { updates } = parseServiceTypeText("21 OR\n21 MOR 5:22-9:37");
    expect(updates).toEqual({ "21": "MOR" });
  });

  it("serviceTypeOf ignores unknown values", () => {
    expect(serviceTypeOf({ "21": "MOR", "5": "x" }, 21)).toBe("MOR");
    expect(serviceTypeOf({ "5": "x" }, 5)).toBeNull();
    expect(serviceTypeOf(null, 5)).toBeNull();
  });
});
