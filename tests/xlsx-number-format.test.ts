import { describe, it, expect } from "vitest";
import { formatValue, builtinFormat, isDateFormat, serialToDate, dateToSerial, parseInput } from "../src/domain/xlsx/number-format";

const f = (v: number | string | boolean, code: string, date1904 = false) => formatValue(v, code, date1904).text;

describe("number formats", () => {
  it("shows General numbers like a spreadsheet: up to 11 significant digits, scientific when huge or tiny", () => {
    expect(f(0, "General")).toBe("0");
    expect(f(1234.5, "General")).toBe("1234.5");
    expect(f(1 / 3, "General")).toBe("0.333333333");
    expect(f(123456789012, "General")).toBe("1.23457E+11");
    expect(f(0.1 + 0.2, "General")).toBe("0.3");
    expect(f(-42, "General")).toBe("-42");
    expect(f(1e-10, "General")).toBe("1E-10");
  });

  it("rounds to placeholders, groups thousands and scales", () => {
    expect(f(1234.567, "0.00")).toBe("1234.57");
    expect(f(1234.567, "#,##0")).toBe("1,235");
    expect(f(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(f(0.5, "0%")).toBe("50%");
    expect(f(0.1234, "0.0%")).toBe("12.3%");
    expect(f(5, "000")).toBe("005");
    expect(f(0.5, "#.##")).toBe(".5");
    expect(f(1234567, "#,##0,")).toBe("1,235");
    expect(f(12345, "0.00E+00")).toBe("1.23E+04");
    expect(f(2.5, "0")).toBe("3");
    expect(f(-2.5, "0")).toBe("-3");
  });

  it("uses positive;negative;zero;text sections, literals and colors", () => {
    const code = '"$"#,##0.00;[Red]("$"#,##0.00);"-";@" units"';
    expect(f(1234.5, code)).toBe("$1,234.50");
    expect(formatValue(-5, code).text).toBe("($5.00)");
    expect(formatValue(-5, code).color).toBe("#FF0000");
    expect(f(0, code)).toBe("-");
    expect(f("ten", code)).toBe("ten units");
    expect(f(3, '0" kg"')).toBe("3 kg");
    expect(f(3, "0\\ \\k\\g")).toBe("3 kg");
    expect(f(-3, "0;0")).toBe("3");
    expect(f(1500, "[>=1000]#,##0,\"K\";0")).toBe("2K");
  });

  it("formats dates and times from serials without timezone shifts", () => {
    expect(f(45292, "yyyy-mm-dd")).toBe("2024-01-01");
    expect(f(45292, "m/d/yyyy")).toBe("1/1/2024");
    expect(f(45292.75, "h:mm AM/PM")).toBe("6:00 PM");
    expect(f(45292.5, "hh:mm:ss")).toBe("12:00:00");
    expect(f(45292, "dddd, mmmm d")).toBe("Monday, January 1");
    expect(f(45292, "mmm-yy")).toBe("Jan-24");
    expect(f(1.5, "[h]:mm")).toBe("36:00");
    expect(f(60, "yyyy-mm-dd")).toBe("1900-02-29");
    expect(f(61, "yyyy-mm-dd")).toBe("1900-03-01");
    expect(f(0, "yyyy-mm-dd", true)).toBe("1904-01-01");
    expect(f(45292.999999, "yyyy-mm-dd hh:mm:ss")).toBe("2024-01-02 00:00:00");
  });

  it("knows built-in format ids and which formats are dates", () => {
    expect(builtinFormat(14)).toBe("m/d/yyyy");
    expect(builtinFormat(10)).toBe("0.00%");
    expect(builtinFormat(49)).toBe("@");
    expect(isDateFormat("yyyy-mm-dd")).toBe(true);
    expect(isDateFormat('0.00" m"')).toBe(false);
    expect(isDateFormat("[Red]0")).toBe(false);
  });

  it("converts serials to calendar dates and back", () => {
    expect(serialToDate(45292)).toEqual({ y: 2024, m: 1, d: 1, h: 0, mi: 0, s: 0, ms: 0, dow: 1 });
    expect(dateToSerial(2024, 1, 1)).toBe(45292);
    expect(dateToSerial(1900, 3, 1)).toBe(61);
  });
});

describe("typed input", () => {
  it("reads numbers, percentages, currency, booleans, dates, errors and text like Sheets", () => {
    expect(parseInput("42")).toEqual({ value: 42 });
    expect(parseInput("-1,234.5")).toEqual({ value: -1234.5 });
    expect(parseInput("12%")).toEqual({ value: 0.12, format: "0%" });
    expect(parseInput("$1,200.50")).toEqual({ value: 1200.5, format: '"$"#,##0.00' });
    expect(parseInput("true")).toEqual({ value: true });
    expect(parseInput("2024-01-01")).toEqual({ value: 45292, format: "yyyy-mm-dd" });
    expect(parseInput("1/2/2024")).toEqual({ value: 45293, format: "m/d/yyyy" });
    expect(parseInput("13:30")).toEqual({ value: 13.5 / 24, format: "h:mm" });
    expect(parseInput("#N/A")).toEqual({ value: { error: "#N/A" } });
    expect(parseInput("1e3")).toEqual({ value: 1000 });
  });

  it("keeps text that only looks like a number when asked", () => {
    expect(parseInput("'0042")).toEqual({ value: "0042", quoted: true });
    expect(parseInput("007")).toEqual({ value: 7 });
    expect(parseInput("abc")).toEqual({ value: "abc" });
    expect(parseInput("")).toEqual({ value: null });
    expect(parseInput("1.2.3")).toEqual({ value: "1.2.3" });
  });
});

describe("decimal places", () => {
  it("adds and removes decimals the way the toolbar buttons do", async () => {
    const { adjustDecimals } = await import("../src/domain/xlsx/number-format");
    expect(adjustDecimals("General", 1, 1.5)).toBe("0.00");
    expect(adjustDecimals("General", -1, 1.25)).toBe("0.0");
    expect(adjustDecimals("General", -1, 3)).toBe("0");
    expect(adjustDecimals("0.00", 1)).toBe("0.000");
    expect(adjustDecimals("0.0", -1)).toBe("0");
    expect(adjustDecimals("0", -1)).toBe("0");
    expect(adjustDecimals("#,##0", 1)).toBe("#,##0.0");
    expect(adjustDecimals('"$"#,##0.00;[Red]("$"#,##0.00)', -1)).toBe('"$"#,##0.0;[Red]("$"#,##0.0)');
    expect(adjustDecimals("0%", 1)).toBe("0.0%");
    expect(adjustDecimals("yyyy-mm-dd", 1)).toBe("yyyy-mm-dd");
  });
});
