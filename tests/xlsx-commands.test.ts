import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { FormulaSyntaxError } from "../src/domain/xlsx/formula";
import { xlsx, S } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const xmlOf = (el: Element) => new XMLSerializer().serializeToString(el).replace(/ xmlns(:\w+)?="[^"]*"/g, "");
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

describe("typed input", () => {
  it("stores numbers, text and formulas, and recalculates dependents", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 1, 0, "15");
    expect(s.cell("A2")!.value).toBe(15);
    expect(s.cell("A5")!.value).toBe(65); // SUM(A2:A4)
    expect(s.cell("C2")!.value).toBe(30); // shared A2*2
    expect(e.wb.sheets[1].cell("B1")!.value).toBe(65);
    e.input(s, 7, 0, "=a2+1");
    expect(s.cell("A8")).toMatchObject({ formula: "a2+1", value: 16 });
    e.input(s, 7, 1, "=SUM(A2:A4");
    expect(s.cell("B8")).toMatchObject({ formula: "SUM(A2:A4)", value: 65 });
  });

  it("rejects malformed formulas without changing the cell", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    expect(() => e.input(s, 1, 0, "=1+*2")).toThrow(FormulaSyntaxError);
    expect(s.cell("A2")!.value).toBe(10);
  });

  it("applies the number format implied by the input only to General cells", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 8, 0, "25%");
    expect(s.cell("A9")!.value).toBe(0.25);
    expect(s.text("A9")).toBe("25%");
    e.input(s, 4, 1, "2024-02-03");
    expect(s.text("B5")).toBe("2024-02-03");
    e.input(s, 8, 1, "'007");
    expect(s.cell("B9")!.value).toBe("007");
  });

  it("clears contents but keeps formatting", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.clear(s, rect(0, 0, 1, 1));
    expect(xmlOf(s.cell("A1")!.el)).toBe('<c r="A1" s="1"/>');
    expect(s.cell("A2")).toBeUndefined();
    expect(s.cell("A5")!.value).toBe(50);
  });
});

describe("undo and redo", () => {
  it("undoes a typed value, including the recalculated dependents and split shared formulas", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const before = new XMLSerializer().serializeToString(s.doc);
    e.input(s, 2, 2, "7"); // C3 is a shared formula member
    expect(s.cell("C3")!.value).toBe(7);
    expect(e.undo()).toBe(true);
    expect(new XMLSerializer().serializeToString(e.wb.sheets[0].doc)).toBe(before);
    expect(e.wb.sheets[0].cell("C3")).toMatchObject({ formula: "A3*2", value: 40, shared: "0" });
    expect(e.redo()).toBe(true);
    expect(e.wb.sheets[0].cell("C3")!.value).toBe(7);
    expect(e.wb.sheets[0].cell("C2")!.formula).toBe("A2*2");
  });

  it("keeps calculation working after undo", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 1, 0, "100");
    e.undo();
    e.input(s, 2, 0, "1");
    expect(s.cell("A5")!.value).toBe(41);
    expect(s.cell("C3")!.value).toBe(2);
  });

  it("undoes a new cell completely", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 20, 5, "x");
    e.undo();
    expect(s.cell("F21")).toBeUndefined();
    expect(s.doc.querySelector('row[r="21"]')).toBeNull();
    expect(e.undo()).toBe(false);
  });

  it("drops redo after a new edit", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 1, 0, "1");
    e.undo();
    e.input(s, 1, 0, "2");
    expect(e.redo()).toBe(false);
  });
});

describe("formatting ranges", () => {
  it("applies a style change to every cell, reusing one new format per original format", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.format(s, rect(0, 0, 1, 1), { italic: true });
    const styles = ["A1", "B1", "A2", "B2"].map((a) => s.cell(a)!.style);
    expect(styles[0]).toBe(styles[1]);
    expect(styles[2]).toBe(styles[3]);
    expect(styles[0]).not.toBe(styles[2]);
    expect(e.wb.styles.format(styles[0])).toMatchObject({ italic: true, bold: true });
    expect(e.wb.styles.format(styles[2])).toMatchObject({ italic: true, bold: false });
    e.undo();
    expect(s.cell("A2")!.style).toBe(0);
  });

  it("formats empty cells in the selection", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.format(s, rect(10, 0), { fill: "#FF0000" });
    expect(e.wb.styles.format(s.cell("A11")!.style).fill).toBe("#FF0000");
  });

  it("toggles bold based on the first cell", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    expect(e.toggled(s, rect(0, 0, 1, 0), "bold")).toBe(false); // A1 is bold, so the toggle turns it off
    expect(e.toggled(s, rect(1, 0), "bold")).toBe(true);
  });
});

describe("fill", () => {
  it("fills formulas down with their references moved", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.fill(s, rect(4, 0), rect(4, 0, 4, 1)); // A5 =SUM(A2:A4) right into B5
    expect(s.cell("B5")!.formula).toBe("SUM(B2:B4)");
  });

  it("continues number series and repeats text", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 10, 0, "1");
    e.input(s, 11, 0, "3");
    e.input(s, 10, 1, "x");
    e.fill(s, rect(10, 0, 11, 1), rect(10, 0, 14, 1));
    expect([12, 13, 14].map((r) => s.at(r, 0)!.value)).toEqual([5, 7, 9]);
    expect([12, 13, 14].map((r) => s.at(r, 1)?.value ?? null)).toEqual(["x", null, "x"]);
  });

  it("copies a single number instead of counting", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 10, 0, "4");
    e.fill(s, rect(10, 0), rect(10, 0, 12, 0));
    expect([11, 12].map((r) => s.at(r, 0)!.value)).toEqual([4, 4]);
  });
});

describe("saving", () => {
  it("survives save and reopen with values, formulas and styles", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 1, 0, "12");
    e.input(s, 9, 2, "=A2*3");
    e.input(s, 9, 3, "hello");
    e.format(s, rect(9, 3), { bold: true });
    const wb = Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
    const t = wb.sheets[0];
    expect(t.cell("A2")!.value).toBe(12);
    expect(t.cell("C10")).toMatchObject({ formula: "A2*3", value: 36 });
    expect(t.cell("D10")!.value).toBe("hello");
    expect(wb.styles.format(t.cell("D10")!.style).bold).toBe(true);
    expect(t.cell("A5")!.value).toBe(62);
    expect(wb.pkg.workbook.getElementsByTagNameNS(S, "calcPr")[0].getAttribute("fullCalcOnLoad")).toBeNull();
  });
});

describe("restyle (paint format, clear formatting)", () => {
  it("tiles copied styles over a range and clears formatting back to the default, undoably", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.format(s, rect(1, 0), { bold: true });
    e.format(s, rect(2, 0), { italic: true });
    const [bold, italic] = [s.at(1, 0)!.style, s.at(2, 0)!.style];
    // Paint A2:A3 onto D1:D4: bold, italic, bold, italic.
    e.restyle(s, rect(0, 3, 3, 3), (row) => (row % 2 === 0 ? bold : italic));
    expect([0, 1, 2, 3].map((r) => s.at(r, 3)?.style)).toEqual([bold, italic, bold, italic]);
    e.undo();
    expect([0, 1, 2, 3].map((r) => s.at(r, 3)?.style ?? 0)).toEqual([0, 0, 0, 0]);
    e.restyle(s, rect(1, 0, 2, 0), () => 0);
    expect([s.at(1, 0)!.style, s.at(2, 0)!.style]).toEqual([0, 0]);
    expect(s.at(1, 0)!.value).toBe(10);
  });
});
