/*
 * A workbook written by LibreOffice (tests/fixtures/libreoffice-budget.xlsx, converted from the .fods next to it
 * with `soffice --headless --convert-to xlsx`): another producer's XML, not our hand-written fixtures.
 */
import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Calc } from "../src/domain/xlsx/calculation";
import { Edits } from "../src/domain/xlsx/commands";

const bytes = () => new Uint8Array(readFileSync("tests/fixtures/libreoffice-budget.xlsx"));

describe("LibreOffice workbook", () => {
  it("reads formats, merges and both sheets", async () => {
    const wb = Workbook.open(await XlsxWorkbook.load(bytes()));
    const s = wb.sheets[0];
    expect(wb.sheets.map((x) => x.name)).toEqual(["Budget", "Summary"]);
    expect(s.text("B3")).toBe("2024-01-01");
    expect(s.text("D4")).toBe("$452.30");
    expect(s.text("E4")).toBe("13.1%");
    expect(s.merges).toContainEqual({ r1: 0, c1: 0, r2: 0, c2: 4 });
    expect(wb.styles.format(s.cell("A2")!.style)).toMatchObject({ bold: true, color: "#FFFFFF", fill: "#1F4E79" });
  });

  it("calculates every formula to LibreOffice's cached result", async () => {
    const wb = Workbook.open(await XlsxWorkbook.load(bytes()));
    const calc = new Calc(wb);
    for (const s of wb.sheets) {
      for (const c of s.cells.values()) {
        if (c.formula === undefined) continue;
        const v = calc.evaluate(s, c.row, c.col, c.formula);
        if (typeof v === "number") expect(v, c.formula).toBeCloseTo(c.value as number, 9);
        else expect(v, c.formula).toEqual(c.value);
      }
    }
  });

  it("recalculates across sheets after an edit", async () => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes())));
    e.input(e.wb.sheets[0], 4, 3, "2000"); // Utilities goes over budget and becomes the biggest item
    expect(e.wb.sheets[1].text("B3")).toBe("Utilities");
    expect(e.wb.sheets[1].cell("B1")!.value).toBe(3);
  });
});
