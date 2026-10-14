import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { insertCheckboxes, isCheckbox, toggleCheckboxes } from "../src/domain/xlsx/checkbox";
import { xlsx } from "./xlsx-fixtures";

const open = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const sheetOf = (e: Edits) => e.wb.sheets[0];

describe("cell checkboxes (Excel's native ones, as Google's Insert › Checkbox)", () => {
  it("formats the range as checkboxes: empty cells become FALSE, values stay; saved with Excel's feature property bags", async () => {
    const e = await open();
    sheetOf(e).setCell(5, 10, { value: true, formula: null });
    insertCheckboxes(e, sheetOf(e), { r1: 4, c1: 10, r2: 6, c2: 10 });
    const s = sheetOf(e);
    for (const row of [4, 5, 6]) expect(isCheckbox(e.wb, s.at(row, 10)!.style), `row ${row}`).toBe(true);
    expect(s.at(4, 10)!.value).toBe(false);
    expect(s.at(5, 10)!.value).toBe(true);
    expect(isCheckbox(e.wb, s.at(4, 11)?.style ?? 0)).toBe(false);
    const zip = await JSZip.loadAsync(await e.wb.pkg.save());
    const bags = await zip.file("xl/featurePropertyBag/featurePropertyBag.xml")!.async("text");
    expect(bags).toContain('<bag type="Checkbox"/>');
    expect(bags).toContain('<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>');
    expect(await zip.file("[Content_Types].xml")!.async("text")).toContain("application/vnd.ms-excel.featurepropertybag+xml");
    expect(await zip.file("xl/_rels/workbook.xml.rels")!.async("text")).toContain("relationships/FeaturePropertyBag");
    expect(await zip.file("xl/styles.xml")!.async("text")).toContain('<xfpb:xfComplement i="0"/>');
    // Reopened, they're still checkboxes.
    const back = Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
    expect(isCheckbox(back, back.sheets[0].at(4, 10)!.style)).toBe(true);
  });

  it("reuses the bags on a second insert, and undo removes the checkboxes", async () => {
    const e = await open();
    insertCheckboxes(e, sheetOf(e), { r1: 0, c1: 10, r2: 0, c2: 10 });
    insertCheckboxes(e, sheetOf(e), { r1: 1, c1: 10, r2: 1, c2: 10 });
    const zip = await JSZip.loadAsync(await e.wb.pkg.save());
    const bags = await zip.file("xl/featurePropertyBag/featurePropertyBag.xml")!.async("text");
    expect(bags.match(/type="Checkbox"/g)).toHaveLength(1);
    e.undo();
    expect(isCheckbox(e.wb, sheetOf(e).at(1, 10)?.style ?? 0)).toBe(false);
    expect(isCheckbox(e.wb, sheetOf(e).at(0, 10)!.style)).toBe(true);
  });

  it("toggles like Sheets: all checked if any is unchecked, else all unchecked", async () => {
    const e = await open();
    insertCheckboxes(e, sheetOf(e), { r1: 0, c1: 10, r2: 1, c2: 10 });
    sheetOf(e).setCell(0, 10, { value: true, formula: null });
    toggleCheckboxes(e, sheetOf(e), { r1: 0, c1: 10, r2: 1, c2: 10 });
    expect([sheetOf(e).at(0, 10)!.value, sheetOf(e).at(1, 10)!.value]).toEqual([true, true]);
    toggleCheckboxes(e, sheetOf(e), { r1: 0, c1: 10, r2: 1, c2: 10 });
    expect([sheetOf(e).at(0, 10)!.value, sheetOf(e).at(1, 10)!.value]).toEqual([false, false]);
    e.undo();
    expect(sheetOf(e).at(1, 10)!.value).toBe(true);
  });
});
