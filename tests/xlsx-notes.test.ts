import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import * as ops from "../src/domain/xlsx/sheet-edit";
import { notesOf, setNote } from "../src/domain/xlsx/notes";
import { xlsx } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const texts = (e: Edits) => Object.fromEntries([...notesOf(e.wb.pkg, e.wb.sheets[0].ref.path)].map(([ref, n]) => [ref, n.text]));

describe("cell notes", () => {
  it("adds a note (Excel's comment part and the drawing Excel shows it with), and saves and reopens it", async () => {
    const e = await setup();
    const sheet = e.wb.sheets[0];
    e.structural(() => setNote(e.wb.pkg, sheet.ref.path, "B2", "Check this figure", "Ada"));
    expect(texts(e)).toEqual({ B2: "Check this figure" });
    expect(notesOf(e.wb.pkg, sheet.ref.path).get("B2")!.author).toBe("Ada");
    const bytes = await e.wb.pkg.save();
    const zip = await JSZip.loadAsync(bytes);
    const vml = Object.keys(zip.files).find((f) => f.endsWith(".vml"))!;
    const drawing = await zip.file(vml)!.async("text");
    expect(drawing).toContain("<x:Row>1</x:Row>");
    expect(drawing).toContain("<x:Column>1</x:Column>");
    const sheetXml = await zip.file(sheet.ref.path)!.async("text");
    expect(sheetXml).toContain("legacyDrawing");
    const back = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    expect(texts(back)).toEqual({ B2: "Check this figure" });
  });

  it("edits and removes notes; removing the last one removes the parts", async () => {
    const e = await setup();
    const path = e.wb.sheets[0].ref.path;
    e.structural(() => setNote(e.wb.pkg, path, "A1", "one", "Ada"));
    e.structural(() => setNote(e.wb.pkg, path, "C3", "two", "Ada"));
    e.structural(() => setNote(e.wb.pkg, path, "A1", "one, edited", "Ada"));
    expect(texts(e)).toEqual({ A1: "one, edited", C3: "two" });
    e.undo();
    expect(texts(e)).toEqual({ A1: "one", C3: "two" });
    e.structural(() => setNote(e.wb.pkg, path, "A1", null, "Ada"));
    e.structural(() => setNote(e.wb.pkg, path, "C3", null, "Ada"));
    expect(texts(e)).toEqual({});
    const zip = await JSZip.loadAsync(await e.wb.pkg.save());
    expect(Object.keys(zip.files).some((f) => f.endsWith(".vml") || /comments\d*\.xml$/.test(f))).toBe(false);
    expect(await zip.file(path)!.async("text")).not.toContain("legacyDrawing");
  });

  it("moves with its cell when rows or columns are inserted or deleted", async () => {
    const e = await setup();
    const path = e.wb.sheets[0].ref.path;
    e.structural(() => setNote(e.wb.pkg, path, "B2", "moves", "Ada"));
    e.structural(() => setNote(e.wb.pkg, path, "B5", "goes", "Ada"));
    ops.insertRows(e, e.wb.sheets[0], 0, 2);
    expect(texts(e)).toEqual({ B4: "moves", B7: "goes" });
    ops.deleteRows(e, e.wb.sheets[0], 6, 1);
    expect(texts(e)).toEqual({ B4: "moves" });
    ops.insertCols(e, e.wb.sheets[0], 0, 1);
    expect(texts(e)).toEqual({ C4: "moves" });
  });
});
