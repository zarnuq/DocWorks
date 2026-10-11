import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { blankXlsx } from "../src/infrastructure/blank";
import { xlsx } from "./xlsx-fixtures";

async function entries(bytes: Uint8Array): Promise<Map<string, string>> {
  const z = await JSZip.loadAsync(bytes);
  const out = new Map<string, string>();
  for (const f of Object.values(z.files)) if (!f.dir) out.set(f.name, await f.async("text"));
  return out;
}

describe("XlsxWorkbook package", () => {
  it("finds sheets through relationships in tab order, keeping ids, names and paths apart", async () => {
    const wb = await XlsxWorkbook.load(await xlsx());
    expect(wb.sheets().map((s) => [s.name, s.sheetId, s.path])).toEqual([
      ["Data", 7, "xl/worksheets/sheet1.xml"],
      ["Totals", 8, "xl/worksheets/sheet2.xml"],
    ]);
    expect(wb.stylesPath).toBe("xl/styles.xml");
    expect(wb.stringsPath).toBe("xl/sharedStrings.xml");
  });

  it("rejects files that aren't workbooks", async () => {
    const z = new JSZip();
    z.file("[Content_Types].xml", "<Types/>");
    await expect(XlsxWorkbook.load(await z.generateAsync({ type: "uint8array" }))).rejects.toThrow(/workbook/i);
  });

  it("writes every part back byte-identical when nothing changed", async () => {
    const original = await xlsx();
    const wb = await XlsxWorkbook.load(original);
    expect(await entries(await wb.save())).toEqual(await entries(original));
  });

  it("rewrites only the parts that were touched", async () => {
    const original = await entries(await xlsx());
    const wb = await XlsxWorkbook.load(await xlsx());
    const path = wb.sheets()[1].path;
    wb.xml(path)!.documentElement.setAttribute("data-x", "1");
    wb.touch(path);
    const saved = await entries(await wb.save());
    for (const [name, text] of original) {
      if (name === path) expect(saved.get(name)).toContain('data-x="1"');
      else expect(saved.get(name), name).toBe(text);
    }
  });

  it("makes a blank workbook with one sheet", async () => {
    const wb = await XlsxWorkbook.load(await blankXlsx());
    expect(wb.sheets().map((s) => s.name)).toEqual(["Sheet1"]);
    expect(wb.xml(wb.stylesPath!)).not.toBeNull();
  });
});
