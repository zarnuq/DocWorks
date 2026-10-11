import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { copyRange, parseTsv, pasteClip, pasteText } from "../src/domain/xlsx/clipboard";
import { xlsx } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

describe("copy", () => {
  it("gives displayed values as tab-separated text and an HTML table", async () => {
    const e = await setup();
    const clip = copyRange(e.wb.sheets[0], rect(3, 0, 4, 1));
    expect(clip.text).toBe("30\tTRUE\n60\t2024-01-01");
    expect(clip.html).toContain("<table");
    expect(clip.html).toContain(">2024-01-01</td>");
    expect(copyRange(e.wb.sheets[0], rect(0, 0)).html).toContain("font-weight:bold");
  });

  it("quotes text with tabs, newlines or quotes", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 10, 0, 'a\nb "c"');
    expect(copyRange(s, rect(10, 0)).text).toBe('"a\nb ""c"""');
  });
});

describe("paste", () => {
  it("pastes formulas with moved references, values and styles", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const clip = copyRange(s, rect(0, 0, 4, 0));
    const area = pasteClip(e, s, rect(0, 4), clip);
    expect(area).toEqual(rect(0, 4, 4, 4));
    expect(s.cell("E5")!.formula).toBe("SUM(E2:E4)");
    expect(s.cell("E5")!.value).toBe(60);
    expect(s.cell("E1")!.style).toBe(1);
    e.undo();
    expect(s.cell("E1")).toBeUndefined();
  });

  it("repeats a copy to fill a larger selection that's a multiple of it", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const area = pasteClip(e, s, rect(10, 0, 13, 1), copyRange(s, rect(1, 0)));
    expect(area).toEqual(rect(10, 0, 13, 1));
    expect(s.cell("B14")!.value).toBe(10);
  });

  it("moves cells on cut: formulas keep their references and the source is cleared", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const clip = copyRange(s, rect(4, 0), true);
    pasteClip(e, s, rect(9, 3), clip);
    expect(s.cell("D10")).toMatchObject({ formula: "SUM(A2:A4)", value: 60 });
    expect(s.cell("A5")).toBeUndefined();
  });

  it("pastes plain text from other apps, reading numbers and dates", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const area = pasteText(e, s, rect(10, 0), "1\t2024-01-02\n\"x\ty\"\t=A11*2\n");
    expect(area).toEqual(rect(10, 0, 11, 1));
    expect(s.cell("A11")!.value).toBe(1);
    expect(s.text("B11")).toBe("2024-01-02");
    expect(s.cell("A12")!.value).toBe("x\ty");
    expect(s.cell("B12")!.value).toBe(2);
  });

  it("parses tab-separated text with quoted fields", () => {
    expect(parseTsv('a\t"b\nc"\td\n1\t""""\t')).toEqual([["a", "b\nc", "d"], ["1", '"', ""]]);
    expect(parseTsv("single")).toEqual([["single"]]);
  });
});
