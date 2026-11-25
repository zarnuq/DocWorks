/* eslint-disable @typescript-eslint/no-explicit-any */
vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { insertRows } from "../src/domain/xlsx/sheet-edit";
import { setProtected } from "../src/domain/xlsx/protection";
import { newRangeName, normalizeSqref, protectedRangeAt, protectedRanges, setProtectedRange } from "../src/domain/xlsx/protected-ranges";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

const ROWS = '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><f>A1*10</f><v>10</v></c></row>';
const PR = '<protectedRanges><protectedRange sqref="A1:A3" name="Totals"/></protectedRanges>';
const book = (after = PR) => xlsx({ sheets: [sheetXml(ROWS, { before: "", after })] });
const edits = async (after = PR) => new Edits(Workbook.open(await XlsxWorkbook.load(await book(after))));

describe("protected ranges (model)", () => {
  it("reads Excel's x:protectedRange list", async () => {
    const e = await edits();
    expect(protectedRanges(e.wb.sheets[0])).toEqual([{ name: "Totals", sqref: "A1:A3" }]);
    expect(protectedRangeAt(e.wb.sheets[0], 2, 0)?.name).toBe("Totals");
    expect(protectedRangeAt(e.wb.sheets[0], 0, 1)).toBeNull();
  });

  it("a protected sheet keeps Excel's own rules: no warnings", async () => {
    const e = await edits();
    setProtected(e, e.wb.sheets[0], true);
    expect(protectedRangeAt(e.wb.sheets[0], 0, 0)).toBeNull();
  });

  it("adds, changes and removes ranges in schema order; the wrapper goes with the last one", async () => {
    const e = await edits("");
    const s = () => e.wb.sheets[0];
    setProtectedRange(e, s(), null, { name: "Inputs", sqref: "B1:B4 D2" });
    expect(protectedRanges(s())).toEqual([{ name: "Inputs", sqref: "B1:B4 D2" }]);
    const xml = new XMLSerializer().serializeToString(s().doc);
    expect(xml.indexOf("<sheetData")).toBeLessThan(xml.indexOf("<protectedRanges"));
    setProtectedRange(e, s(), 0, { name: "Inputs 2", sqref: "C1" });
    expect(protectedRanges(s())).toEqual([{ name: "Inputs 2", sqref: "C1" }]);
    setProtectedRange(e, s(), 0, null);
    expect(new XMLSerializer().serializeToString(s().doc)).not.toContain("protectedRanges");
    e.undo();
    expect(protectedRanges(s())).toEqual([{ name: "Inputs 2", sqref: "C1" }]);
  });

  it("changing a range keeps Excel's other attributes (its password hash)", async () => {
    const e = await edits('<protectedRanges><protectedRange sqref="A1" name="R" password="CC1A"/></protectedRanges>');
    setProtectedRange(e, e.wb.sheets[0], 0, { name: "R", sqref: "A1:A2" });
    expect(new XMLSerializer().serializeToString(e.wb.sheets[0].doc)).toMatch(/<protectedRange[^>]*password="CC1A"/);
  });

  it("moves with inserted rows", async () => {
    const e = await edits();
    insertRows(e, e.wb.sheets[0], 0, 2);
    expect(protectedRanges(e.wb.sheets[0])[0].sqref).toBe("A3:A5");
  });

  it("names and typed ranges", async () => {
    const e = await edits();
    expect(newRangeName(e.wb.sheets[0])).toBe("Range 1");
    expect(normalizeSqref("a1:b2, $C$3")).toBe("A1:B2 C3");
    expect(normalizeSqref("nonsense!")).toBeNull();
    expect(normalizeSqref("")).toBeNull();
  });
});

async function mount(after = PR) {
  const app = new App();
  const bytes = await book(after);
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f, data) => void saved.push(data);
  const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
  document.body.appendChild(view.containerEl);
  const file = new TFile("protected.xlsx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const input = view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-editor")!;
  const type = (row: number, col: number, text: string) => {
    v.grid.select({ row, col });
    input.value = text;
    input.dispatchEvent(new Event("input"));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  };
  const modal = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1);
  const button = (text: string) => Array.from(modal()!.querySelectorAll("button")).find((b) => b.textContent === text)!;
  const flush = () => new Promise((r) => setTimeout(r, 0));
  const value = (row: number, col: number) => v.sheet.at(row, col)?.value;
  return { view, v, type, modal, button, flush, value, saved };
}

describe("Data › Protected sheets and ranges", () => {
  it("editing a protected cell asks first; Cancel leaves it as it was", async () => {
    const { type, modal, button, flush, value } = await mount();
    type(0, 0, "99");
    expect(modal()!.textContent).toContain("Heads up!");
    expect(modal()!.textContent).toContain("Totals");
    expect(value(0, 0)).toBe(1);
    button("Cancel").click();
    await flush();
    expect(value(0, 0)).toBe(1);
  });

  it("the cell editor closes while the warning is up (the typed text doesn't linger for the next cell)", async () => {
    const { v, type, button, flush, value } = await mount();
    type(0, 0, "99");
    expect(v.grid.edit).toBeFalsy();
    button("OK").click();
    await flush();
    expect(value(0, 0)).toBe(99);
    type(0, 1, "5");
    expect(value(0, 1)).toBe(5);
  });

  it("OK makes the edit (and its formulas follow); the next edit asks again", async () => {
    const { type, modal, button, flush, value } = await mount();
    type(0, 0, "7");
    button("OK").click();
    await flush();
    expect(value(0, 0)).toBe(7);
    expect(value(1, 0)).toBe(70);
    type(0, 0, "8");
    expect(modal()!.textContent).toContain("Heads up!");
  });

  it("\"Don't show this again for 5 minutes\" lets the next edits through", async () => {
    const { type, modal, button, flush, value } = await mount();
    type(0, 0, "7");
    const toggle = modal()!.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    toggle.checked = true;
    toggle.dispatchEvent(new Event("change"));
    button("OK").click();
    await flush();
    const before = document.querySelectorAll(".test-modal").length;
    type(2, 0, "5");
    expect(document.querySelectorAll(".test-modal").length).toBe(before);
    expect(value(2, 0)).toBe(5);
  });

  it("cells outside the range and formulas recalculated inside it don't ask", async () => {
    const { type, value } = await mount('<protectedRanges><protectedRange sqref="A2" name="Formula"/></protectedRanges>');
    type(0, 0, "3");
    expect(document.querySelector(".test-modal")).toBeNull();
    expect(value(1, 0)).toBe(30);
  });

  it("formatting a protected cell asks too", async () => {
    const { v, modal } = await mount();
    v.grid.select({ row: 0, col: 0 });
    v.format({ bold: true });
    expect(modal()?.textContent).toContain("Heads up!");
  });

  it("the dialog lists the ranges, adds the selection, and removes one", async () => {
    const { view, v, modal, button } = await mount();
    v.grid.select({ row: 4, col: 1 }, { row: 6, col: 2 });
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Data")!.click();
    Menu.last!.items.find((i) => i.title === "Protected sheets and ranges…")!.run!();
    const inputs = () => Array.from(modal()!.querySelectorAll<HTMLInputElement>("input"));
    expect(inputs().map((i) => i.value)).toEqual(["Totals", "A1:A3", "", "B5:C7"]);
    button("Add").click();
    expect(protectedRanges(v.sheet)).toEqual([{ name: "Totals", sqref: "A1:A3" }, { name: "Range 1", sqref: "B5:C7" }]);
    Array.from(modal()!.querySelectorAll("button")).filter((b) => b.textContent === "Remove")[0].click();
    expect(protectedRanges(v.sheet)).toEqual([{ name: "Range 1", sqref: "B5:C7" }]);
  });
});
