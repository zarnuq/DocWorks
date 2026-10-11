import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { xlsx } from "./xlsx-fixtures";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

async function mount() {
  const app = new App();
  const bytes = await xlsx();
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f, data) => {
    saved.push(data);
  };
  const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
  document.body.appendChild(view.containerEl);
  const file = new TFile("book.xlsx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const input = view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-editor")!;
  const type = (text: string) => {
    input.value = text;
    input.dispatchEvent(new Event("input"));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  };
  return { view, file, saved, input, type };
}

const reopen = async (data: ArrayBuffer) => Workbook.open(await XlsxWorkbook.load(new Uint8Array(data)));

describe("XlsxView", () => {
  it("opens a workbook: cells, sheet tabs and the formula bar for the first cell below frozen rows", async () => {
    const { view } = await mount();
    const el = view.contentEl;
    expect(Array.from(el.querySelectorAll(".xlsx-tab")).map((t) => t.textContent)).toEqual(["Data", "Totals"]);
    expect(el.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Data");
    const grid = (view as unknown as { grid: { ops: Map<string, { kind: string; text?: string }[]> } }).grid;
    expect([...grid.ops.values()].flat().map((o) => o.text)).toContain("Name");
    expect(el.querySelector<HTMLInputElement>(".xlsx-name-box")!.value).toBe("A2");
    expect(el.querySelector<HTMLTextAreaElement>(".xlsx-formula-input")!.value).toBe("10");
  });

  it("doesn't save when the file is only opened, scrolled and clicked around", async () => {
    const { view, file, saved, input } = await mount();
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    view.contentEl.querySelectorAll<HTMLElement>(".xlsx-tab")[1].click();
    await view.onUnloadFile(file);
    expect(saved).toHaveLength(0);
  });

  it("saves typed values and recalculated formulas when closed", async () => {
    const { view, file, saved, type } = await mount();
    type("100");
    await view.onUnloadFile(file);
    expect(saved).toHaveLength(1);
    const wb = await reopen(saved[0]);
    expect(wb.sheets[0].cell("A2")!.value).toBe(100);
    expect(wb.sheets[0].cell("A5")!.value).toBe(150);
    expect(wb.sheets[1].cell("B1")!.value).toBe(150);
  });

  it("shows formula errors without changing the cell", async () => {
    const { view, file, saved, type } = await mount();
    type("=1+*2");
    expect(view.contentEl.querySelector(".xlsx-editor.is-editing")).not.toBeNull();
    await view.onUnloadFile(file);
    expect(saved).toHaveLength(0);
  });

  it("goes to a cell or another sheet from the name box", async () => {
    const { view } = await mount();
    const box = view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!;
    box.value = "Totals!B1";
    box.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(view.contentEl.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Totals");
    expect(view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-formula-input")!.value).toBe("=SUM(Data!A2:A4)");
  });

  it("adds a sheet from the + button and saves it", async () => {
    const { view, file, saved } = await mount();
    view.contentEl.querySelector<HTMLElement>(".xlsx-tab-add")!.click();
    expect(Array.from(view.contentEl.querySelectorAll(".xlsx-tab")).map((t) => t.textContent)).toEqual(["Data", "Sheet3", "Totals"]);
    expect(view.contentEl.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Sheet3");
    await view.onUnloadFile(file);
    expect((await reopen(saved[0])).sheets.map((s) => s.name)).toEqual(["Data", "Sheet3", "Totals"]);
  });

  it("undoes with Ctrl+Z through the view's scope", async () => {
    const { view, type } = await mount();
    type("5");
    const scope = view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => boolean }[] };
    view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-editor")!.focus();
    scope.keys.find((k) => k.key === "z" && k.mods.length === 1)!.fn();
    const box = view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!;
    box.value = "A2";
    box.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-formula-input")!.value).toBe("10");
  });

  it("finds cells across sheets", async () => {
    const { view } = await mount();
    const scope = view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => boolean }[] };
    scope.keys.find((k) => k.key === "f")!.fn();
    const bar = view.contentEl.querySelector(".xlsx-find")!;
    const input = bar.querySelector<HTMLInputElement>("input[type=search]")!;
    bar.querySelectorAll<HTMLInputElement>("input[type=checkbox]")[3].click(); // all sheets
    input.value = "SUM";
    bar.querySelectorAll<HTMLInputElement>("input[type=checkbox]")[2].click(); // formulas
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(view.contentEl.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Totals");
    expect(view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!.value).toBe("B1");
  });
});

describe("XlsxView lifecycle", () => {
  it("listens for outside changes once per view, not once per opened file", async () => {
    const app = new App();
    const bytes = await xlsx();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const on = vi.spyOn(app.vault, "on");
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    for (const path of ["a.xlsx", "b.xlsx"]) {
      const file = new TFile(path) as unknown as RealTFile;
      view.file = file;
      await view.onLoadFile(file);
      await view.onUnloadFile(file);
    }
    expect(on.mock.calls.filter(([name]) => name === "modify")).toHaveLength(1);
  });
});

describe("XlsxView conditional formatting and dropdowns", () => {
  const S_ = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  async function mountWith(after: string) {
    const { sheetXml } = await import("./xlsx-fixtures");
    const app = new App();
    const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>5</v></c><c r="B1" t="inlineStr"><is><t>Yes</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>`, { before: "", after })], names: ["Data"] });
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const saved: ArrayBuffer[] = [];
    app.vault.modifyBinary = async (_f, data) => void saved.push(data);
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    const file = new TFile("rules.xlsx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    const grid = (view as any).grid;
    const ops = () => [...grid.ops.values()].flat() as { kind: string; color?: string; text?: string }[];
    return { view, file, saved, grid, ops };
  }

  it("paints conditional fills, shows dropdown buttons and refuses items not on the list", async () => {
    const { view, grid, ops } = await mountWith(
      `<conditionalFormatting sqref="A1:A2"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF57BB8A"/></colorScale></cfRule></conditionalFormatting>` +
        `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B1:B5"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>`
    );
    expect(ops().some((o) => o.kind === "rect" && o.color === "#57BB8A")).toBe(true);
    grid.render();
    expect(view.contentEl.querySelectorAll(".xlsx-dropdown-btn").length).toBeGreaterThanOrEqual(5);
    const v = view as any;
    expect(v.commit(1, 1, "Maybe")).toBe(false);
    expect(v.sheet.at(1, 1)).toBeUndefined();
    expect(v.commit(1, 1, "no")).toBe(true);
    expect(v.sheet.at(1, 1).value).toBe("no");
  });

  it("adds a conditional format rule and a dropdown from the dialogs, and saves them", async () => {
    const { view, file, saved, ops } = await mountWith("");
    const v = view as any;
    v.grid.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    v.conditionalDialog();
    let modal = document.querySelector(".xlsx-rules-modal") as HTMLElement;
    const set = (label: string, value: string) => {
      const el = modal.querySelector<HTMLInputElement | HTMLSelectElement>(`[aria-label="${label}"]`)!;
      el.value = value;
      el.dispatchEvent(new Event("change"));
    };
    set("Condition", "cellIs:greaterThan");
    set("Value or formula", "3");
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Add rule")!.click();
    expect(modal.textContent).toContain("Greater than 3");
    v.grid.render();
    expect(ops().some((o) => o.kind === "rect" && o.color === "#B7E1CD")).toBe(true);
    modal.remove();
    v.grid.select({ row: 0, col: 2 }, { row: 3, col: 2 });
    v.validationDialog();
    modal = document.querySelector(".xlsx-rules-modal") as HTMLElement;
    set("Items", "Low, High");
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Done")!.click();
    expect(v.sheet.doc.getElementsByTagNameNS(S_, "dataValidation")[0].getAttribute("sqref")).toBe("C1:C4");
    await view.onUnloadFile(file);
    const back = await reopen(saved.at(-1)!);
    expect(back.sheets[0].doc.getElementsByTagNameNS(S_, "cfRule")).toHaveLength(1);
    expect(back.sheets[0].doc.getElementsByTagNameNS(S_, "formula1")[1]?.textContent ?? back.sheets[0].doc.getElementsByTagNameNS(S_, "formula1")[0].textContent).toContain("Low,High");
  });
});
