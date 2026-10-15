vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { addComment } from "../src/domain/xlsx/comments";
import { setNote } from "../src/domain/xlsx/notes";
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

function menuTitles(view: { contentEl: HTMLElement }, name: string): string[] {
  Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!.click();
  return Menu.last!.items.map((i) => i.title).filter((t) => t !== "---");
}
function runMenu(view: { contentEl: HTMLElement }, name: string, ...path: string[]): void {
  menuTitles(view, name);
  let menu = Menu.last!;
  for (const title of path) {
    const item = menu.items.find((i) => i.title === title)!;
    if (item.submenu) menu = item.submenu;
    else item.run!(new MouseEvent("click"));
  }
}

const reopen = async (data: ArrayBuffer) => Workbook.open(await XlsxWorkbook.load(new Uint8Array(data)));

describe("XlsxView", () => {
  it("finds notes and replies in empty cells, opens the matching message and excludes annotations from replace", async () => {
    const { view, file } = await mount();
    const v = view as any;
    const [a, b] = v.wb.sheets;
    setNote(v.wb.pkg, a.ref.path, 'Z99', 'needle note', 'Author');
    const id = addComment(v.edits, b, 'Y98', 'needle root', 'Author');
    addComment(v.edits, b, 'Y98', 'reply needle', 'Reply author', id);
    v.openFind();
    const bar = view.contentEl.querySelector<HTMLElement>('.xlsx-find')!;
    bar.querySelector<HTMLInputElement>('[aria-label="Search all sheets"]')!.click();
    const input = bar.querySelector<HTMLInputElement>('input[type=search]')!;
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('1 of 3');
    expect(bar.querySelector('.dw-find-source')!.textContent).toContain('Data!Z99 · Note');
    const buttons = () => Array.from(bar.querySelectorAll<HTMLButtonElement>('button'));
    expect(buttons().find(b => b.textContent === 'Replace')!.disabled).toBe(true);
    expect(buttons().find(b => b.textContent === 'Replace all')!.disabled).toBe(true);
    for (let i = 0; i < 2; i++) input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(view.contentEl.querySelector('.xlsx-tab.is-active')!.textContent).toBe('Totals');
    expect(bar.querySelector('.dw-find-source')!.textContent).toContain('Reply author');
    expect(document.activeElement).toBe(input);
    buttons().find(b => b.textContent === 'Open comment')!.click();
    expect(document.querySelector('.xlsx-comment-text mark')!.textContent).toBe('needle');
    Array.from(document.querySelectorAll<HTMLButtonElement>('.test-modal button')).find(b => b.textContent === 'Close')!.click();
    bar.querySelector<HTMLInputElement>('[aria-label="Search comments"]')!.click();
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('1 of 1');
    bar.querySelector<HTMLInputElement>('[aria-label="Search cell notes"]')!.click();
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('No results');
    await view.onUnloadFile(file);
  });
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
    bar.querySelector<HTMLInputElement>('[aria-label="Search all sheets"]')!.click();
    input.value = "SUM";
    bar.querySelector<HTMLInputElement>('[aria-label="Also search within formulas"]')!.click();
    expect(view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!.value).toBe("A5");
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(view.contentEl.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Totals");
    expect(view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!.value).toBe("B1");
    expect(document.activeElement).toBe(input);
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(view.contentEl.querySelector(".xlsx-tab.is-active")!.textContent).toBe("Data");
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

  it("paints icon set icons, without the value when the rule hides it", async () => {
    const { ops } = await mountWith(
      `<conditionalFormatting sqref="A1:A2"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`
    );
    const texts = ops().filter((o) => o.kind === "text").map((o) => o.text);
    expect(texts).toContain("↑");
    expect(texts).toContain("↓");
    expect(texts.filter((t) => t === "5")).toHaveLength(1); // the row header only
  });

  it("paints sparklines in their cells", async () => {
    const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", XM = "http://schemas.microsoft.com/office/excel/2006/main";
    const { ops } = await mountWith(
      `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}"><x14:sparklineGroup type="column"><x14:colorSeries rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>Data!A1:A2</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`
    );
    const spark = ops().find((o) => o.kind === "spark") as unknown as { x: number; w: number; s: { values: number[] } };
    expect(spark.s.values).toEqual([5, 1]);
    expect(spark.w).toBeGreaterThan(0);
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

describe('spreadsheet printing', () => {
  it('includes charts outside the used cells in the printed pages', async () => {
    const { view, file } = await mount();
    const v = view as unknown as { edits: import('../src/domain/xlsx/commands').Edits; sheet: import('../src/domain/xlsx/workbook-model').Sheet };
    const { insertChart, moveChart } = await import('../src/domain/xlsx/charts');
    const path = insertChart(v.edits, v.sheet, {r1: 0, c1: 0, r2: 3, c2: 2});
    moveChart(v.edits, path, {x: 50, y: 400, w: 300, h: 200});
    const exporting = view.exportAs('pdf');
    const button = Array.from(document.querySelectorAll<HTMLButtonElement>('.test-modal button')).find(b => b.textContent === 'Export')!;
    expect(button).toBeTruthy(); button.click();
    await exporting;
    const { renderPdf } = await import('../src/presentation/common/pdf');
    const pages = vi.mocked(renderPdf).mock.calls.at(-1)![0];
    expect(pages.some(p => p.querySelector('.xlsx-chart svg'))).toBe(true);
    expect(Math.max(...pages.map(p => parseFloat(p.querySelector<HTMLElement>('.xlsx-print-sheet')!.style.getPropertyValue('--xlsx-print-h'))))).toBeGreaterThanOrEqual(600);
    await view.onUnloadFile(file);
  });
});

describe('spreadsheet comment dialog', () => {
  it('posts a thread and reply, resolves it, and saves both messages', async () => {
    const {view,file,saved}=await mount();
    const v=view as unknown as {commentDialog(row:number,col:number):void;grid:{render():void}};
    v.commentDialog(1,0);
    const modal=document.querySelector<HTMLElement>('.xlsx-comments-modal')!;
    const post=(label:string,text:string) => {
      const input=modal.querySelector<HTMLTextAreaElement>(`[aria-label="${label}"]`)!;
      input.value=text;
      (input.nextElementSibling as HTMLButtonElement).click();
    };
    post('New comment','Check the amount');post('Reply','Confirmed');
    expect(modal.textContent).toContain('Confirmed');
    Array.from(modal.querySelectorAll('button')).find(b=>b.textContent==='Resolve')!.click();
    expect(modal.textContent).toContain('Resolved');
    expect(modal.querySelector('[aria-label="Reply"]')).toBeNull();
    await view.onUnloadFile(file);
    const wb=await reopen(saved[0]);
    const {commentsOf}=await import('../src/domain/xlsx/comments');
    expect(commentsOf(wb.pkg,wb.sheets[0].ref.path).get('A2')![0]).toMatchObject({resolved:true,messages:[{text:'Check the amount'},{text:'Confirmed'}]});
  });
});


it("F6 moves focus between the cell editor and toolbar without committing navigation", async () => {
  const {view, file, input, saved} = await mount();
  const scope = view.scope as unknown as { keys: {mods: string[]; key: string; fn: () => boolean}[] };
  const f6 = scope.keys.find(k => k.key === "F6" && !k.mods.length)!;
  expect(f6).toBeDefined();
  const back = scope.keys.find(k => k.key === "F6" && k.mods.join() === "Shift")!;
  input.focus(); f6.fn();
  expect((view as any).tabsEl.contains(document.activeElement)).toBe(true);
  back.fn(); expect(document.activeElement).toBe(input);
  expect(saved).toEqual([]);
  await view.onUnloadFile(file);
});

describe("XlsxView keys Obsidian hotkeys could take", () => {
  it("keeps Shift+Tab for the grid", async () => {
    const app = new App();
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    const k = (view as any).scope.keys.find((x: any) => x.key === "Tab" && x.mods.join() === "Shift");
    expect(k).toBeTruthy();
  });
});

describe("XlsxView focus after menus", () => {
  it("gives the grid its focus back when the cell menu closes, so Ctrl+Z and typing keep working", async () => {
    const { Menu } = await import("./obsidian-stub");
    const app = new App();
    const bytes = await xlsx();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    const file = new TFile("menu.xlsx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    const v = view as any;
    v.contextMenu(new MouseEvent("contextmenu"), { area: "cell", row: 1, col: 1 });
    (document.activeElement as HTMLElement | null)?.blur();
    Menu.last!.hide();
    await new Promise((r) => setTimeout(r, 0));
    expect(document.activeElement).toBe(v.grid.input);
  });
});

describe("XlsxView data tools", () => {
  const click = (root: ParentNode, text: string) => Array.from(root.querySelectorAll<HTMLButtonElement>("button")).find((b) => b.textContent === text)!.click();

  it("groups selected rows from the row header menu and collapses them with the gutter button", async () => {
    const { Menu } = await import("./obsidian-stub");
    const { view } = await mount();
    const v = view as any;
    v.grid.select({ row: 1, col: 0 }, { row: 2, col: 16383 });
    v.contextMenu(new MouseEvent("contextmenu"), { area: "rowHeader", row: 1, col: 0 });
    Menu.last!.items.find((i) => i.title === "Group rows 2 - 3")!.run!();
    expect(v.sheet.rowLevels.get(1)).toBe(1);
    const [g] = v.grid.outline().row;
    v.grid.host.toggleGroup("row", g);
    expect(v.sheet.hiddenRows.has(1) && v.sheet.hiddenRows.has(2)).toBe(true);
    v.grid.select({ row: 1, col: 0 }, { row: 2, col: 16383 });
    v.contextMenu(new MouseEvent("contextmenu"), { area: "rowHeader", row: 1, col: 0 });
    Menu.last!.items.find((i) => i.title === "Ungroup rows 2 - 3")!.run!();
    expect(v.sheet.rowLevels.size).toBe(0);
  });

  it("removes duplicates from the dialog and reports the counts", async () => {
    const { Notice } = await import("./obsidian-stub");
    const notices = Notice.messages;
    const { view, type } = await mount();
    const v = view as any;
    v.grid.select({ row: 10, col: 0 });
    type("a");
    type("a");
    type("b");
    v.grid.select({ row: 10, col: 0 }, { row: 12, col: 0 });
    v.removeDuplicatesDialog();
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    click(modal, "Remove duplicates");
    expect([v.sheet.cell("A11")?.value, v.sheet.cell("A12")?.value, v.sheet.cell("A13")?.value ?? null]).toEqual(["a", "b", null]);
    expect(notices.at(-1)).toBe("1 duplicate row found and removed. 2 unique rows remain.");
  });

  it("splits text to columns with the detected separator", async () => {
    const { view, type } = await mount();
    const v = view as any;
    v.grid.select({ row: 10, col: 5 });
    type("x;y");
    v.grid.select({ row: 10, col: 5 });
    v.splitTextDialog();
    click(document.querySelector(".test-modal")!, "Split");
    expect([v.sheet.cell("F11")!.value, v.sheet.cell("G11")!.value]).toEqual(["x", "y"]);
  });

  it("adds a named range for the selection", async () => {
    const { view } = await mount();
    const v = view as any;
    v.grid.select({ row: 1, col: 0 }, { row: 3, col: 0 });
    v.namedRangesDialog();
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    const name = modal.querySelector<HTMLInputElement>('input[aria-label="New name"]')!;
    expect(modal.querySelector<HTMLInputElement>('input[aria-label="New range"]')!.value).toBe("Data!A2:A4");
    name.value = "Amounts";
    name.dispatchEvent(new Event("change"));
    click(modal, "Add");
    expect(v.wb.names().map((n: any) => [n.name, n.formula])).toEqual([["Amounts", "Data!$A$2:$A$4"]]);
    expect(modal.querySelector<HTMLInputElement>('input[aria-label="Name"]')!.value).toBe("Amounts");
  });

  it("downloads the sheet as CSV next to the workbook and imports a CSV as a new tab", async () => {
    const { view, file } = await mount();
    const v = view as any;
    const written: [string, string][] = [];
    v.app.vault.create = async (path: string, data: string) => { written.push([path, data]); return new TFile(path); };
    (file as any).path = "folder/book.xlsx";
    const exporting = v.exportAs("csv");
    Array.from(document.querySelectorAll<HTMLButtonElement>(".test-modal button")).find((b) => b.textContent === "Export")!.click();
    await exporting;
    expect(written[0][0]).toBe("folder/book - Data.csv");
    expect(written[0][1].startsWith("Name,Notes,Double\r\n")).toBe(true);
    v.app.vault.read = async () => "a,b\n1,2\n";
    await v.importCsvFile(new TFile("in/prices.csv"));
    expect(v.sheet.name).toBe("prices");
    expect(v.sheet.cell("B2")!.value).toBe(2);
  });
});

describe("XlsxView Copy as Markdown table", () => {
  it("copies the selection as a Markdown table from Edit and the cell menu", async () => {
    const { view } = await mount();
    const written: string[] = [];
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (t: string) => void written.push(t) } });
    (view as any).grid.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    runMenu(view, "Edit", "Copy as Markdown table");
    await Promise.resolve();
    const lines = written[0].trimEnd().split("\n");
    expect(lines).toHaveLength(3);
    expect(lines[1]).toMatch(/^\| -{3}:? \| -{3}:? \|$/);
  });
});

describe("XlsxView save status", () => {
  it("shows Couldn't save with Retry when writing fails, and Saved after a retry works", async () => {
    const { view, type } = await mount();
    const v = view as any;
    let fail = true;
    const write = v.app.vault.modifyBinary;
    v.app.vault.modifyBinary = async (f: unknown, d: ArrayBuffer) => {
      if (fail) throw new Error("read-only file");
      return write(f, d);
    };
    type("5");
    await v.requestSave();
    const status = view.contentEl.querySelector<HTMLElement>(".dw-status")!;
    expect(status.textContent).toContain("Couldn't save");
    fail = false;
    status.querySelector<HTMLButtonElement>("button")!.click();
    await v.autosave.idle();
    expect(status.textContent).toBe("Saved");
  });
});

describe("XlsxView keyboard", () => {
  it("has one toolbar Tab stop and cycles F6 through toolbar, formula bar, grid and sheet tabs", async () => {
    const { view } = await mount();
    const v = view as any;
    const stops = Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-toolbar button, .dw-toolbar select, .dw-toolbar input")).filter((el) => el.tabIndex === 0);
    expect(stops).toHaveLength(1);
    const f6 = v.scope.keys.find((k: any) => k.key === "F6" && !k.mods.length).fn;
    v.grid.focusGrid();
    f6();
    expect(v.tabsEl.contains(document.activeElement)).toBe(true);
    f6();
    expect(view.contentEl.querySelector(".dw-menubar")!.contains(document.activeElement)).toBe(true);
    f6();
    expect(view.contentEl.querySelector(".dw-toolbar")!.contains(document.activeElement)).toBe(true);
    f6();
    expect(document.activeElement).toBe(v.nameBox);
    f6();
    expect(document.activeElement).toBe(v.grid.input);
  });
});

describe("XlsxView menus", () => {
  it("has Google's menus; the data tools live in Data, not the row", async () => {
    const { view, file } = await mount();
    expect(Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent)).toEqual(["File", "Edit", "View", "Insert", "Format", "Data", "Tools", "Help"]);
    expect(menuTitles(view, "File")).toEqual(["Export…"]);
    expect(menuTitles(view, "Insert")).toEqual(["Sheet", "Image", "Chart", "Pivot table", "Function", "Link", "Checkbox", "Comment", "Note", "Dropdown", "Special characters"]);
    expect(menuTitles(view, "Format")).toEqual(["Number", "Text", "Alignment", "Wrap text", "Merge cells", "Conditional formatting", "Clear formatting"]);
    expect(menuTitles(view, "Data")).toEqual(["Sort A → Z", "Sort Z → A", "Create a filter", "Data validation", "Remove duplicates…", "Split text to columns…", "Named ranges…", "Group rows 2 - 2", "Group columns A - A", "Import CSV as new sheet…"]);
    for (const l of ["Data", "Freeze", "Insert pivot table", "Conditional formatting", "Find and replace (Ctrl+F)"]) expect(view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`), l).toBeNull();
    for (const l of ["Format as currency", "Format as percent", "Decrease decimal places", "Increase decimal places", "Remove fill color"]) expect(view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`)!.classList.contains("dw-icon"), l).toBe(true);
    for (const l of ["Undo (Ctrl+Z)", "Number format", "Font", "Bold (Ctrl+B)", "Fill color", "Borders", "Horizontal align", "Wrap text", "Add comment (Ctrl+Alt+M)", "Insert chart", "Create or remove a filter", "Functions"]) expect(view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`), l).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-segmented")).toBeNull();
    await view.onUnloadFile(file);
  });

  it("freezes from View and switches the theme there", async () => {
    const { view, file } = await mount();
    const v = view as any;
    runMenu(view, "View", "Freeze", "1 row");
    expect(v.sheet.frozen.rows).toBe(1);
    runMenu(view, "View", "Dark");
    expect(view.contentEl.classList.contains("xlsx-theme-dark")).toBe(true);
    menuTitles(view, "View");
    expect(Menu.last!.items.find((i) => i.title === "Dark")!.checked).toBe(true);
    await view.onUnloadFile(file);
  });

  it("has Select all, Delete and a Tools menu like the other editors", async () => {
    const { view, file } = await mount();
    const v = view as any;
    expect(menuTitles(view, "Edit")).toEqual(["Undo", "Redo", "Cut", "Copy", "Copy as Markdown table", "Copy link to cells", "Paste", "Paste values only", "Select all", "Delete", "Find and replace"]);
    expect(menuTitles(view, "Tools")).toEqual(["Paint format", "Find and replace"]);
    runMenu(view, "Edit", "Select all");
    expect(v.grid.selection()).toMatchObject({ r1: 0, c1: 0, r2: 1048575, c2: 16383 });
    v.grid.select({ row: 0, col: 0 });
    runMenu(view, "Edit", "Delete");
    expect(v.sheet.cell(0, 0)?.value ?? null).toBeNull();
    await view.onUnloadFile(file);
  });

  it("paints format from Tools › Paint format onto the next range selected, and clears formatting", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 1, col: 0 });
    v.toggle("bold");
    const bold = v.sheet.at(1, 0).style;
    const labels = Array.from(view.contentEl.querySelectorAll(".dw-toolbar button, .dw-toolbar label.dw-color")).map((b) => b.getAttribute("aria-label"));
    expect(labels.slice(labels.indexOf("Text color") + 1, labels.indexOf("Text color") + 3)).toEqual(["Insert link (Ctrl+K)", "Add comment (Ctrl+Alt+M)"]);
    expect(labels[labels.indexOf("Clear formatting") - 1]).toBe("Paint format (double-click to keep painting)");
    runMenu(view, "Tools", "Paint format");
    expect(view.contentEl.querySelector('[aria-label^="Paint format"]')!.classList.contains("is-active")).toBe(true);
    v.grid.select({ row: 5, col: 3 }, { row: 6, col: 3 });
    v.grid.el.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    expect([v.sheet.at(5, 3)?.style, v.sheet.at(6, 3)?.style]).toEqual([bold, bold]);
    expect(view.contentEl.querySelector('[aria-label^="Paint format"]')!.classList.contains("is-active")).toBe(false);
    v.grid.select({ row: 1, col: 0 });
    runMenu(view, "Format", "Clear formatting");
    expect(v.sheet.at(1, 0).style).toBe(0);
    await view.onUnloadFile(file);
  });

  it("links a cell from Insert › Link: drawn as a link, Ctrl+click opens it, an empty address removes it", async () => {
    const { view, file } = await mount();
    const v = view as any;
    const opened = vi.spyOn(window, "open").mockImplementation(() => null);
    v.grid.select({ row: 1, col: 1 });
    const fill = async (value: string) => {
      await new Promise((r) => setTimeout(r, 0));
      const input = document.querySelector<HTMLInputElement>(".test-modal input")!;
      input.value = value;
      input.dispatchEvent(new Event("change"));
      Array.from(document.querySelectorAll<HTMLButtonElement>(".test-modal button")).find((b) => b.textContent === "Save")!.click();
      await new Promise((r) => setTimeout(r, 0));
    };
    runMenu(view, "Insert", "Link");
    await fill("https://example.com/");
    expect(v.links().get("B2")).toBe("https://example.com/");
    expect(v.grid.host.conditional(1, 1)).toMatchObject({ underline: true });
    expect(v.grid.host.openLink(1, 1)).toBe(true);
    expect(opened).toHaveBeenCalledWith("https://example.com/");
    expect(view.contentEl.querySelector('.dw-toolbar [aria-label="Insert link (Ctrl+K)"]')).not.toBeNull();
    runMenu(view, "Insert", "Link");
    await fill("");
    expect(v.links().size).toBe(0);
    await view.onUnloadFile(file);
  });

  it("inserts an image over the cells at the active cell (Insert › Image)", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 3, col: 2 });
    const png = new File([new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0])], "Logo.png", { type: "image/png" });
    await v.insertImageFile(png);
    const pics = v.pictures();
    expect(pics).toHaveLength(1);
    expect(pics[0]).toMatchObject({ name: "Logo", x: v.sheet.colX(2), y: v.sheet.rowY(3) });
    v.undo();
    expect(v.pictures()).toHaveLength(0);
    await view.onUnloadFile(file);
  });

  it("Insert › Special characters types the character into the active cell (or the open editor)", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 0, col: 0 });
    runMenu(view, "Insert", "Special characters");
    document.querySelector<HTMLButtonElement>('.test-modal .dw-char[aria-label="Section sign"]')!.click();
    expect(v.grid.edit).not.toBeNull();
    // Focus stays in the dialog: Escape there closes it rather than cancelling the cell edit, and typing searches.
    expect(document.activeElement?.closest(".test-modal")).not.toBeNull();
    expect(v.grid.input.value.endsWith("§")).toBe(true);
    document.querySelectorAll(".test-modal").forEach((m) => m.remove());
    await view.onUnloadFile(file);
  });

  it("menu bar items show icons, and other menus (right-click) don't change", async () => {
    const { view, file } = await mount();
    for (const [menu, title, icon] of [["Edit", "Copy", "copy"], ["Insert", "Chart", "chart-column"], ["Data", "Create a filter", "filter"], ["Format", "Wrap text", "wrap-text"]]) {
      menuTitles(view, menu);
      expect(Menu.last!.items.find((i) => i.title === title)?.icon, `${menu} › ${title}`).toBe(icon);
    }
    await view.onUnloadFile(file);
  });

  it("Help › Keyboard shortcuts lists the sheet's shortcuts", async () => {
    const { view, file } = await mount();
    runMenu(view, "Help", "Keyboard shortcuts");
    const modal = Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
    const text = modal.textContent ?? "";
    for (const s of ["Fill down", "Ctrl+D", "Insert date", "Edit cell", "F2"]) expect(text).toContain(s);
    modal.remove();
    await view.onUnloadFile(file);
  });

  it("Insert › Checkbox draws checkboxes; Space and a click on the box check / uncheck them", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 2, col: 10 }, { row: 3, col: 10 });
    runMenu(view, "Insert", "Checkbox");
    v.grid.render();
    const boxes = () => [...v.grid.ops.values()].flat().filter((o: any) => o.kind === "checkbox");
    expect(boxes()).toHaveLength(2);
    expect(boxes().every((o: any) => !o.checked)).toBe(true);
    v.grid.input.dispatchEvent(new KeyboardEvent("keydown", { key: " ", bubbles: true, cancelable: true }));
    expect([v.sheet.at(2, 10).value, v.sheet.at(3, 10).value]).toEqual([true, true]);
    v.grid.render();
    expect(boxes().every((o: any) => o.checked)).toBe(true);
    expect(v.grid.edit).toBeNull();
    // A click on row 3's box unchecks just it.
    const x = v.grid.hw + v.sheet.colX(10) + v.sheet.colWidth(10) / 2, y = v.grid.hh + v.sheet.rowY(3) + v.sheet.rowHeight(3) / 2;
    v.grid.el.querySelector(".xlsx-canvas")!.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, clientX: x, clientY: y, button: 0 }));
    window.dispatchEvent(new PointerEvent("pointerup", { clientX: x, clientY: y }));
    expect([v.sheet.at(2, 10).value, v.sheet.at(3, 10).value]).toEqual([true, false]);
    await view.onUnloadFile(file);
  });

  it("pastes a clipboard image (Ctrl+V of a screenshot) as a picture at the active cell", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 2, col: 1 });
    const shot = new File([new Uint8Array([137, 80, 78, 71])], "image", { type: "image/png" });
    const paste = new Event("paste", { bubbles: true, cancelable: true });
    Object.defineProperty(paste, "clipboardData", { value: { getData: () => "", files: [shot] } });
    v.grid.input.dispatchEvent(paste);
    expect(paste.defaultPrevented).toBe(true);
    await vi.waitFor(() => expect(v.pictures()).toHaveLength(1));
    expect(v.pictures()[0]).toMatchObject({ x: v.sheet.colX(1), y: v.sheet.rowY(2) });
    await view.onUnloadFile(file);
  });

  it("picks a picture with a click: it shows selected, drags to move, and Delete removes it", async () => {
    const { view, file } = await mount();
    const v = view as any;
    v.grid.select({ row: 1, col: 1 });
    await v.insertImageFile(new File([new Uint8Array([137, 80, 78, 71])], "Pic.png", { type: "image/png" }));
    v.grid.render();
    const picEl = () => v.grid.el.querySelector(".xlsx-picture") as HTMLElement;
    expect(picEl().querySelector(".xlsx-chart-handle")).not.toBeNull();
    picEl().dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
    window.dispatchEvent(new PointerEvent("pointerup", { clientX: 10, clientY: 10 }));
    v.grid.render();
    expect(picEl().classList.contains("is-selected")).toBe(true);
    v.grid.input.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    expect(v.pictures()).toHaveLength(0);
    await view.onUnloadFile(file);
  });

  it("names the filter item after the sheet's state", async () => {
    const { view, file } = await mount();
    runMenu(view, "Data", "Create a filter");
    expect(menuTitles(view, "Data")).toContain("Remove filter");
    await view.onUnloadFile(file);
  });
});
