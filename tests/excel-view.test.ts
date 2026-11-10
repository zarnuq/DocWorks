import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS, type DocWorksSettings } from "../src/presentation/common/settings";
import { XL_COMMAND_IDS } from "../src/presentation/xlsx/excel-ui";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

const XL: DocWorksSettings = { ...DEFAULT_SETTINGS, officeXlsx: true, author: "Ada" };

async function openXl(settings: DocWorksSettings = { ...XL }) {
  const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>4</v></c><c r="C1" t="inlineStr"><is><t>x</t></is></c></row>`)] });
  const app = new App();
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const modify = vi.spyOn(app.vault, "modifyBinary");
  const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
  document.body.appendChild(view.containerEl);
  const file = new TFile("a.xlsx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return { view, v: view as any, file, modify, settings };
}

const tabs = (view: XlsxView) => Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-rb-tab")).filter((t) => !t.hidden).map((t) => t.textContent);
const ribbonButton = (view: XlsxView, id: string) => view.contentEl.querySelector<HTMLButtonElement>(`.dw-rb-panel [data-cmd="${id}"]`)!;
const status = (view: XlsxView, id: string) => view.contentEl.querySelector<HTMLElement>(`.dw-statusbar [data-item="${id}"]`)!;

describe("XlsxView Excel-style interface", () => {
  it("opens with the ribbon above the formula bar and the status bar below the sheet tabs", async () => {
    const { view } = await openXl();
    expect(view.contentEl.querySelector(".dw-menubar")).toBeNull();
    expect(view.contentEl.querySelector(".xlsx-toolbar")).toBeNull();
    expect(tabs(view)).toEqual(["Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Help"]);
    const kids = Array.from(view.contentEl.children).map((c) => c.className.split(" ")[0]);
    expect(kids.indexOf("dw-rb-host")).toBeLessThan(kids.indexOf("xlsx-formula-bar"));
    expect(kids.indexOf("xlsx-bottom")).toBeLessThan(kids.indexOf("dw-statusbar"));
  });

  it("provides every command Excel's layout names", async () => {
    const { v } = await openXl();
    const cmds = v.xlCommands() as Record<string, unknown>;
    expect(XL_COMMAND_IDS.filter((id) => !cmds[id])).toEqual([]);
  });

  it("Home › Bold bolds the selected cells and shows as on", async () => {
    const { view, v } = await openXl();
    v.grid.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    ribbonButton(view, "Bold").click();
    expect(v.wb.styles.format(v.sheet.at(0, 0).style).bold).toBe(true);
    expect(ribbonButton(view, "Bold").classList.contains("is-active")).toBe(true);
  });

  it("shows Average, Count and Sum of the selection, and Ready", async () => {
    const { view, v } = await openXl();
    v.grid.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    expect(status(view, "average").textContent).toBe("Average: 3");
    expect(status(view, "count").textContent).toBe("Count: 3");
    expect(status(view, "sum").textContent).toBe("Sum: 6");
    expect(status(view, "mode").textContent).toBe("Ready");
    view.contentEl.querySelector(".dw-statusbar")!.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Cell Mode", "Average", "Count", "Sum"]);
  });

  it("keeps read-only files to the commands that look at them, and never saves", async () => {
    const { view, file, modify } = await openXl({ ...XL, readOnly: true });
    expect(tabs(view)).not.toContain("Insert");
    expect(tabs(view)).not.toContain("Data");
    await view.onUnloadFile(file);
    expect(modify).not.toHaveBeenCalled();
  });

  it("switches layouts on reopen without saving", async () => {
    const { view, file, modify, settings } = await openXl();
    settings.officeXlsx = false;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(view.contentEl.querySelector(".dw-rb")).toBeNull();
    expect(view.contentEl.querySelector(".xlsx-toolbar")).not.toBeNull();
    expect(modify).not.toHaveBeenCalled();
  });
});

describe("XlsxView Excel-style right-click and keys", () => {
  const items = () => Array.from(document.querySelectorAll(".dw-ctx > .menu:not(.dw-ctx-sub) > *")).map((e) => (e.classList.contains("menu-separator") ? "---" : e.querySelector(".menu-item-title")!.textContent));
  const scopeKey = (v: { scope: { keys: { mods: string[]; key: string; fn: () => unknown }[] } }, mods: string, key: string) => v.scope.keys.filter((k) => k.mods.join() === mods && k.key === key);

  it("right-click on a cell shows Excel's menu and mini toolbar", async () => {
    const { v } = await openXl();
    v.contextMenu(new MouseEvent("contextmenu", { clientX: 10, clientY: 10 }), { area: "cell", row: 0, col: 0 });
    expect(items().slice(0, 4)).toEqual(["Cut", "Copy", "Paste", "Paste Values"]);
    expect(items()).toEqual(expect.arrayContaining(["Insert", "Delete", "Clear Contents", "Sort", "New Comment", "New Note", "Link…", "Create Note from Selection"]));
    expect(document.querySelector(".dw-ctx-mini [aria-label^='Bold']")).not.toBeNull();
  });

  it("Ctrl+5 strikes through and Ctrl+Shift+L toggles the filter while the grid has focus", async () => {
    const { v } = await openXl();
    v.grid.focusGrid();
    v.grid.select({ row: 0, col: 0 });
    expect(scopeKey(v, "Mod", "5").map((k) => k.fn())).toContain(false);
    expect(v.wb.styles.format(v.sheet.at(0, 0).style).strike).toBe(true);
    expect(scopeKey(v, "Mod,Shift", "l").map((k) => k.fn())).toContain(false);
    expect(v.sheet.filter).toBeTruthy();
  });

  it("Excel's keys do nothing in the Google layout", async () => {
    const { v } = await openXl({ ...XL, officeXlsx: false });
    v.grid.focusGrid();
    // Not registered at all: Ctrl+5 stays the app's.
    expect(scopeKey(v, "Mod", "5")).toEqual([]);
  });
});

describe("XlsxView Excel-style screenshot fixes", () => {
  it("calls the General number format General", async () => {
    const { view } = await openXl();
    expect(view.contentEl.querySelector('.dw-rb-select[data-cmd="NumberFormatGallery"] .dw-rb-select-label')!.textContent).toBe("General");
  });

  it("shows Average and Sum in the selection's number format", async () => {
    const { view, v } = await openXl();
    v.grid.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    v.format({ numFmt: '"$"#,##0.00' });
    v.grid.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    expect(status(view, "average").textContent).toBe("Average: $3.00");
    expect(status(view, "sum").textContent).toBe("Sum: $6.00");
  });
});

describe("XlsxView Excel-style review fixes", () => {
  const scopeKey = (v: { scope: { keys: { mods: string[]; key: string; fn: () => unknown }[] } }, mods: string, key: string) => v.scope.keys.filter((k) => k.mods.join() === mods && k.key === key);

  it("registers Excel's keys only while its layout is on", async () => {
    const google = await openXl({ ...XL, officeXlsx: false });
    expect(scopeKey(google.v, "Mod", "5")).toEqual([]);
    const { v, view, file } = await openXl();
    expect(scopeKey(v, "Mod", "5").length).toBe(1);
    await view.onUnloadFile(file);
    expect(scopeKey(v, "Mod", "5")).toEqual([]);
  });

  it("reads the selection once per selection change, and not while typing in a cell", async () => {
    const { v } = await openXl();
    const scan = vi.spyOn(v.sheet, "cellsIn");
    v.grid.select({ row: 0, col: 0 }, { row: 0, col: 2 });
    expect(scan).toHaveBeenCalledTimes(1);
    scan.mockClear();
    v.grid.host.editing("1");
    v.grid.host.editing("12");
    expect(scan).not.toHaveBeenCalled();
  });

  it("Google Sheets' strikethrough key is off in Excel mode", async () => {
    const { v } = await openXl();
    v.grid.focusGrid();
    v.grid.select({ row: 0, col: 0 });
    scopeKey(v, "Mod,Shift", "x")[0].fn();
    expect(v.wb.styles.format(v.sheet.at(0, 0).style).strike).toBeFalsy();
  });

  it("offers Ungroup on headers only when something there is grouped, and counts rows", async () => {
    const { v } = await openXl();
    v.grid.select({ row: 0, col: 0 }, { row: 2, col: 16383 });
    v.contextMenu(new MouseEvent("contextmenu", { clientX: 10, clientY: 10 }), { area: "rowHeader", row: 0, col: 0 });
    const titles = () => Array.from(document.querySelectorAll(".dw-ctx .menu-item-title")).map((t) => t.textContent);
    expect(titles()).not.toContain("Ungroup");
    expect(titles()).toContain("Hide 3 Rows");
  });
});
