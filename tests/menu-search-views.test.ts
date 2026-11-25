/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS, type DocWorksSettings } from "../src/presentation/common/settings";
import type { MenuEntry } from "../src/presentation/common/menu-search";
import { fixture } from "./helpers";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

type Kind = "docx" | "pptx" | "xlsx";
const KINDS: Kind[] = ["docx", "pptx", "xlsx"];
const GOOGLE: DocWorksSettings = { ...DEFAULT_SETTINGS, officeDocx: false, officePptx: false, officeXlsx: false };

async function bytesOf(kind: Kind): Promise<Uint8Array> {
  if (kind !== "xlsx") return fixture(`sample.${kind}`);
  return xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>4</v></c></row>`)] });
}

async function openView(kind: Kind, settings: DocWorksSettings = { ...GOOGLE }) {
  const bytes = await bytesOf(kind);
  const app = new App();
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const modify = vi.spyOn(app.vault, "modifyBinary");
  const View = { docx: DocxView, pptx: PptxView, xlsx: XlsxView }[kind];
  const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
  document.body.appendChild(view.containerEl);
  const file = new TFile(`search.${kind}`) as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const input = () => view.contentEl.querySelector<HTMLInputElement>(".dw-menu-search input");
  const altSlash = () => (v.scope.keys as { mods: string[]; key: string; fn: () => boolean }[]).find((k) => k.key === "/" && k.mods.join() === "Alt");
  const type = (text: string) => {
    const el = input()!;
    el.focus();
    el.value = text;
    el.dispatchEvent(new Event("input"));
  };
  const shown = () => Array.from(document.querySelectorAll(".dw-menu-search-results .suggestion-item")).map((el) => `${el.querySelector(".dw-menu-search-path")!.textContent} › ${el.querySelector(".dw-menu-search-title")!.textContent}`);
  const enter = () => input()!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  const entries = (): MenuEntry[] => v.toolbar.search.collect();
  return { view, v, file, modify, input, altSlash, type, shown, enter, entries };
}

/** Every item of a stub menu that runs something, with its path, as the menu itself shows them. */
function leaves(menu: Menu, path: string[], out: string[] = []): string[] {
  for (const item of menu.items) {
    if (item.title === "---") continue;
    if (item.submenu) leaves(item.submenu, [...path, item.title], out);
    else if (item.title && item.run) out.push([...path, item.title].join(" › "));
  }
  return out;
}

describe.each(KINDS)("menu search in the %s editor (Google layout)", (kind) => {
  it("puts the 'Menus' box first in the toolbar row", async () => {
    const { view, input } = await openView(kind);
    const bar = view.contentEl.querySelector(".dw-toolbar")!;
    expect(bar.firstElementChild!.classList.contains("dw-menu-search")).toBe(true);
    expect(input()!.placeholder).toBe("Menus");
  });

  it("Alt+/ focuses the box and is handled (no default)", async () => {
    const { altSlash, input } = await openView(kind);
    const key = altSlash();
    expect(key).toBeDefined();
    expect(key!.fn()).toBe(false);
    expect(document.activeElement).toBe(input());
  });

  it("finds every item the menu row's menus show, with the same paths", async () => {
    const { view, entries } = await openView(kind);
    const found = new Set(entries().map((e) => [...e.path, e.title].join(" › ")));
    const names = Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menubar .dw-menu-btn")).map((b) => b.textContent!);
    expect(names.length).toBeGreaterThan(4);
    let total = 0;
    for (const b of Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menubar .dw-menu-btn"))) {
      Menu.last = null;
      b.click();
      // File may open the view's own ⋮ menu instead (none in the stub).
      if (!Menu.last) continue;
      for (const path of leaves(Menu.last, [b.textContent!])) {
        expect(found.has(path), path).toBe(true);
        total++;
      }
    }
    expect(total).toBeGreaterThan(30);
  });

  it("lists Help › Keyboard shortcuts by name and Enter opens it (listing the box's own Alt+/)", async () => {
    const { type, shown, enter } = await openView(kind);
    type("keyboard shortcuts");
    expect(shown()[0]).toBe("Help › Keyboard shortcuts");
    enter();
    const modal = Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
    expect(modal.textContent).toContain("Search the menus");
    expect(modal.textContent).toContain("Alt+/");
  });

  it("finds items through their menu path (View › Zoom levels)", async () => {
    const { type, shown } = await openView(kind);
    type("zoom 150");
    expect(shown()).toContain("View › Zoom › 150%");
  });

  it("reading the menus to search them never changes or saves the file", async () => {
    const { v, type, modify, entries } = await openView(kind);
    for (const q of ["a", "e", "insert", "format", "zoom", "x"]) type(q);
    entries();
    await new Promise((r) => setTimeout(r, 20));
    expect(v.autosave.dirty).toBe(false);
    expect(modify).not.toHaveBeenCalled();
  });

  it("read-only files: finds what their menus keep, not what edits", async () => {
    const { entries, input } = await openView(kind, { ...GOOGLE, readOnly: true });
    expect(input()).not.toBeNull();
    const titles = entries().map((e) => e.title);
    expect(titles).toContain("Keyboard shortcuts");
    expect(titles).not.toContain("Undo");
    expect(titles).not.toContain("Paste");
  });
});

describe.each(KINDS)("menu search in the %s editor (Office layout)", (kind) => {
  it("has no box, and Alt+/ passes on", async () => {
    const settings = { ...GOOGLE, officeDocx: true, officePptx: true, officeXlsx: true };
    const { input, altSlash } = await openView(kind, settings);
    expect(input()).toBeNull();
    expect(altSlash()!.fn()).toBe(true);
  });
});

describe("menu search runs real commands", () => {
  it("Slides: 'new slide' adds a slide", async () => {
    const { v, type, shown, enter } = await openView("pptx");
    const before = v.doc.slides().length;
    type("new slide");
    expect(shown()[0]).toMatch(/› New slide$/);
    enter();
    expect(v.doc.slides().length).toBe(before + 1);
  });

  it("Docs: 'page break' finds Insert › Break › Page break and inserts one", async () => {
    const { v, type, shown, enter } = await openView("docx");
    const breaks = () => {
      let n = 0;
      v.editor.state.doc.descendants((node: any) => void (node.type.name === "page_break" && n++));
      return n;
    };
    const before = breaks();
    type("page break");
    expect(shown()[0]).toBe("Insert › Break › Page break");
    enter();
    expect(breaks()).toBe(before + 1);
  });

  it("Sheets: 'sheet' finds Insert › Sheet first and adds a tab", async () => {
    const { v, type, shown, enter } = await openView("xlsx");
    const before = v.wb.sheets.length;
    type("sheet");
    const first = shown()[0];
    expect(first).toBe("Insert › Sheet");
    enter();
    expect(v.wb.sheets.length).toBe(before + 1);
  });

  it("switching to the Office layout and back keeps Alt+/ working on the new toolbar", async () => {
    const settings = { ...GOOGLE };
    const { view, file, altSlash, input } = await openView("docx", settings);
    settings.officeDocx = true;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(altSlash()!.fn()).toBe(true);
    settings.officeDocx = false;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(altSlash()!.fn()).toBe(false);
    expect(document.activeElement).toBe(input());
  });
});
