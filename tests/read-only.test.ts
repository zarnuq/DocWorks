vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
import { afterEach, describe, expect, it, vi } from "vitest";
import type { EditorView } from "prosemirror-view";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { fixture } from "./helpers";
import { xlsx } from "./xlsx-fixtures";

const READ_ONLY = { ...DEFAULT_SETTINGS, readOnly: true };

afterEach(() => {
  document.body.empty();
  vi.useRealTimers();
  vi.restoreAllMocks();
});

type AnyView = DocxView | PptxView | XlsxView;

async function mount<V extends AnyView>(View: new (leaf: RealLeaf, settings: typeof READ_ONLY) => V, bytes: Uint8Array, path: string) {
  const app = new App();
  app.vault.readBinary = async () => bytes.slice().buffer;
  const saved = vi.spyOn(app.vault, "modifyBinary");
  const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf, READ_ONLY);
  document.body.appendChild(view.containerEl);
  const file = new TFile(path) as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  return { view, saved };
}

const menuNames = (view: AnyView) => Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent);
function menuTitles(view: AnyView, name: string, sub?: string): string[] {
  Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!.click();
  let menu = Menu.last!;
  if (sub) menu = menu.items.find((i) => i.title === sub)!.submenu!;
  return menu.items.map((i) => i.title);
}
const toolbarLabels = (view: AnyView) => Array.from(view.contentEl.querySelectorAll(".dw-toolbar [aria-label]")).map((b) => b.getAttribute("aria-label"));

describe("read-only setting", () => {
  it("opens a document as a viewer: view menus only, Zoom and Page view, no edits, never saved", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const { view, saved } = await mount(DocxView, fixture("sample.docx"), "doc.docx");
    expect(view.contentEl.hasClass("dw-read-only")).toBe(true);
    expect(menuNames(view)).toEqual(["File", "Edit", "View", "Tools", "Help"]);
    expect(menuTitles(view, "Edit")).toEqual(["Copy", "Select all", "Copy link to heading", "---", "Find and replace"]);
    expect(menuTitles(view, "View")).toEqual(["Page view", "Outline", "Linked from", "Zoom", "---", "Light", "Dark"]);
    expect(menuTitles(view, "View", "Zoom").length).toBeGreaterThan(3);
    expect(menuTitles(view, "Tools")).toEqual(["Word count", "Find and replace"]);
    const labels = toolbarLabels(view);
    expect(labels).toContain("Zoom");
    expect(labels).toContain("Page view");
    expect(labels).not.toContain("Bold (Ctrl+B)");
    expect(labels).not.toContain("Undo (Ctrl+Z)");
    expect(labels).not.toContain("Editing mode");
    expect(view.contentEl.querySelector(".dw-end")!.textContent).toContain("View only");
    expect(view.contentEl.querySelector<HTMLElement>('.dw-end [role="status"]')!.style.display).toBe("none");
    expect(view.contentEl.querySelector<HTMLElement>(".docx-ruler-viewport")!.style.display).toBe("none");
    const editor = (view as unknown as { editor: EditorView }).editor;
    expect(editor.editable).toBe(false);
    const before = editor.state.doc;
    editor.dispatch(editor.state.tr.insertText("x", 1));
    expect(editor.state.doc.eq(before)).toBe(true);
    vi.runAllTimers();
    await view.onUnloadFile(view.file!);
    expect(saved).not.toHaveBeenCalled();
  });

  it("Page view is a toolbar button while editing too", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").slice().buffer;
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
    document.body.appendChild(view.containerEl);
    const file = new TFile("doc.docx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    const button = view.contentEl.querySelector<HTMLButtonElement>('.dw-toolbar button[aria-label="Page view"]')!;
    const paged = () => view.contentEl.querySelector(".is-paged") !== null;
    const was = paged();
    expect(button.hasClass("is-active")).toBe(was);
    button.click();
    expect(paged()).toBe(!was);
    expect(button.hasClass("is-active")).toBe(!was);
  });

  it("opens a deck as a viewer: slides can be watched, not changed", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const { view, saved } = await mount(PptxView, fixture("sample.pptx"), "deck.pptx");
    expect(menuNames(view)).toEqual(["File", "Edit", "View", "Tools", "Help"]);
    expect(menuTitles(view, "View")).toEqual(["Present", "Present from start", "---", "Grid view", "Linked from", "Zoom", "---", "Light", "Dark"]);
    const labels = toolbarLabels(view);
    expect(labels).toContain("Zoom");
    expect(labels).not.toContain("Insert text box");
    expect(view.contentEl.querySelector('.pptx-slide-host [contenteditable="true"]')).toBeNull();
    expect(view.contentEl.querySelector<HTMLTextAreaElement>(".pptx-notes")!.readOnly).toBe(true);
    expect(view.contentEl.querySelector(".pptx-thumb[draggable='true'], .pptx-thumb-add")).toBeNull();
    // Select all and Delete do nothing.
    const stage = view.contentEl.querySelector<HTMLElement>(".pptx-stage")!;
    const shapes = view.contentEl.querySelectorAll(".pptx-slide-host .pptx-shape").length;
    (view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => void }[] }).keys.filter((k) => k.mods.join() === "Mod" && k.key === "a").forEach((k) => k.fn());
    stage.focus();
    stage.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    expect(view.contentEl.querySelectorAll(".pptx-slide-host .pptx-shape").length).toBe(shapes);
    vi.runAllTimers();
    await view.onUnloadFile(view.file!);
    expect(saved).not.toHaveBeenCalled();
  });

  it("opens a workbook as a viewer: cells can be selected and copied, not typed into", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const { view, saved } = await mount(XlsxView, await xlsx(), "book.xlsx");
    expect(menuNames(view)).toEqual(["File", "Edit", "View", "Tools", "Help"]);
    expect(menuTitles(view, "View")).toEqual(["Zoom", "Linked from", "---", "Light", "Dark"]);
    expect(toolbarLabels(view)).toContain("Zoom");
    expect(toolbarLabels(view)).not.toContain("Bold (Ctrl+B)");
    const input = view.contentEl.querySelector<HTMLTextAreaElement>(".xlsx-editor")!;
    const cell = () => (view as unknown as { sheet: { at(r: number, c: number): { value: unknown } | undefined } }).sheet.at(0, 0)?.value;
    const before = cell();
    input.value = "changed";
    input.dispatchEvent(new Event("input"));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));
    expect(cell()).toBe(before);
    vi.runAllTimers();
    await view.onUnloadFile(view.file!);
    expect(saved).not.toHaveBeenCalled();
  });
});

describe("opening a file twice at once", () => {
  for (const [name, View, bytes, path] of [
    ["document", DocxView, () => fixture("sample.docx"), "doc.docx"],
    ["deck", PptxView, () => fixture("sample.pptx"), "deck.pptx"],
    ["workbook", XlsxView, xlsx, "book.xlsx"],
  ] as const) {
    it(`draws the ${name} once`, async () => {
      const data = await bytes();
      const app = new App();
      app.vault.readBinary = async () => data.slice().buffer;
      const view = new (View as new (leaf: RealLeaf, settings: typeof DEFAULT_SETTINGS) => AnyView)(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
      document.body.appendChild(view.containerEl);
      const file = new TFile(path) as unknown as RealTFile;
      view.file = file;
      await Promise.all([view.onLoadFile(file), view.onLoadFile(file)]);
      expect(view.contentEl.querySelectorAll(".dw-menubar").length).toBe(1);
    });
  }
});
