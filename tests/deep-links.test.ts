vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
import { afterEach, describe, expect, it, vi } from "vitest";
import type { EditorView } from "prosemirror-view";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { fixture } from "./helpers";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => document.body.empty());

async function mount<V extends DocxView | PptxView | XlsxView>(View: new (leaf: RealLeaf, settings: typeof DEFAULT_SETTINGS) => V, bytes: Uint8Array, path: string) {
  const app = new App();
  app.vault.readBinary = async () => bytes.slice().buffer;
  const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
  document.body.appendChild(view.containerEl);
  const file = new TFile(path) as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  return view;
}

describe("links to a place in a file", () => {
  it("[[deck.pptx#3]] opens slide 3", async () => {
    const view = await mount(PptxView, fixture("sample.pptx"), "deck.pptx");
    await view.setEphemeralState({ subpath: "#3" });
    expect((view as unknown as { index: number }).index).toBe(2);
  });

  it("[[book.xlsx#Two!B2:C3]] opens that sheet with the cells selected", async () => {
    const view = await mount(XlsxView, await xlsx({ sheets: [sheetXml(""), sheetXml("")], names: ["One", "Two"] }), "book.xlsx");
    await view.setEphemeralState({ subpath: "#Two!B2:C3" });
    const v = view as unknown as { sheetIndex: number; grid: { selection(): object } };
    expect(v.sheetIndex).toBe(1);
    expect(v.grid.selection()).toEqual({ r1: 1, c1: 1, r2: 2, c2: 2 });
  });

  it("[[report.docx#Heading]] puts the cursor on that heading", async () => {
    const view = await mount(DocxView, fixture("sample.docx"), "report.docx");
    const editor = (view as unknown as { editor: EditorView }).editor;
    const { Selection } = await import("prosemirror-state");
    editor.dispatch(editor.state.tr.setSelection(Selection.atEnd(editor.state.doc)));
    await view.setEphemeralState({ subpath: "#Doc Heading" });
    expect(editor.state.doc.resolve(editor.state.selection.from).parent.textContent).toBe("Doc Heading");
  });

  it("waits for the file when the link arrives before it has loaded", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.pptx").slice().buffer;
    const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
    document.body.appendChild(view.containerEl);
    const file = new TFile("deck.pptx") as unknown as RealTFile;
    view.file = file;
    const loading = view.onLoadFile(file);
    await view.setEphemeralState({ subpath: "#2" });
    await loading;
    expect((view as unknown as { index: number }).index).toBe(1);
  });
});

describe("Copy link to here", () => {
  const clipboard = () => {
    const written: string[] = [];
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (t: string) => void written.push(t) } });
    return written;
  };
  const run = async (view: { contentEl: HTMLElement }, title: string) => {
    const { Menu } = await import("./obsidian-stub");
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((b) => b.textContent === "Edit")!.click();
    Menu.last!.items.find((i) => i.title === title)!.run!(new MouseEvent("click"));
    await Promise.resolve();
  };

  it("links to the current slide", async () => {
    const view = await mount(PptxView, fixture("sample.pptx"), "deck.pptx");
    await view.setEphemeralState({ subpath: "#2" });
    const written = clipboard();
    await run(view, "Copy link to slide");
    expect(written).toEqual(["[[deck.pptx#2]]"]);
  });

  it("links to the selected cells, with the sheet", async () => {
    const view = await mount(XlsxView, await xlsx({ sheets: [sheetXml(""), sheetXml("")], names: ["One", "Two"] }), "book.xlsx");
    await view.setEphemeralState({ subpath: "#Two!B2:C3" });
    const written = clipboard();
    await run(view, "Copy link to cells");
    expect(written).toEqual(["[[book.xlsx#Two!B2:C3]]"]);
  });

  it("links to the heading the cursor is under", async () => {
    const view = await mount(DocxView, fixture("sample.docx"), "report.docx");
    const editor = (view as unknown as { editor: EditorView }).editor;
    const { TextSelection } = await import("prosemirror-state");
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, editor.state.doc.child(0).nodeSize + 2)));
    const written = clipboard();
    await run(view, "Copy link to heading");
    expect(written).toEqual(["[[report.docx#Doc Heading]]"]);
  });
});

describe("buttons at the right of the menu row", () => {
  it("each editor has Comments there; Sheets lists every comment and jumps to its cell", async () => {
    const { addComment } = await import("../src/domain/xlsx/comments");
    const view = await mount(XlsxView, await xlsx({ sheets: [sheetXml(""), sheetXml("")], names: ["One", "Two"] }), "book.xlsx");
    const v = view as unknown as { edits: never; wb: { sheets: unknown[] }; sheetIndex: number; grid: { selection(): object } };
    addComment(v.edits, v.wb.sheets[1] as never, "C4", "check this", "Ada");
    const side = () => Array.from(view.contentEl.querySelectorAll(".dw-menubar-end button")).map((b) => b.getAttribute("aria-label"));
    expect(side()).toContain("Comments");
    view.contentEl.querySelector<HTMLButtonElement>('.dw-menubar-end button[aria-label="Comments"]')!.click();
    const rows = Array.from(document.querySelectorAll(".dw-comment-list-item"));
    expect(rows.map((r) => r.querySelector(".dw-comment-list-where")!.textContent)).toEqual(["Two!C4"]);
    expect(rows[0].textContent).toContain("check this");
    (rows[0] as HTMLElement).click();
    expect(v.sheetIndex).toBe(1);
    expect(v.grid.selection()).toEqual({ r1: 3, c1: 2, r2: 3, c2: 2 });
  });

  it("Docs and Slides have Comments there too, and none in read-only mode", async () => {
    const docx = await mount(DocxView, fixture("sample.docx"), "report.docx");
    expect(docx.contentEl.querySelector('.dw-menubar-end button[aria-label="Comments"]')).not.toBeNull();
    const pptx = await mount(PptxView, fixture("sample.pptx"), "deck.pptx");
    expect(pptx.contentEl.querySelector('.dw-menubar-end button[aria-label="Comments"]')).not.toBeNull();
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").slice().buffer;
    const ro = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, readOnly: true });
    document.body.appendChild(ro.containerEl);
    const file = new TFile("ro.docx") as unknown as RealTFile;
    ro.file = file;
    await ro.onLoadFile(file);
    expect(ro.contentEl.querySelector('.dw-menubar-end button[aria-label="Comments"]')).toBeNull();
  });
});
