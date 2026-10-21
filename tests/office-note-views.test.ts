import { afterEach, expect, it, vi } from "vitest";
import type { WorkspaceLeaf as RealLeaf, TFile as RealFile } from "obsidian";
import { TextSelection } from "prosemirror-state";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { fixture } from "./helpers";
import { xlsx } from "./xlsx-fixtures";
import type { RenderedShape } from "../src/presentation/pptx/slide-renderer";

afterEach(() => { document.body.empty(); vi.restoreAllMocks(); });

async function mount(kind: "docx" | "pptx" | "xlsx", readOnly = false) {
  const app = new App();
  const bytes = kind === "xlsx" ? await xlsx() : fixture(`sample.${kind}`);
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const created = vi.spyOn(app.vault, "create");
  const saved = vi.spyOn(app.vault, "modifyBinary");
  const opened = vi.fn(async () => {});
  Object.assign(app.workspace, { getLeaf: () => ({ openFile: opened }) });
  const View = { docx: DocxView, pptx: PptxView, xlsx: XlsxView }[kind];
  const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf, { readOnly } as never);
  document.body.append(view.containerEl);
  const file = new TFile(`source.${kind}`) as unknown as RealFile;
  view.file = file;
  await view.onLoadFile(file);
  return { view, file, created, saved, opened };
}

function extract(view: { contentEl: HTMLElement }) {
  Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find(b => b.textContent === "Edit")!.click();
  const item = Menu.last!.items.find(i => i.title === "Create note from selection");
  expect(item).toBeDefined();
  item!.run!(new MouseEvent("click"));
}

it.each(["docx", "pptx", "xlsx"] as const)("%s has a companion button beside comments, also in read-only mode", async kind => {
  const { view, file, created, saved, opened } = await mount(kind, true);
  const button = view.contentEl.querySelector<HTMLButtonElement>('button[aria-label="Companion note"]');
  expect(button).not.toBeNull();
  button!.click();
  await vi.waitFor(() => expect(opened).toHaveBeenCalledTimes(1));
  expect(created.mock.calls[0][0]).toBe(`source.${kind} notes.md`);
  expect(saved).not.toHaveBeenCalled();
  await view.onUnloadFile(file);
});

it("extracts only selected document text and links to its heading", async () => {
  const { view, file, created, saved } = await mount("docx");
  const v = view as unknown as { editor: import("prosemirror-view").EditorView };
  const start = v.editor.state.doc.firstChild!.nodeSize + 1;
  v.editor.dispatch(v.editor.state.tr.setSelection(TextSelection.create(v.editor.state.doc, start, start + 5)));
  extract(view);
  await vi.waitFor(() => expect(created).toHaveBeenCalledTimes(1));
  expect(created.mock.calls[0][1]).toContain("[[source.docx#Doc Heading]]\n\nPlain");
  expect(created.mock.calls[0][1]).not.toContain("italic");
  expect(saved).not.toHaveBeenCalled();
  await view.onUnloadFile(file);
});

it("extracts the current slide rather than the entire presentation", async () => {
  const { view, file, created, saved } = await mount("pptx");
  extract(view);
  await vi.waitFor(() => expect(created).toHaveBeenCalledTimes(1));
  expect(created.mock.calls[0][1]).toContain("[[source.pptx#1]]");
  expect(created.mock.calls[0][1]).toContain("Hello Deck");
  expect(created.mock.calls[0][1]).not.toContain("Slide 2");
  expect(saved).not.toHaveBeenCalled();
  await view.onUnloadFile(file);
});

it("extracts a selected slide shape and an exact text selection", async () => {
  const { view, file, created } = await mount("pptx");
  const v = view as unknown as { rendered: RenderedShape[]; activeText: RenderedShape | null; setSelection(items: Element[]): void };
  const shape = v.rendered.find(r => r.textEl?.textContent?.includes("Subtitle here"))!;
  v.setSelection([shape.shape.el]);
  extract(view);
  await vi.waitFor(() => expect(created).toHaveBeenCalledTimes(1));
  expect(created.mock.calls[0][1]).toContain("Subtitle here");
  expect(created.mock.calls[0][1]).not.toContain("Hello Deck");
  const text = shape.textEl!.querySelector("span")!.firstChild!;
  const range = document.createRange();
  range.setStart(text, 0); range.setEnd(text, 3);
  document.getSelection()!.removeAllRanges();
  document.getSelection()!.addRange(range);
  v.activeText = shape;
  extract(view);
  await vi.waitFor(() => expect(created).toHaveBeenCalledTimes(2));
  expect(created.mock.calls[1][1]).toContain("\n\nSub\n");
  expect(created.mock.calls[1][1]).not.toContain("Subtitle here");
  document.getSelection()!.removeAllRanges();
  await view.onUnloadFile(file);
});

it("extracts selected spreadsheet cells as a Markdown table with a range link", async () => {
  const { view, file, created, saved } = await mount("xlsx");
  const box = view.contentEl.querySelector<HTMLInputElement>(".xlsx-name-box")!;
  box.value = "A1:B2";
  box.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  extract(view);
  await vi.waitFor(() => expect(created).toHaveBeenCalledTimes(1));
  expect(created.mock.calls[0][1]).toContain("[[source.xlsx#Data!A1:B2]]");
  expect(created.mock.calls[0][1]).toContain("| Name | Notes |");
  expect(saved).not.toHaveBeenCalled();
  await view.onUnloadFile(file);
});
