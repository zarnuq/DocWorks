/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, Notice, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { FilePicker } from "../src/presentation/common/file-picker";
import { revisions } from "../src/domain/docx/revisions";
import { blankDocx } from "../src/infrastructure/blank";
import { fixture } from "./helpers";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

async function setup(otherBytes?: Uint8Array, readOnly = false) {
  const app = new App();
  const mine = new TFile("Drafts/Draft.docx");
  const other = new TFile("Final.docx");
  const note = new TFile("Notes.md");
  const files = new Map<string, Uint8Array>([[mine.path, fixture("sample.docx")], [other.path, otherBytes ?? (await blankDocx())]]);
  (mine as any).parent = { path: "Drafts" };
  app.vault.readBinary = async (f: any) => {
    const b = files.get(f.path)!;
    return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
  };
  app.vault.getFiles = () => [mine, other, note] as any;
  const created: { path: string; data: ArrayBuffer }[] = [];
  (app.vault as any).createBinary = async (path: string, data: ArrayBuffer) => {
    created.push({ path, data });
    return new TFile(path);
  };
  app.vault.getAbstractFileByPath = (path: string) => (created.some((c) => c.path === path) || path === "Drafts/Comparison of Draft and Final.docx" && (app as any).taken ? new TFile(path) : null) as any;
  const opened: string[] = [];
  (app.workspace as any).getLeaf = () => ({ openFile: async (f: any) => void opened.push(f.path) });
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officeDocx: false, author: "Ada", readOnly });
  document.body.appendChild(view.containerEl);
  view.file = mine as unknown as RealTFile;
  await view.onLoadFile(mine as unknown as RealTFile);
  let picker: any = null;
  vi.spyOn(FilePicker.prototype, "open").mockImplementation(function (this: any) { picker = this; });
  const flush = () => new Promise((r) => setTimeout(r, 0));
  const top = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
  const button = (text: string) => Array.from(top().querySelectorAll("button")).find((b) => b.textContent === text)!;
  const openMenu = () => {
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Tools")!.click();
    Menu.last!.items.find((i) => i.title === "Compare documents")!.run!();
  };
  return { app, view, created, opened, picker: () => picker, flush, top, button, openMenu };
}

describe("Tools › Compare documents", () => {
  it("offers only the other Word documents, Compare disabled until one is chosen", async () => {
    const { openMenu, button, picker } = await setup();
    openMenu();
    expect(button("Compare").disabled).toBe(true);
    button("Choose…").click();
    expect(picker().getItems().map((f: any) => f.path)).toEqual(["Final.docx"]);
    picker().onChooseItem(picker().getItems()[0]);
    expect(button("Compare").disabled).toBe(false);
  });

  it("makes 'Comparison of Draft and Final.docx' beside the file, with the differences as suggestions by the named author, and opens it", async () => {
    const { openMenu, button, picker, top, created, opened } = await setup();
    openMenu();
    button("Choose…").click();
    picker().onChooseItem(picker().getItems()[0]);
    const author = top().querySelector<HTMLInputElement>('input[type="text"]')!;
    author.value = "Reviewer";
    author.dispatchEvent(new Event("change"));
    button("Compare").click();
    await vi.waitFor(() => expect(opened).toHaveLength(1), { timeout: 5000 });
    expect(Notice.messages.some((m) => m.startsWith("Couldn't"))).toBe(false);
    expect(created.map((c) => c.path)).toEqual(["Drafts/Comparison of Draft and Final.docx"]);
    expect(opened).toEqual(["Drafts/Comparison of Draft and Final.docx"]);
    const result = await DocxDocument.load(new Uint8Array(created[0].data));
    const revs = revisions(result.body);
    expect(revs.length).toBeGreaterThan(0);
    expect(new Set(revs.map((r) => r.author))).toEqual(new Set(["Reviewer"]));
  });

  it("identical documents: says so, makes nothing", async () => {
    const { openMenu, button, picker, created } = await setup(fixture("sample.docx"));
    openMenu();
    button("Choose…").click();
    picker().onChooseItem(picker().getItems()[0]);
    const before = Notice.messages.length;
    button("Compare").click();
    await vi.waitFor(() => expect(Notice.messages.length).toBeGreaterThan(before), { timeout: 5000 });
    expect(Notice.messages.slice(before)).toEqual(["No differences found."]);
    expect(created).toHaveLength(0);
  });

  it("a taken name gets a number", async () => {
    const { app, openMenu, button, picker, created, opened } = await setup();
    (app as any).taken = true;
    openMenu();
    button("Choose…").click();
    picker().onChooseItem(picker().getItems()[0]);
    button("Compare").click();
    await vi.waitFor(() => expect(opened).toHaveLength(1), { timeout: 5000 });
    expect(created[0].path).toBe("Drafts/Comparison of Draft and Final 2.docx");
  });

  it("works on a read-only file (it doesn't change it)", async () => {
    const { view } = await setup(undefined, true);
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Tools")!.click();
    expect(Menu.last!.items.map((i) => i.title)).toContain("Compare documents");
  });
});
