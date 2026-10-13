import { expect, it, vi } from "vitest";
import { App, Menu, TFile, WorkspaceLeaf } from "obsidian";
import { DocxView } from "../src/presentation/docx/docx-view";
import { App as TestApp, WorkspaceLeaf as TestLeaf } from "./obsidian-stub";
import { exportDocxToMarkdown } from "../src/presentation/common/export-markdown";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { docxSchema as s } from "../src/domain/docx/schema";

const file = (path: string) => Object.assign(new TFile(), { path });

it("offers one Export… item in the document pane menu", () => {
  const view = new DocxView(new TestLeaf(new TestApp()) as unknown as WorkspaceLeaf);
  const run = vi.spyOn(view, "exportAs").mockResolvedValue();
  const menu = new Menu();
  const titles: string[] = [], callbacks: (() => void)[] = [];
  vi.spyOn(menu, "addItem").mockImplementation(cb => {
    const item = { setTitle: (text: string) => { titles.push(text); return item; }, setIcon: () => item, setSection: () => item, onClick: (fn: () => void) => { callbacks.push(fn); return item; } };
    cb(item as never); return menu;
  });
  view.onPaneMenu(menu, "more-options");
  expect(titles).toEqual(["Export…"]);
  callbacks[0]();
  expect(run).toHaveBeenCalledOnce();
});
function vaultApp(existing: string[] = []) {
  const files = new Map(existing.map(path => [path, file(path)]));
  const openFile = vi.fn();
  const app = {
    vault: {
      getAbstractFileByPath: (path: string) => files.get(path) ?? null,
      createFolder: vi.fn(async (path: string) => { const f = file(path); files.set(path, f); return f; }),
      create: vi.fn(async (path: string, _text: string) => { const f = file(path); files.set(path, f); return f; }),
      createBinary: vi.fn(async (path: string, _data: ArrayBuffer) => { const f = file(path); files.set(path, f); return f; }),
    },
    workspace: { getLeaf: vi.fn(() => ({ openFile })) },
    fileManager: { trashFile: vi.fn(async (f: TFile) => { files.delete(f.path); }) },
  };
  return { app: app as unknown as App, files, openFile, vault: app.vault, workspace: app.workspace };
}

it("creates a new note with images beside the source and avoids existing paths", async () => {
  const a = vaultApp(["folder/Report.md", "folder/Report 1 attachments"]);
  const docx = await DocxDocument.load(await blankDocx());
  const id = docx.addImage(new Uint8Array([1, 2]), "png");
  const doc = s.nodes.doc.create(null, s.nodes.paragraph.create(null, [s.text("Current text"), s.nodes.image.create({ rId: id })]));
  const result = await exportDocxToMarkdown(a.app, file("folder/Report.docx"), { document: docx, content: doc });
  expect(result?.path).toBe("folder/Report 2.md");
  expect(a.vault.createBinary).toHaveBeenCalledWith("folder/Report 2 attachments/image-1.png", expect.any(ArrayBuffer));
  expect(a.vault.create).toHaveBeenCalledWith("folder/Report 2.md", "Current text![](<Report%202%20attachments/image-1.png>)\n");
  expect(a.workspace.getLeaf).toHaveBeenCalledWith("tab");
  expect(a.openFile).toHaveBeenCalledWith(result);
  expect(a.files.has("folder/Report.md")).toBe(true);
});

it("creates no attachment folder for a text-only document", async () => {
  const a = vaultApp();
  const docx = await DocxDocument.load(await blankDocx());
  await exportDocxToMarkdown(a.app, file("Report.docx"), { document: docx, content: s.nodes.doc.create(null, s.nodes.paragraph.create(null, s.text("Text"))) });
  expect(a.vault.createFolder).not.toHaveBeenCalled();
  expect(a.vault.create).toHaveBeenCalledWith("Report.md", "Text\n");
});

it("removes its new attachments if writing the note fails", async () => {
  const a = vaultApp(["Existing.md"]);
  const docx = await DocxDocument.load(await blankDocx());
  const id = docx.addImage(new Uint8Array([1]), "png");
  a.vault.create.mockRejectedValueOnce(new Error("Write failed"));
  expect(await exportDocxToMarkdown(a.app, file("Report.docx"), { document: docx, content: s.nodes.doc.create(null, s.nodes.paragraph.create(null, s.nodes.image.create({ rId: id }))) })).toBeNull();
  expect([...a.files.keys()]).toEqual(["Existing.md"]);
  expect(a.openFile).not.toHaveBeenCalled();
});
