/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { TextSelection } from "prosemirror-state";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { BUILDING_BLOCKS, buildingBlock } from "../src/domain/docx/building-blocks";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { isoDay, formatDate } from "../src/domain/docx/chips";
import { fixture, W } from "./helpers";

afterEach(() => document.body.empty());

const OPTS = { taskList: "7", bulletList: "8", today: "2026-11-21" };
const chips = (nodes: any[]) => nodes.flatMap((n) => { const out: any[] = []; n.descendants((c: any) => void (c.type === schema.nodes.chip && out.push(c.attrs))); return out; });

describe("buildingBlock", () => {
  it("Google's six blocks", () => {
    expect(BUILDING_BLOCKS.map(([, l]) => l)).toEqual(["Meeting notes", "Email draft", "Product roadmap", "Review tracker", "Project assets", "Content launch tracker"]);
  });

  it("meeting notes: dated title, attendees, notes bullet, action item checkbox", () => {
    const nodes = buildingBlock("meeting", OPTS);
    expect(nodes.map((n) => [n.attrs.style, n.textContent])).toEqual([
      ["Heading1", "Nov 21, 2026 | Meeting title"], [null, "Attendees: "], ["Heading2", "Notes"], [null, ""], ["Heading2", "Action items"], [null, "☐ "],
    ]);
    expect(nodes[3].attrs.numId).toBe("8");
    expect(nodes[5].attrs.numId).toBe("7");
    expect(chips(nodes)).toMatchObject([{ kind: "date", date: "2026-11-21" }]);
  });

  it("email draft: To / Cc / Bcc / Subject rows, then a paragraph for the body", () => {
    const [table, after] = buildingBlock("email", OPTS);
    expect(table.type).toBe(schema.nodes.table);
    const rows: string[][] = [];
    table.forEach((r) => { const cells: string[] = []; r.forEach((c) => void cells.push(c.textContent)); rows.push(cells); });
    expect(rows).toEqual([["To", ""], ["Cc", ""], ["Bcc", ""], ["Subject", ""]]);
    expect(after.type).toBe(schema.nodes.paragraph);
    expect(table.attrs.borders).toBe("all");
  });

  it("trackers: bold headers and a first row with status dropdowns and date chips", () => {
    const [table] = buildingBlock("launch", OPTS);
    expect(table.firstChild!.textContent).toBe("ContentChannelOwnerStatusLaunch date");
    let bold = true;
    table.firstChild!.descendants((n) => void (n.isText && !n.marks.some((m) => m.type === schema.marks.bold) && (bold = false)));
    expect(bold).toBe(true);
    expect(chips([table.child(1)])).toMatchObject([{ kind: "dropdown", text: "Not started" }, { kind: "date", date: "2026-11-21" }]);
    expect(chips(buildingBlock("review", OPTS))[0].options.map((o: any) => o.text)).toEqual(["Not started", "In review", "Changes requested", "Approved"]);
    for (const kind of ["roadmap", "assets"] as const) expect(chips(buildingBlock(kind, OPTS))[0].options.map((o: any) => o.text)).toEqual(["Not started", "In progress", "Blocked", "Done"]);
  });
});

async function openDoc() {
  const app = new App();
  const bytes = fixture("sample.docx");
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void saved.push(data);
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officeDocx: false });
  document.body.appendChild(view.containerEl);
  const file = new TFile("blocks.docx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const insert = (title: string) => {
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Insert")!.click();
    Menu.last!.items.find((i) => i.title === "Building blocks")!.submenu!.items.find((i) => i.title === title)!.run!();
  };
  const caretAt = (pos: number) => v.editor.dispatch(v.editor.state.tr.setSelection(TextSelection.create(v.editor.state.doc, pos)));
  const savedPkg = async () => {
    await v.autosave.flush();
    return DocxDocument.load(new Uint8Array(saved.at(-1)!));
  };
  return { view, v, insert, caretAt, savedPkg };
}

describe("Insert › Building blocks", () => {
  it("lists the blocks", async () => {
    const { view } = await openDoc();
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Insert")!.click();
    expect(Menu.last!.items.find((i) => i.title === "Building blocks")!.submenu!.items.map((i) => i.title)).toEqual(BUILDING_BLOCKS.map(([, l]) => l));
  });

  it("a tracker goes below the caret's paragraph, caret in its first empty cell; saved as a table with a dropdown control", async () => {
    const { v, insert, caretAt, savedPkg } = await openDoc();
    const first = v.editor.state.doc.child(0);
    caretAt(first.nodeSize - 1);
    insert("Review tracker");
    const doc = v.editor.state.doc;
    expect(doc.child(0).eq(first)).toBe(true);
    expect(doc.child(1).type).toBe(schema.nodes.table);
    expect(v.editor.state.selection.$from.parent.textContent).toBe("");
    expect(v.editor.state.selection.$from.node(-1).type).toBe(schema.nodes.table_cell);
    const body = (await savedPkg()).body;
    const tbl = body.getElementsByTagNameNS(W, "tbl")[0];
    expect(Array.from(tbl.getElementsByTagNameNS(W, "tr")[0].getElementsByTagNameNS(W, "t")).map((t) => t.textContent)).toEqual(["Reviewer", "Status", "Notes"]);
    expect(tbl.getElementsByTagNameNS(W, "tblBorders")[0].getElementsByTagNameNS(W, "insideV")[0].getAttributeNS(W, "val")).toBe("single");
    const list = tbl.getElementsByTagNameNS(W, "dropDownList")[0];
    expect(Array.from(list.getElementsByTagNameNS(W, "listItem")).map((i) => i.getAttributeNS(W, "displayText"))).toEqual(["Not started", "In review", "Changes requested", "Approved"]);
  });

  it("meeting notes replace an empty paragraph; the lists are real Word lists, the title has today's date", async () => {
    const { v, insert, caretAt, savedPkg } = await openDoc();
    const tr = v.editor.state.tr.insert(0, schema.nodes.paragraph.create());
    v.editor.dispatch(tr);
    caretAt(1);
    const count = v.editor.state.doc.childCount;
    insert("Meeting notes");
    expect(v.editor.state.doc.childCount).toBe(count - 1 + 6);
    expect(v.editor.state.doc.child(0).textContent).toBe(`${formatDate(isoDay(), "MMM d, yyyy")} | Meeting title`);
    const pkg = await savedPkg();
    const paras = Array.from(pkg.body.getElementsByTagNameNS(W, "p")).slice(0, 6);
    expect(paras[0].getElementsByTagNameNS(W, "pStyle")[0].getAttributeNS(W, "val")).toBe("Heading1");
    expect(paras[0].getElementsByTagNameNS(W, "date")).toHaveLength(1);
    const numIds = [paras[3], paras[5]].map((p) => p.getElementsByTagNameNS(W, "numId")[0]?.getAttributeNS(W, "val"));
    expect(numIds.every(Boolean)).toBe(true);
    expect(numIds[0]).not.toBe(numIds[1]);
    expect(pkg.listFormat(numIds[1]!, 0)).toBe("none"); // the checklist definition: no bullet, the box is text
  });

  it("from the @ menu: the @ goes, the block is inserted", async () => {
    const { v, view, caretAt } = await openDoc();
    const end = v.editor.state.doc.child(0).nodeSize - 1;
    caretAt(end);
    v.editor.dispatch(v.editor.state.tr.insertText(" @"));
    v.atMenu();
    const item = Menu.last!.items.find((i) => i.title === "Email draft")!;
    item.run!();
    expect(v.editor.state.doc.child(0).textContent.endsWith("@")).toBe(false);
    expect(v.editor.state.doc.child(1).type).toBe(schema.nodes.table);
    expect(view).toBeTruthy();
  });

  it("undo takes the whole block out", async () => {
    const { v, insert, caretAt } = await openDoc();
    const before = v.editor.state.doc;
    caretAt(before.child(0).nodeSize - 1);
    insert("Product roadmap");
    const { undo } = await import("prosemirror-history");
    undo(v.editor.state, v.editor.dispatch);
    expect(v.editor.state.doc.eq(before)).toBe(true);
  });
});
