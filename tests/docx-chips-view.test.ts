/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { TextSelection } from "prosemirror-state";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { FilePicker } from "../src/presentation/common/file-picker";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { isoDay, formatDate } from "../src/domain/docx/chips";
import { fixture, W } from "./helpers";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
  vi.useRealTimers();
});

async function openDoc(readOnly = false) {
  const app = new App();
  const bytes = fixture("sample.docx");
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  (app.vault as any).getName = () => "notes";
  app.vault.getFiles = () => [new TFile("Projects/Plan.md") as never];
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void saved.push(data);
  const opened: string[] = [];
  (app.workspace as any).openLinkText = async (link: string) => void opened.push(link);
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officeDocx: false, readOnly });
  document.body.appendChild(view.containerEl);
  const file = new TFile("chips.docx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const editor = () => v.editor;
  /** Caret at the end of the first paragraph. */
  const caretAtEnd = () => {
    const ed = editor();
    const end = ed.state.doc.child(0).nodeSize - 1;
    ed.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, end)));
    return end;
  };
  const insertMenu = (...path: string[]) => {
    const b = Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Insert")!;
    b.click();
    let menu = Menu.last!;
    for (const title of path) {
      const item = menu.items.find((i) => i.title === title);
      if (!item) throw new Error(`No "${title}": ${menu.items.map((i) => i.title).join(", ")}`);
      if (item.submenu) menu = item.submenu;
      else item.run!(new MouseEvent("click"));
    }
  };
  const chipNodes = () => {
    const out: any[] = [];
    editor().state.doc.descendants((n: any, pos: number) => void (n.type === schema.nodes.chip && out.push({ n, pos })));
    return out;
  };
  const savedXml = async () => {
    await v.autosave.flush();
    const docx = await DocxDocument.load(new Uint8Array(saved.at(-1)!));
    return docx.body;
  };
  const modal = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
  const clickButton = (text: string) => Array.from(modal().querySelectorAll("button")).find((b) => b.textContent === text)!.click();
  const flush = () => new Promise((r) => setTimeout(r, 0));
  return { view, v, editor, caretAtEnd, insertMenu, chipNodes, savedXml, modal, clickButton, flush, opened, app };
}

describe("smart chips in the document editor", () => {
  it("Insert › Smart chips lists Google's chips", async () => {
    const { insertMenu } = await openDoc();
    insertMenu("Smart chips");
    const b = Array.from(document.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Insert")!;
    b.click();
    const sub = Menu.last!.items.find((i) => i.title === "Smart chips")!.submenu!;
    expect(sub.items.map((i) => i.title)).toEqual(["Today", "Tomorrow", "Yesterday", "Date…", "---", "Dropdown…", "File…", "Person…"]);
  });

  it("Today inserts a date chip and saves a Word date picker", async () => {
    const { caretAtEnd, insertMenu, chipNodes, savedXml } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Today");
    const [c] = chipNodes();
    expect(c.n.attrs).toMatchObject({ kind: "date", date: isoDay(), text: formatDate(isoDay(), "MMM d, yyyy") });
    const body = await savedXml();
    const sdt = body.getElementsByTagNameNS(W, "sdt")[0];
    expect(sdt.getElementsByTagNameNS(W, "date")[0].getAttributeNS(W, "fullDate")).toBe(`${isoDay()}T00:00:00Z`);
    expect(sdt.parentElement!.localName).toBe("p");
  });

  it("Tomorrow and Yesterday are a day either side", async () => {
    const { caretAtEnd, insertMenu, chipNodes } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Tomorrow");
    insertMenu("Smart chips", "Yesterday");
    expect(chipNodes().map((c) => c.n.attrs.date)).toEqual([isoDay(1), isoDay(-1)]);
  });

  it("Date… picks a day and format in a dialog", async () => {
    const { caretAtEnd, insertMenu, chipNodes, modal, clickButton, flush } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Date…");
    const date = modal().querySelector<HTMLInputElement>('input[type="date"]')!;
    date.value = "2026-12-25";
    date.dispatchEvent(new Event("change"));
    const format = modal().querySelector<HTMLSelectElement>("select")!;
    expect(Array.from(format.options).map((o) => o.text)).toContain("Dec 25, 2026");
    format.value = "yyyy-MM-dd";
    format.dispatchEvent(new Event("change"));
    clickButton("Insert");
    await flush();
    expect(chipNodes()[0].n.attrs).toMatchObject({ date: "2026-12-25", format: "yyyy-MM-dd", text: "2026-12-25" });
  });

  it("Cancel inserts nothing", async () => {
    const { caretAtEnd, insertMenu, chipNodes, clickButton, flush } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Date…");
    clickButton("Cancel");
    await flush();
    expect(chipNodes()).toHaveLength(0);
  });

  it("Dropdown… makes a dropdown from a ready-made list and saves a Word drop-down list", async () => {
    const { caretAtEnd, insertMenu, chipNodes, modal, clickButton, flush, savedXml } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Dropdown…");
    const preset = modal().querySelector<HTMLSelectElement>("select")!;
    preset.value = "Project status";
    preset.dispatchEvent(new Event("change"));
    expect(modal().querySelector("textarea")!.value).toBe("Not started\nIn progress\nBlocked\nDone");
    clickButton("Insert");
    await flush();
    expect(chipNodes()[0].n.attrs.text).toBe("Not started");
    const items = Array.from((await savedXml()).getElementsByTagNameNS(W, "listItem")).map((i) => i.getAttributeNS(W, "displayText"));
    expect(items).toEqual(["Not started", "In progress", "Blocked", "Done"]);
  });

  it("an empty dropdown list inserts nothing", async () => {
    const { caretAtEnd, insertMenu, chipNodes, clickButton, flush } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Dropdown…");
    clickButton("Insert");
    await flush();
    expect(chipNodes()).toHaveLength(0);
  });

  it("clicking a dropdown chip lists its items; picking one changes it", async () => {
    const { caretAtEnd, insertMenu, chipNodes, modal, clickButton, flush, v } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Dropdown…");
    modal().querySelector("textarea")!.value = "Low\nHigh";
    clickButton("Insert");
    await flush();
    v.chipClicked(chipNodes()[0].pos, new MouseEvent("click"));
    expect(Menu.last!.items.map((i) => [i.title, i.checked])).toEqual([["Low", true], ["High", false], ["---", undefined], ["Edit options…", undefined]]);
    Menu.last!.items[1].run!();
    expect(chipNodes()[0].n.attrs.text).toBe("High");
  });

  it("Edit options… keeps the choice when it's still listed, else takes the first", async () => {
    const { caretAtEnd, insertMenu, chipNodes, modal, clickButton, flush, v } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Dropdown…");
    modal().querySelector("textarea")!.value = "Low\nHigh";
    clickButton("Insert");
    await flush();
    v.chipClicked(chipNodes()[0].pos, new MouseEvent("click"));
    Menu.last!.items.find((i) => i.title === "Edit options…")!.run!();
    modal().querySelector("textarea")!.value = "Medium\nLow";
    clickButton("Done");
    await flush();
    expect(chipNodes()[0].n.attrs.options.map((o: any) => o.text)).toEqual(["Medium", "Low"]);
    expect(chipNodes()[0].n.attrs.text).toBe("Low");
  });

  it("clicking a date chip opens the date dialog with its day", async () => {
    const { caretAtEnd, insertMenu, chipNodes, modal, clickButton, flush, v } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "Today");
    v.chipClicked(chipNodes()[0].pos, new MouseEvent("click"));
    const date = modal().querySelector<HTMLInputElement>('input[type="date"]')!;
    expect(date.value).toBe(isoDay());
    date.value = "2027-01-01";
    date.dispatchEvent(new Event("change"));
    clickButton("Done");
    await flush();
    expect(chipNodes()[0].n.attrs).toMatchObject({ date: "2027-01-01", text: "Jan 1, 2027" });
  });

  it("File… links a vault file by name, opened in a new tab with Ctrl+click", async () => {
    let picker: any = null;
    vi.spyOn(FilePicker.prototype, "open").mockImplementation(function (this: any) { picker = this; });
    const { caretAtEnd, insertMenu, editor, view, opened } = await openDoc();
    caretAtEnd();
    insertMenu("Smart chips", "File…");
    picker.onChooseItem(picker.getItems()[0]);
    const ed = editor();
    const text = ed.state.doc.child(0).lastChild;
    expect(text.text).toBe("Plan");
    expect(text.marks.find((m: any) => m.type === schema.marks.link).attrs.href).toBe("obsidian://open?vault=notes&file=Projects%2FPlan.md");
    const a = view.contentEl.querySelector<HTMLAnchorElement>('a.docx-link[href^="obsidian:"]')!;
    a.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, ctrlKey: true }));
    expect(opened).toEqual(["Projects/Plan.md"]);
  });

  it("Person… links an email address, and refuses something else", async () => {
    const { caretAtEnd, insertMenu, editor, modal, flush } = await openDoc();
    caretAtEnd();
    const answer = async (text: string) => {
      insertMenu("Smart chips", "Person…");
      const input = modal().querySelector<HTMLInputElement>("input")!;
      input.value = text;
      input.dispatchEvent(new Event("change"));
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      await flush();
    };
    await answer("not an address");
    expect(editor().state.doc.textContent).not.toContain("not an address");
    await answer("ada@example.com");
    const last = editor().state.doc.child(0).lastChild;
    expect(last.text).toBe("ada@example.com");
    expect(last.marks.find((m: any) => m.type === schema.marks.link).attrs.href).toBe("mailto:ada@example.com");
  });

  it("typing @ at a word start opens the chip menu at the caret; a pick replaces the @", async () => {
    const { caretAtEnd, editor, chipNodes } = await openDoc();
    vi.useFakeTimers();
    caretAtEnd();
    const ed = editor();
    const from = ed.state.selection.from;
    ed.dispatch(ed.state.tr.insertText(" @"));
    Menu.last = null;
    const handled = ed.someProp("handleTextInput", (f: any) => f(ed, from + 1, from + 1, "@"));
    expect(handled).toBeFalsy();
    vi.runAllTimers();
    expect(Menu.last!.items.map((i) => i.title)).toContain("Today");
    Menu.last!.items.find((i) => i.title === "Today")!.run!();
    expect(chipNodes()).toHaveLength(1);
    expect(ed.state.doc.child(0).textContent.endsWith(` ${formatDate(isoDay(), "MMM d, yyyy")}`)).toBe(true);
    expect(ed.state.doc.textContent).not.toContain("@");
  });

  it("an @ inside a word (an email address) opens nothing", async () => {
    const { caretAtEnd, editor } = await openDoc();
    vi.useFakeTimers();
    caretAtEnd();
    const ed = editor();
    ed.dispatch(ed.state.tr.insertText("x@"));
    Menu.last = null;
    ed.someProp("handleTextInput", (f: any) => f(ed, 1, 1, "@"));
    vi.runAllTimers();
    expect(Menu.last).toBeNull();
  });

  it("read-only files: chips don't react to clicks and Insert isn't offered", async () => {
    const { view, v } = await openDoc(true);
    expect(Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent)).not.toContain("Insert");
    const handled = v.editor.someProp("handleClickOn", (f: any) => f(v.editor, 0, schema.nodes.chip.create(), 0, new MouseEvent("click"), true));
    expect(handled).toBeFalsy();
  });

  it("an unchanged file with no chips still saves nothing on open", async () => {
    const { v } = await openDoc();
    expect(v.autosave.dirty).toBe(false);
  });
});
