import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DEFAULT_SETTINGS, type DocWorksSettings } from "../src/presentation/common/settings";
import { PPT_COMMAND_IDS } from "../src/presentation/pptx/powerpoint-ui";
import { fixture } from "./helpers";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

const PPT: DocWorksSettings = { ...DEFAULT_SETTINGS, officePptx: true, author: "Ada" };

async function openPpt(settings: DocWorksSettings = { ...PPT }) {
  const app = new App();
  app.vault.readBinary = async () => fixture("sample.pptx").buffer as ArrayBuffer;
  const modify = vi.spyOn(app.vault, "modifyBinary");
  const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
  document.body.appendChild(view.containerEl);
  const file = new TFile("a.pptx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return { view, v: view as any, file, modify, settings };
}

const tabs = (view: PptxView) => Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-rb-tab")).filter((t) => !t.hidden).map((t) => t.textContent);
const clickTab = (view: PptxView, label: string) => Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-rb-tab")).find((t) => t.textContent === label && !t.hidden)!.click();
const ribbonButton = (view: PptxView, id: string) => view.contentEl.querySelector<HTMLButtonElement>(`.dw-rb-panel [data-cmd="${id}"]`)!;

describe("PptxView PowerPoint-style interface", () => {
  it("opens with the ribbon and status bar instead of the Google toolbar", async () => {
    const { view } = await openPpt();
    expect(view.contentEl.querySelector(".dw-rb")).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-statusbar")).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-menubar")).toBeNull();
    expect(view.contentEl.querySelector(".pptx-toolbar")).toBeNull();
    expect(tabs(view)).toEqual(["Home", "Insert", "Design", "Transitions", "Animations", "Slide Show", "Review", "View", "Help"]);
  });

  it("provides every command PowerPoint's layout names", async () => {
    const { v } = await openPpt();
    const cmds = v.pptCommands() as Record<string, unknown>;
    expect(PPT_COMMAND_IDS.filter((id) => !cmds[id])).toEqual([]);
  });

  it("shows Slide N of M and follows slide changes", async () => {
    const { view, v } = await openPpt();
    const slide = () => view.contentEl.querySelector('.dw-statusbar [data-item="slide"]')!.textContent;
    const total = v.doc.slides().length;
    expect(slide()).toBe(`Slide 1 of ${total}`);
    v.goTo(1);
    expect(slide()).toBe(`Slide 2 of ${total}`);
  });

  it("Home › New Slide adds a slide", async () => {
    const { view, v } = await openPpt();
    const before = v.doc.slides().length;
    ribbonButton(view, "SlideNew").click();
    expect(v.doc.slides().length).toBe(before + 1);
  });

  it("shows Shape Format while a shape is selected", async () => {
    const { view, v } = await openPpt();
    const shape = v.rendered.find((r: { shape: { kind: string; el: Element; inherited?: boolean } }) => r.shape.kind === "sp" && !r.shape.inherited) ?? v.rendered.find((r: { shape: { el: Element; inherited?: boolean } }) => r.shape.el.localName === "sp" && !r.shape.inherited);
    v.select(shape);
    expect(tabs(view)).toContain("Shape Format");
  });

  it("keeps read-only files to the commands that look at them, and never saves", async () => {
    const { view, file, modify } = await openPpt({ ...PPT, readOnly: true });
    expect(tabs(view)).not.toContain("Insert");
    expect(tabs(view)).toContain("Slide Show");
    await view.onUnloadFile(file);
    expect(modify).not.toHaveBeenCalled();
  });

  it("switches layouts on reopen without saving", async () => {
    const { view, file, modify, settings } = await openPpt();
    settings.officePptx = false;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(view.contentEl.querySelector(".dw-rb")).toBeNull();
    expect(view.contentEl.querySelector(".pptx-toolbar")).not.toBeNull();
    expect(modify).not.toHaveBeenCalled();
  });

  it("Slide Show › From Beginning starts the show from slide 1", async () => {
    const { view, v } = await openPpt();
    const start = vi.spyOn(v, "startPresentation").mockImplementation(() => {});
    clickTab(view, "Slide Show");
    ribbonButton(view, "SlideShowFromBeginning").click();
    expect(start).toHaveBeenCalledWith(true);
  });
});

describe("PptxView PowerPoint-style right-click and keys", () => {
  const items = () => Array.from(document.querySelectorAll(".dw-ctx > .menu:not(.dw-ctx-sub) > *")).map((e) => (e.classList.contains("menu-separator") ? "---" : e.querySelector(".menu-item-title")!.textContent));
  const scopeKey = (v: { scope: { keys: { mods: string[]; key: string; fn: () => unknown }[] } }, mods: string, key: string) => v.scope.keys.filter((k) => k.mods.join() === mods && k.key === key);

  it("right-click on a shape shows PowerPoint's menu and mini toolbar", async () => {
    const { v } = await openPpt();
    const rs = v.rendered.find((r: { shape: { el: Element; inherited?: boolean } }) => r.shape.el.localName === "sp" && !r.shape.inherited);
    rs.el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    expect(items().slice(0, 4)).toEqual(["Cut", "Copy", "Paste", "Keep Text Only"]);
    expect(items()).toEqual(expect.arrayContaining(["Group", "Bring to Front", "Send to Back", "Link…", "New Comment", "Format Shape…", "Create Note from Selection"]));
    expect(document.querySelector(".dw-ctx-mini [aria-label^='Bold']")).not.toBeNull();
  });

  it("Ctrl+M adds a slide while the slide has focus", async () => {
    const { v } = await openPpt();
    const before = v.doc.slides().length;
    v.stageEl.focus();
    const handled = scopeKey(v, "Mod", "m").map((k) => k.fn());
    expect(handled).toContain(false);
    expect(v.doc.slides().length).toBe(before + 1);
  });

  it("leaves Google's list keys alone in the PowerPoint layout, and keeps them in Google's", async () => {
    const { v } = await openPpt();
    v.stageEl.focus();
    expect(scopeKey(v, "Mod,Shift", "8").map((k) => k.fn())).toEqual([undefined]);
    const google = await openPpt({ ...PPT, officePptx: false });
    google.v.stageEl.focus();
    expect(scopeKey(google.v, "Mod,Shift", "8").map((k) => k.fn())).toEqual([false]);
    // Not registered at all in the Google layout: Ctrl+M stays the app's.
    expect(scopeKey(google.v, "Mod", "m")).toEqual([]);
  });

  it("closes the right-click menu when the file closes", async () => {
    const { v, view, file } = await openPpt();
    v.stageEl.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
    await view.onUnloadFile(file);
    expect(document.querySelector(".dw-ctx")).toBeNull();
  });
});

describe("PptxView PowerPoint-style screenshot fixes", () => {
  it("fills Font and Size from a selected shape's text", async () => {
    const { view, v } = await openPpt();
    const rs = v.rendered.find((r: { shape: { el: Element; inherited?: boolean; text?: { paras: { runs: unknown[] }[] } } }) => r.shape.el.localName === "sp" && !r.shape.inherited && r.shape.text?.paras.some((p) => p.runs.length));
    v.select(rs);
    expect(view.contentEl.querySelector('.dw-rb-select[data-cmd="Font"] .dw-rb-select-label')!.textContent).not.toBe("");
    expect(view.contentEl.querySelector('.dw-rb-select[data-cmd="FontSize"] .dw-rb-select-label')!.textContent).toMatch(/^\d/);
  });
});

describe("PptxView PowerPoint-style review fixes", () => {
  const scopeKey = (v: { scope: { keys: { mods: string[]; key: string; fn: () => unknown }[] } }, mods: string, key: string) => v.scope.keys.filter((k) => k.mods.join() === mods && k.key === key);
  const firstTextShape = (v: { rendered: { shape: { el: Element; inherited?: boolean; text?: { paras: { runs: unknown[] }[] } }; textEl?: HTMLElement }[] }) =>
    v.rendered.find((r) => r.shape.el.localName === "sp" && !r.shape.inherited && r.shape.text?.paras.some((p) => p.runs.length))!;

  it("registers PowerPoint's keys only while its layout is on", async () => {
    const google = await openPpt({ ...PPT, officePptx: false });
    expect(scopeKey(google.v, "Mod", "m")).toEqual([]);
    const { v, view, file } = await openPpt();
    expect(scopeKey(v, "Mod", "m").length).toBe(1);
    await view.onUnloadFile(file);
    expect(scopeKey(v, "Mod", "m")).toEqual([]);
  });

  it("leaves keys to the speaker notes while typing in them", async () => {
    const { v } = await openPpt();
    v.select(firstTextShape(v));
    const before = new XMLSerializer().serializeToString(firstTextShape(v).shape.el);
    v.notesEl.focus();
    expect(scopeKey(v, "Mod", "e").map((k) => k.fn())).toEqual([undefined]);
    expect(new XMLSerializer().serializeToString(firstTextShape(v).shape.el)).toBe(before);
  });

  it("Ctrl+Shift+C copies formatting without arming the painter; Ctrl+Shift+V pastes it on the selection", async () => {
    const { v } = await openPpt();
    const shapes = v.rendered.filter((r: { shape: { el: Element; inherited?: boolean } }) => r.shape.el.localName === "sp" && !r.shape.inherited);
    v.select(shapes[0]);
    v.stageEl.focus();
    scopeKey(v, "Mod,Shift", "c")[0].fn();
    expect(v.painter).toBeNull();
    expect(v.copiedFormat).toBeTruthy();
    v.select(shapes[1]);
    v.stageEl.focus();
    const paint = vi.spyOn(v, "applyCopiedFormat");
    scopeKey(v, "Mod,Shift", "v")[0].fn();
    expect(paint).toHaveBeenCalled();
  });

  it("Ctrl+M adds a slide from the thumbnails too", async () => {
    const { v } = await openPpt();
    const before = v.doc.slides().length;
    v.thumbsEl.focus();
    scopeKey(v, "Mod", "m")[0].fn();
    expect(v.doc.slides().length).toBe(before + 1);
  });

  it("Google's grid view key is off in PowerPoint mode", async () => {
    const { v } = await openPpt();
    scopeKey(v, "Mod,Alt", "1")[0].fn();
    expect(v.gridView).toBe(false);
  });

  it("right-click in text being edited shows PowerPoint's menu and mini toolbar", async () => {
    const { v } = await openPpt();
    const rs = firstTextShape(v);
    v.select(rs);
    rs.textEl!.setAttribute("contenteditable", "true");
    rs.textEl!.focus();
    rs.textEl!.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
    expect(document.querySelector(".dw-ctx-mini")).not.toBeNull();
  });

  it("right-click on a table cell keeps the cell's insert / delete / merge actions, and Select All", async () => {
    const { v } = await openPpt();
    v.goTo(2);
    const cell = v.rendered.find((r: { shape: { kind: string } }) => r.shape.kind === "cell");
    cell.el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    const titles = Array.from(document.querySelectorAll(".dw-ctx .menu-item-title")).map((t) => t.textContent);
    expect(titles).toContain("Table");
    expect(titles).toContain("Select All");
  });

  it("the status bar's Notes button hides and shows the speaker notes", async () => {
    const { view, v } = await openPpt();
    const notes = view.contentEl.querySelector<HTMLButtonElement>('.dw-statusbar [aria-label="Notes"]')!;
    notes.click();
    expect(v.notesEl.parentElement.hidden).toBe(true);
    view.contentEl.querySelector<HTMLButtonElement>('.dw-statusbar [aria-label="Notes"]')!.click();
    expect(v.notesEl.parentElement.hidden).toBe(false);
  });
});
