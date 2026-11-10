import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { NodeSelection, TextSelection } from "prosemirror-state";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DEFAULT_SETTINGS, DocxView, DocWorksSettings } from "../src/presentation/docx/docx-view";
import { WORD_COMMAND_IDS } from "../src/presentation/docx/word-ui";
import { fixture } from "./helpers";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

const WORD: DocWorksSettings = { ...DEFAULT_SETTINGS, officeDocx: true, author: "Ada" };

async function openWord(settings: DocWorksSettings = WORD) {
  const app = new App();
  app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
  const modify = vi.spyOn(app.vault, "modifyBinary");
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
  document.body.appendChild(view.containerEl);
  const file = new TFile("a.docx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const v = view as any;
  return { view, v, app, file, modify, editor: v.editor as import("prosemirror-view").EditorView };
}

const tabs = (view: DocxView) => Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-rb-tab")).filter((t) => !t.hidden).map((t) => t.textContent);
const clickTab = (view: DocxView, label: string) => Array.from(view.contentEl.querySelectorAll<HTMLElement>(".dw-rb-tab")).find((t) => t.textContent === label && !t.hidden)!.click();
const ribbonButton = (view: DocxView, id: string) => view.contentEl.querySelector<HTMLButtonElement>(`.dw-rb-panel [data-cmd="${id}"]`)!;
/** Put the cursor at the first text position inside the first node of `type`. */
function cursorIn(editor: import("prosemirror-view").EditorView, type: string) {
  let at = -1;
  editor.state.doc.descendants((n, pos) => {
    if (at < 0 && n.type.name === type) at = pos;
    return at < 0;
  });
  editor.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(at + 1))));
}

describe("DocxView Word-style interface", () => {
  it("opens with the ribbon and status bar instead of the Google toolbar", async () => {
    const { view } = await openWord();
    expect(view.contentEl.querySelector(".dw-rb")).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-statusbar")).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-menubar")).toBeNull();
    expect(view.contentEl.querySelector(".dw-toolbar")).toBeNull();
    expect(tabs(view)).toEqual(["Home", "Insert", "Design", "Layout", "References", "Review", "View", "Help"]);
    expect(Array.from(view.contentEl.querySelectorAll(".dw-rb-qat .dw-rb-btn")).map((b) => b.getAttribute("data-cmd"))).toEqual(["FileSave", "Undo", "Redo"]);
  });

  it("provides every command Word's layout names", async () => {
    const { v } = await openWord();
    const cmds = v.wordCommands() as Record<string, unknown>;
    expect(WORD_COMMAND_IDS.filter((id) => !cmds[id])).toEqual([]);
  });

  it("Home › Bold bolds the selection and shows as on", async () => {
    const { view, editor } = await openWord();
    cursorIn(editor, "paragraph");
    const { from } = editor.state.selection;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 3)));
    ribbonButton(view, "Bold").click();
    expect(editor.state.doc.rangeHasMark(from, from + 3, editor.state.schema.marks.bold)).toBe(true);
    expect(ribbonButton(view, "Bold").classList.contains("is-active")).toBe(true);
  });

  it("shows Table Design and Layout in a table, Picture Format on a picture", async () => {
    const { view, editor } = await openWord();
    cursorIn(editor, "table_cell");
    expect(tabs(view)).toEqual(["Home", "Insert", "Design", "Layout", "References", "Review", "View", "Help", "Table Design", "Layout"]);
    let image = -1;
    editor.state.doc.descendants((n, pos) => {
      if (image < 0 && n.type.name === "image") image = pos;
      return image < 0;
    });
    editor.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, image)));
    expect(tabs(view)).toContain("Picture Format");
    expect(tabs(view)).not.toContain("Table Design");
  });

  it("shows the page and word count in the status bar, and can hide the word count", async () => {
    const { view } = await openWord();
    const bar = view.contentEl.querySelector<HTMLElement>(".dw-statusbar")!;
    expect(bar.querySelector('[data-item="page"]')!.textContent).toMatch(/^Page \d+ of \d+$/);
    expect(bar.querySelector('[data-item="words"]')!.textContent).toMatch(/^\d[\d,]* words$/);
    bar.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    Menu.last!.items.find((i) => i.title === "Word Count")!.run!();
    expect(bar.querySelector<HTMLElement>('[data-item="words"]')!.hidden).toBe(true);
  });

  it("Review › Track Changes switches to suggesting", async () => {
    const { view, v } = await openWord();
    clickTab(view, "Review");
    ribbonButton(view, "ReviewTrackChanges").click();
    await vi.waitFor(() => expect(v.suggesting).toBe(true));
    expect(view.contentEl.classList.contains("is-suggesting")).toBe(true);
  });

  it("keeps read-only files to the commands that look at them, and never saves", async () => {
    const { view, modify, file } = await openWord({ ...WORD, readOnly: true });
    expect(tabs(view)).toEqual(["Home", "References", "Review", "View", "Help"]);
    clickTab(view, "Review");
    expect(Array.from(view.contentEl.querySelectorAll(".dw-rb-caption-label")).map((c) => c.textContent)).toEqual(["Proofing", "Compare"]);
    expect(Array.from(view.contentEl.querySelectorAll(".dw-rb-qat .dw-rb-btn")).map((b) => b.getAttribute("data-cmd"))).not.toContain("FileSave");
    await view.onUnloadFile(file);
    expect(modify).not.toHaveBeenCalled();
  });

  it("switches layouts on reopen without changing or saving the document", async () => {
    const settings = { ...WORD };
    const { view, v, file, modify } = await openWord(settings);
    const text = v.editor.state.doc.textContent;
    settings.officeDocx = false;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(view.contentEl.querySelector(".dw-rb")).toBeNull();
    expect(view.contentEl.querySelector(".dw-toolbar")).not.toBeNull();
    expect(v.editor.state.doc.textContent).toBe(text);
    expect(modify).not.toHaveBeenCalled();
  });
});

describe("DocxView Word-style right-click", () => {
  const rightClick = (target: Element) => target.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 5, clientY: 5 }));
  const items = () => Array.from(document.querySelectorAll(".dw-ctx > .menu:not(.dw-ctx-sub) > *")).map((e) => (e.classList.contains("menu-separator") ? "---" : e.querySelector(".menu-item-title")!.textContent));
  const item = (label: string) => Array.from(document.querySelectorAll<HTMLElement>(".dw-ctx .menu-item")).find((e) => e.querySelector(".menu-item-title")!.textContent === label)!;
  /** The element holding the text at `pos` (events target elements, not text nodes). */
  const textNode = (editor: import("prosemirror-view").EditorView, pos: number) => {
    const n = editor.domAtPos(pos).node;
    return n.nodeType === Node.TEXT_NODE ? n.parentElement! : n;
  };

  it("shows Word's menu with the mini toolbar for selected text; the mini toolbar's Bold keeps the menu open", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    const { from } = editor.state.selection;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 3)));
    rightClick(textNode(editor, from + 1) as Element);
    expect(items()).toEqual(["Cut", "Copy", "Paste", "Keep Text Only", "---", "Link…", "New Comment", "---", "Paragraph…", "Create Note from Selection"]);
    const mini = Array.from(document.querySelectorAll<HTMLElement>(".dw-ctx-mini [aria-label]")).map((b) => b.getAttribute("aria-label")!.replace(/ \(.*\)$/, ""));
    expect(mini).toEqual(["Font", "Font Size", "Increase Font Size", "Decrease Font Size", "Styles", "Bold", "Italic", "Underline", "Text Highlight Color", "Font Color", "Bullets", "Numbering", "New Comment"]);
    document.querySelector<HTMLElement>('.dw-ctx-mini [aria-label^="Bold"]')!.click();
    expect(editor.state.doc.rangeHasMark(from, from + 3, editor.state.schema.marks.bold)).toBe(true);
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
  });

  it("opens on plain text with no selection too", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    rightClick(textNode(editor, editor.state.selection.from) as Element);
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
    expect(items()[0]).toBe("Cut");
  });

  it("adds Word's table items in a table", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "table_cell");
    rightClick(textNode(editor, editor.state.selection.from) as Element);
    expect(items()).toEqual(expect.arrayContaining(["Insert", "Delete Cells", "Merge Cells", "Split Cells", "Cell Shading"]));
    item("Insert").click();
    expect(Array.from(document.querySelectorAll(".dw-ctx-sub .menu-item-title")).map((t) => t.textContent)).toEqual(["Insert Rows Above", "Insert Rows Below", "Insert Columns to the Left", "Insert Columns to the Right"]);
  });

  it("shows picture items and no mini toolbar on a picture", async () => {
    const { view } = await openWord();
    const img = view.contentEl.querySelector(".docx-image")!;
    rightClick(img.querySelector("img") ?? img);
    expect(items()).toEqual(["Cut", "Copy", "Change Picture…", "Wrap Text", "Crop", "Edit Alt Text…", "Delete"]);
    expect(document.querySelector(".dw-ctx-mini")).toBeNull();
  });

  it("leaves the Google layout's right-click as it was", async () => {
    const { editor } = await openWord({ ...WORD, officeDocx: false });
    cursorIn(editor, "table_cell");
    rightClick(textNode(editor, editor.state.selection.from) as Element);
    expect(document.querySelector(".dw-ctx")).toBeNull();
    expect(Menu.last!.items.map((i) => i.title)).toContain("Insert row above");
  });
});

describe("DocxView Word-style shortcuts", () => {
  /** Press `key` in the editor (ProseMirror reads keydown on its DOM). */
  const press = (editor: import("prosemirror-view").EditorView, key: string, mods: { ctrl?: boolean; shift?: boolean; alt?: boolean } = {}) =>
    editor.dom.dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, altKey: !!mods.alt, bubbles: true, cancelable: true }));
  const align = (editor: import("prosemirror-view").EditorView) => editor.state.selection.$from.parent.attrs.align as string | null;
  const style = (editor: import("prosemirror-view").EditorView) => editor.state.selection.$from.parent.attrs.style as string | null;

  it("align with Ctrl+E / J / R / L", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    press(editor, "e", { ctrl: true });
    expect(align(editor)).toBe("center");
    press(editor, "j", { ctrl: true });
    expect(align(editor)).toBe("justify");
    press(editor, "r", { ctrl: true });
    expect(align(editor)).toBe("right");
    press(editor, "l", { ctrl: true });
    expect(align(editor)).toBe("left");
  });

  it("grow and shrink the font with Ctrl+] and Ctrl+[", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    const { from } = editor.state.selection;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 3)));
    const size = () => editor.state.doc.nodeAt(from)!.marks.find((mk) => mk.type.name === "size")?.attrs.pt as number | undefined;
    press(editor, "]", { ctrl: true });
    const grown = size()!;
    expect(grown).toBeGreaterThan(11);
    press(editor, "[", { ctrl: true });
    expect(size()!).toBeLessThan(grown);
  });

  it("Ctrl+Shift+L makes a bulleted list and Ctrl+Shift+8 does nothing", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    const doc = editor.state.doc;
    press(editor, "8", { ctrl: true, shift: true });
    expect(editor.state.doc.eq(doc)).toBe(true);
    press(editor, "L", { ctrl: true, shift: true });
    expect(editor.state.selection.$from.parent.attrs.numId).not.toBeNull();
  });

  it("Ctrl+Alt+1 sets Heading 1 and Ctrl+Shift+N sets Normal", async () => {
    const { editor } = await openWord();
    cursorIn(editor, "paragraph");
    press(editor, "1", { ctrl: true, alt: true });
    expect(style(editor)).toBe("Heading1");
    press(editor, "N", { ctrl: true, shift: true });
    expect(style(editor)).toBeNull();
  });

  it("Ctrl+Shift+E turns on Track Changes", async () => {
    const { editor, v } = await openWord();
    cursorIn(editor, "paragraph");
    press(editor, "E", { ctrl: true, shift: true });
    await vi.waitFor(() => expect(v.suggesting).toBe(true));
  });

  it("keeps Word's keys from Obsidian's hotkeys only in the Word interface", async () => {
    const { v } = await openWord({ ...WORD });
    const scopeKey = (on: boolean) => {
      const k = (v.scope.keys as { mods: string[]; key: string; fn: () => unknown }[]).find((x) => x.mods.join() === "Mod" && x.key === "j")!;
      v.settings.officeDocx = on;
      document.body.focus();
      v.editor.dom.focus();
      return k.fn();
    };
    // (Registered only while the Word layout is on: see "registers Word's extra keys only while its layout is on".)
    expect(scopeKey(true)).toBe(true);
  });

  it("lists Word's keys in Help › Keyboard Shortcuts", async () => {
    const { view } = await openWord();
    clickTab(view, "Help");
    ribbonButton(view, "KeyboardShortcuts").click();
    const text = document.querySelector(".test-modal")!.textContent!;
    expect(text).toContain("Ctrl+E");
    expect(text).toContain("Ctrl+Shift+L");
    expect(text).not.toContain("Ctrl+Shift+8");
  });

  it("keeps Google's keys in the Google layout", async () => {
    const { editor } = await openWord({ ...WORD, officeDocx: false });
    cursorIn(editor, "paragraph");
    press(editor, "8", { ctrl: true, shift: true });
    expect(editor.state.selection.$from.parent.attrs.numId).not.toBeNull();
  });
});

describe("DocxView Word-style font boxes", () => {
  it("show the cursor's font and size as soon as the file opens", async () => {
    const { view } = await openWord();
    const font = view.contentEl.querySelector('.dw-rb-select[data-cmd="Font"] .dw-rb-select-label')!.textContent;
    const size = view.contentEl.querySelector('.dw-rb-select[data-cmd="FontSize"] .dw-rb-select-label')!.textContent;
    expect(font).not.toBe("");
    expect(size).toMatch(/^\d+(\.\d+)?$/);
  });
});

describe("DocxView Word-style status bar pages", () => {
  it("updates Page N of M when the page layout is measured", async () => {
    const { view, editor } = await openWord();
    const { pagesKey } = await import("../src/presentation/docx/docx-pages");
    const { Decoration, DecorationSet } = await import("prosemirror-view");
    const decorations = DecorationSet.create(editor.state.doc, [Decoration.widget(1, () => document.createElement("span"))]);
    editor.dispatch(editor.state.tr.setMeta(pagesKey, { pages: 3, decorations, bands: [0, 0], pageStarts: [{ pos: 3, page: 2 }, { pos: 6, page: 3 }] }));
    expect(view.contentEl.querySelector('.dw-statusbar [data-item="page"]')!.textContent).toBe("Page 1 of 3");
  });
});

describe("DocxView Word-style review fixes", () => {
  it("a reload after an outside change rebuilds the commands and drops the old ribbon", async () => {
    const { view, v, file } = await openWord({ ...WORD });
    const oldCmds = v.cmds, oldRibbon = v.ribbon;
    const destroy = vi.spyOn(oldRibbon, "destroy");
    await view.onLoadFile(file);
    expect(destroy).toHaveBeenCalled();
    expect(v.cmds).not.toBe(oldCmds);
    expect(view.contentEl.querySelectorAll(".dw-rb").length).toBe(1);
  });

  it("closes its own open right-click menu when the file closes", async () => {
    const { view, editor, file } = await openWord({ ...WORD });
    cursorIn(editor, "paragraph");
    const n = editor.domAtPos(editor.state.selection.from).node;
    (n.nodeType === Node.TEXT_NODE ? n.parentElement! : (n as Element)).dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 5, clientY: 5 }));
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
    await view.onUnloadFile(file);
    expect(document.querySelector(".dw-ctx")).toBeNull();
  });

  it("Ctrl+Shift+V does nothing (leaving the key to the app) when no formatting was copied", async () => {
    const { editor } = await openWord({ ...WORD });
    cursorIn(editor, "paragraph");
    const e = new KeyboardEvent("keydown", { key: "V", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    editor.dom.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(false);
  });

  it("forgets the Google toolbar after switching to the Word layout", async () => {
    const settings = { ...WORD, officeDocx: false };
    const { view, v, file } = await openWord(settings);
    settings.officeDocx = true;
    await view.onUnloadFile(file);
    await view.onLoadFile(file);
    expect(v.toolbar).toBeNull();
    expect(v.fontSelect).toBeNull();
  });
});

describe("DocxView Word-style screenshot fixes", () => {
  it("draws each Styles tile in its style's look", async () => {
    const { view } = await openWord({ ...WORD });
    const tile = (v: string) => view.contentEl.querySelector<HTMLElement>(`.dw-rb-gallery-item[data-value="${v}"]`)!;
    expect(tile("Heading1").style.fontWeight).toBe("bold");
    expect(parseFloat(tile("Heading1").style.fontSize)).toBeGreaterThan(parseFloat(tile("Normal").style.fontSize));
  });

  it("leaves out the page count in Web Layout", async () => {
    const { view, v } = await openWord({ ...WORD });
    v.setPageView(false);
    expect(view.contentEl.querySelector('.dw-statusbar [data-item="page"]')!.textContent).toBe("");
  });

  it("double-clicking Format Painter keeps painting", async () => {
    const { view, v, editor } = await openWord({ ...WORD });
    cursorIn(editor, "paragraph");
    view.contentEl.querySelector<HTMLElement>('.dw-rb-panel [data-cmd="FormatPainter"]')!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(v.painter?.sticky).toBe(true);
  });
});

describe("DocxView Word-style keys registration", () => {
  it("registers Word's extra keys only while its layout is on", async () => {
    const keys = (v: { scope: { keys: { mods: string[]; key: string }[] } }) => v.scope.keys.filter((k) => k.mods.join() === "Mod" && k.key === "]");
    const google = await openWord({ ...WORD, officeDocx: false });
    expect(keys(google.v)).toEqual([]);
    const { v, view, file } = await openWord({ ...WORD });
    expect(keys(v).length).toBe(1);
    await view.onUnloadFile(file);
    expect(keys(v)).toEqual([]);
  });

  it("closing one file leaves a menu another view opened", async () => {
    const { showContextMenu } = await import("../src/presentation/common/context-menu");
    const { view, file } = await openWord({ ...WORD });
    showContextMenu({ x: 0, y: 0 }, [{ label: "Elsewhere" }]);
    await view.onUnloadFile(file);
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
  });
});
