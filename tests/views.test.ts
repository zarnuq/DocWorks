import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, localData, Menu, Modal, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DEFAULT_SETTINGS, DocxView, DocWorksSettings } from "../src/presentation/docx/docx-view";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { readAnimations, readTransition } from "../src/domain/pptx/animation";
import { addSlideComment } from "../src/domain/pptx/comments";
import { writeFill } from "../src/domain/pptx/shape-style";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
import { fixture } from "./helpers";

/** A view in a leaf attached to the document, reading `name` from tests/fixtures for every file. */
function mount<V extends PptxView | DocxView>(View: new (leaf: RealLeaf) => V, name: string): V {
  const app = new App();
  app.vault.readBinary = async () => fixture(name).buffer as ArrayBuffer;
  const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf);
  document.body.appendChild(view.containerEl);
  return view;
}

async function open(view: PptxView | DocxView, path: string) {
  const file = new TFile(path) as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  return file;
}

afterEach(() => {
  document.body.empty();
  vi.useRealTimers();
  vi.restoreAllMocks();
});

/** Open a menu in the view's menu row; returns its item titles (separators left out). */
function menuTitles(view: { contentEl: HTMLElement }, name: string): string[] {
  const b = Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!;
  b.click();
  return Menu.last!.items.map((i) => i.title).filter((t) => t !== "---");
}
/** Run a menu item by path, e.g. runMenu(view, "Format", "Text", "Bold"). */
function runMenu(view: { contentEl: HTMLElement }, name: string, ...path: string[]): void {
  menuTitles(view, name);
  let menu = Menu.last!;
  for (const title of path) {
    const item = menu.items.find((i) => i.title === title);
    if (!item) throw new Error(`No "${title}" in ${name}: ${menu.items.map((i) => i.title).join(", ")}`);
    if (item.submenu) menu = item.submenu;
    else item.run!(new MouseEvent("click"));
  }
}
const menuNames = (view: { contentEl: HTMLElement }) => Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent);

describe("PptxView Copy as Markdown outline", () => {
  it("copies the deck as headings and bullets from the Edit menu", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "deck.pptx");
    const written: string[] = [];
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (t: string) => void written.push(t) } });
    runMenu(view, "Edit", "Copy as Markdown outline");
    await Promise.resolve();
    expect(written[0]).toMatch(/^## Hello Deck\n\n- Subtitle here\n/);
  });
});

describe("PptxView motion editor", () => {
  it("keeps slide creation in the Slide dropdown without a duplicate toolbar button", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "motion.pptx");
    const labels = Array.from(view.contentEl.querySelectorAll(".dw-toolbar button")).map((b) => b.textContent);
    expect(menuTitles(view, "Slide")).toContain("New slide");
    expect(labels).not.toContain("New slide");
    await view.onUnloadFile(file);
  });
  it("adds animations from the toolbar, supports undo/redo and persists them for present mode", async () => {
    const view = mount(PptxView, "sample.pptx");
    const saved = vi.spyOn(view.app.vault, "modifyBinary");
    const file = await open(view, "motion.pptx");
    const v = view as any;
    v.setSelection([v.rendered.find((r: any) => !r.shape.inherited && r.shape.kind === "sp").shape.el]);
    const button = (name: string) => Array.from(view.contentEl.querySelectorAll("button")).find((b) => b.textContent === name)!;
    runMenu(view, "Slide", "Motion…");
    expect(saved).not.toHaveBeenCalled();
    expect(button("Add animation").disabled).toBe(false);
    button("Add animation").click();
    expect(readAnimations(v.currentSlideXml())).toHaveLength(1);
    expect(view.contentEl.querySelector(".pptx-props-heading")!.textContent).toBe("Motion");
    v.undo();
    expect(readAnimations(v.currentSlideXml())).toEqual([]);
    v.redo();
    expect(readAnimations(v.currentSlideXml())).toHaveLength(1);
    await view.requestSave();
    expect(saved).toHaveBeenCalled();
    const bytes = saved.mock.calls.at(-1)![1];
    const reloaded = await PptxDocument.load(bytes);
    expect(readAnimations(reloaded.xml(reloaded.slides()[0].path)!)).toHaveLength(1);
    view.startPresentation(false);
    const overlay = document.querySelector<HTMLElement>(".dw-show")!;
    expect(overlay.querySelector<HTMLElement>(".pptx-shape[style*='hidden']")).not.toBeNull();
    overlay.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    overlay.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await view.onUnloadFile(file);
  });

  it("sets transitions, applies only their visual settings to all slides and undoes in one step", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "motion.pptx");
    const v = view as any;
    v.openMotion();
    const field = view.contentEl.querySelector<HTMLSelectElement>('[data-field="Transition"]')!;
    field.value = "push";
    field.dispatchEvent(new Event("change"));
    expect(readTransition(v.currentSlideXml())).toMatchObject({ type: "push", dur: 500 });
    Array.from(view.contentEl.querySelectorAll("button")).find((b) => b.textContent === "Apply to all slides")!.click();
    expect(v.doc.slides().every((s: any) => readTransition(v.doc.xml(s.path))?.type === "push")).toBe(true);
    v.undo();
    expect(readTransition(v.doc.xml(v.doc.slides()[1].path))).toBeNull();
    await view.onUnloadFile(file);
  });
});

describe("PptxView switching files", () => {
  it("Ctrl+F opens find again after switching files with find open", async () => {
    const view = mount(PptxView, "sample.pptx");
    const a = await open(view, "a.pptx");
    (view as any).openFind();
    await view.onUnloadFile(a);
    await open(view, "b.pptx");
    (view as any).openFind();
    const bar = view.contentEl.querySelector(".pptx-find");
    expect(bar?.isConnected).toBe(true);
  });

  it("the shape gallery opens after switching files, and its window listener is removed", async () => {
    const view = mount(PptxView, "sample.pptx");
    const a = await open(view, "a.pptx");
    const anchor = () => view.contentEl.querySelector<HTMLElement>(".dw-toolbar button")!;
    (view as any).toggleGallery(anchor());
    const removed = vi.spyOn(window, "removeEventListener");
    await view.onUnloadFile(a);
    expect(removed).toHaveBeenCalledWith("pointerdown", expect.any(Function), { capture: true });
    await open(view, "b.pptx");
    (view as any).toggleGallery(anchor());
    expect(view.contentEl.querySelector(".pptx-gallery")?.isConnected).toBe(true);
  });
});

describe("PptxView theme", () => {
  it("opens with the dark workspace by default", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    expect(view.contentEl.classList.contains("pptx-theme-dark")).toBe(true);
  });
});

describe("PptxView slide copy and paste", () => {
  it("copies a slide in one deck and pastes it into another, from the slide menu and with Ctrl+C / Ctrl+V on the thumbnails", async () => {
    const a = mount(PptxView, "sample.pptx");
    await open(a, "a.pptx");
    const b = mount(PptxView, "sample.pptx");
    await open(b, "b.pptx");
    const va = a as any, vb = b as any;
    const count = vb.doc.slides().length;
    // Menu: Copy slide on A's second slide, Paste on B's first.
    const MenuStub = (await import("./obsidian-stub")).Menu as any;
    const items: any[] = [];
    vi.spyOn(MenuStub.prototype, "addItem").mockImplementation(function (this: any, cb: any) {
      const item: any = { title: "", onClick: (f: any) => ((item.run = f), item), setTitle: (t: string) => ((item.title = t), item), setIcon: () => item, setChecked: () => item, setDisabled: () => item };
      cb(item);
      items.push(item);
      return this;
    });
    va.slideMenu(1);
    expect(items.map((i) => i.title)).toEqual(expect.arrayContaining(["Copy slide", "Cut slide"]));
    items.find((i) => i.title === "Copy slide").run();
    await va.copying;
    items.length = 0;
    vb.slideMenu(0);
    await items.find((i) => i.title === "Paste slide").run();
    expect(vb.doc.slides()).toHaveLength(count + 1);
    expect(vb.index).toBe(1);
    // Keyboard on the thumbnails: Ctrl+C in A, Ctrl+V in B.
    va.thumbsEl.focus();
    const keys = (v: any) => v.scope.keys as { mods: string[]; key: string; fn: () => boolean | void }[];
    expect(keys(va).find((k) => k.key === "c" && k.mods.length === 1)!.fn()).toBe(false);
    await va.copying;
    vb.thumbsEl.focus();
    expect(keys(vb).find((k) => k.key === "v" && k.mods.length === 1)!.fn()).toBe(false);
    expect(vb.doc.slides()).toHaveLength(count + 2);
  });
});

describe("PptxView thumbnails", () => {
  it("has a + button under the last thumbnail that adds a slide at the end", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    const count = v.doc.slides().length;
    const add = view.contentEl.querySelector<HTMLElement>(".pptx-thumb-add")!;
    expect(add.getAttribute("aria-label")).toBe("New slide");
    add.click();
    expect(v.doc.slides().length).toBe(count + 1);
    expect(v.index).toBe(count);
  });
});

describe("PptxView deleting slides from the thumbnails", () => {
  it("deletes the current slide with Delete or Backspace, but never the last one", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    v.mutate((d: any) => d.addSlide(0));
    const count = v.doc.slides().length;
    v.thumbsEl.focus();
    v.thumbsEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    expect(v.doc.slides()).toHaveLength(count - 1);
    v.thumbsEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", bubbles: true }));
    expect(v.doc.slides()).toHaveLength(count - 2);
    while (v.doc.slides().length > 1) v.thumbsEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    v.thumbsEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    expect(v.doc.slides()).toHaveLength(1);
    expect(v.index).toBe(0);
  });
});

describe("PptxView empty text boxes", () => {
  it("puts the caret inside the empty line when an empty placeholder is clicked", async () => {
    const { blankPptx } = await import("../src/infrastructure/blank");
    const bytes = await blankPptx();
    const app = new App();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    await open(view, "blank.pptx");
    const textEl = (view as any).rendered.find((r: any) => r.textEl?.classList.contains("is-empty")).textEl as HTMLElement;
    textEl.focus();
    // Clicking the prompt text leaves the browser's caret on the box itself, before its paragraph.
    getSelection()!.collapse(textEl, 0);
    textEl.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(getSelection()!.anchorNode).toBe(textEl.querySelector("p"));
    // Or nowhere in the box at all, e.g. when the click first selected the shape.
    getSelection()!.collapse(document.body, 0);
    textEl.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(getSelection()!.anchorNode).toBe(textEl.querySelector("p"));
  });

  it("does the same for an empty shape that isn't a placeholder (a new rectangle)", async () => {
    const { blankPptx } = await import("../src/infrastructure/blank");
    const { PptxDocument } = await import("../src/infrastructure/pptx/pptx-document");
    const { addShape } = await import("../src/domain/pptx/slide-edit");
    const deck = await PptxDocument.load(await blankPptx());
    addShape(deck.xml(deck.slides()[0].path)!, "rect", { x: 100000, y: 100000, w: 2000000, h: 1000000 });
    const bytes = await deck.save();
    const app = new App();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    await open(view, "rect.pptx");
    const rect = (view as any).rendered.find((r: any) => r.shape.name.startsWith("Shape "));
    const textEl = rect.textEl as HTMLElement;
    expect(textEl.textContent).toBe("");
    getSelection()!.collapse(document.body, 0);
    textEl.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(getSelection()!.anchorNode).toBe(textEl.querySelector("p"));
  });

  it("drags an empty shape by its body (it has no text to select yet)", async () => {
    const { blankPptx } = await import("../src/infrastructure/blank");
    const { PptxDocument } = await import("../src/infrastructure/pptx/pptx-document");
    const { addShape } = await import("../src/domain/pptx/slide-edit");
    const deck = await PptxDocument.load(await blankPptx());
    addShape(deck.xml(deck.slides()[0].path)!, "ellipse", { x: 100000, y: 100000, w: 2000000, h: 1000000 });
    const bytes = await deck.save();
    const app = new App();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    await open(view, "drag.pptx");
    const v = view as any;
    const move = vi.spyOn(v, "startMoveItems").mockImplementation(() => undefined);
    const shape = v.rendered.find((r: any) => r.shape.name.startsWith("Shape "));
    shape.textEl.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
    expect(move).toHaveBeenCalled();
  });
});

describe("PptxView comments", () => {
  it("posts, replies to and resolves a comment thread on the current slide", async () => {
    const { slideComments } = await import("../src/domain/pptx/comments");
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    v.commentsDialog();
    const modal = () => Array.from(document.querySelectorAll(".test-modal")).pop()!;
    const post = (label: string, text: string) => {
      const box = modal().querySelector<HTMLTextAreaElement>(`textarea[aria-label='${label}']`)!;
      box.value = text;
      (box.nextElementSibling as HTMLButtonElement).click();
    };
    post("New comment", "Tighten this slide");
    post("Reply", "On it");
    Array.from(modal().querySelectorAll("button")).find((b) => b.textContent === "Resolve")!.click();
    const [t] = slideComments(v.doc, v.doc.slides()[0].path);
    expect(t.messages.map((m: { text: string }) => m.text)).toEqual(["Tighten this slide", "On it"]);
    expect(t.resolved).toBe(true);
    expect(modal().textContent).toContain("Resolved");
    v.undo();
    expect(slideComments(v.doc, v.doc.slides()[0].path)[0].resolved).toBe(false);
  });
});

describe("PptxView notes", () => {
  it("sizes the notes box to the slide's width so the two line up on wide screens", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    Object.defineProperty(v.stageEl, "clientWidth", { get: () => 2400, configurable: true });
    v.renderCurrent();
    expect(v.stageEl.style.getPropertyValue("--pptx-slide-width")).toBe("1400px");
  });
});

describe("PptxView fit", () => {
  it("refits a slide that was resized while its text was being edited, once editing ends", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    let width = 1200;
    Object.defineProperty(v.stageEl, "clientWidth", { get: () => width, configurable: true });
    v.renderCurrent();
    const before = v.scale;
    const rs = v.rendered.find((r: any) => r.textEl);
    rs.textEl.focus();
    v.activeText = rs;
    width = 800; // e.g. the format panel opened
    rs.textEl.dispatchEvent(new FocusEvent("blur"));
    expect(v.scale).toBeLessThan(before);
  });
});

describe("view preferences", () => {
  it("keep the panels' open state in the vault's storage, not the browser's", async () => {
    const view = mount(DocxView, "sample.docx");
    await open(view, "a.docx");
    (view as any).setReview(true);
    expect(localData.get("doc-works-review-panel")).toBe("on");
    expect(window.localStorage.getItem("doc-works-review-panel")).toBeNull();
    expect((mount(DocxView, "sample.docx") as any).reviewOpen).toBe(true);
    localData.clear();
  });
});

describe("DocxView switching files", () => {
  it("pending word-count and review timers don't fire after unload", async () => {
    localData.set("doc-works-review-panel", "on");
    const view = mount(DocxView, "sample.docx");
    const a = await open(view, "a.docx");
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const words = vi.spyOn(view as any, "updateWordCount");
    const review = vi.spyOn(view as any, "renderReview");
    (view as any).scheduleWordCount();
    (view as any).scheduleReview();
    await view.onUnloadFile(a);
    expect(vi.getTimerCount()).toBe(0);
    vi.runAllTimers();
    expect(words).not.toHaveBeenCalled();
    expect(review).not.toHaveBeenCalled();
    localData.delete("doc-works-review-panel");
  });
});

describe("DocxView find", () => {
  it("finds comment replies, opens their review cards and leaves comments unchanged by replace", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "comments.docx");
    const v = view as any;
    const id = v.docx.addComment('needle root', 'Author');
    const reply = v.docx.replyToComment(id, 'reply needle', 'Reply author');
    v.openFind();
    const bar = view.contentEl.querySelector<HTMLElement>('.pptx-find')!;
    const input = bar.querySelector<HTMLInputElement>('input[type=search]')!;
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('1 of 2');
    expect(bar.querySelector('.dw-find-preview mark')!.textContent).toBe('needle');
    const buttons = () => Array.from(bar.querySelectorAll<HTMLButtonElement>('button'));
    expect(buttons().find(b => b.textContent === 'Replace')!.disabled).toBe(true);
    expect(buttons().find(b => b.textContent === 'Replace all')!.disabled).toBe(true);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    expect(bar.querySelector('.dw-find-source')!.textContent).toContain('Reply author');
    buttons().find(b => b.textContent === 'Open comment')!.click();
    const area = view.contentEl.querySelector<HTMLTextAreaElement>(`[data-comment-id="${reply}"] textarea`)!;
    expect(document.activeElement).toBe(area);
    expect(area.value.slice(area.selectionStart, area.selectionEnd)).toBe('needle');
    expect(v.docx.comments().filter((c: any) => c.id === id || c.id === reply).map((c: any) => c.text)).toEqual(['needle root', 'reply needle']);
    bar.querySelector<HTMLInputElement>('[aria-label="Search comments"]')!.click();
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('No results');
    expect(bar.querySelector('.dw-find-preview')!.textContent).toBe('');
    await view.onUnloadFile(file);
  });
  it("opens the find bar above the document (Ctrl+F / Ctrl+H)", async () => {
    const view = mount(DocxView, "sample.docx");
    await open(view, "a.docx");
    (view as any).openFind();
    const bar = view.contentEl.querySelector(".pptx-find");
    expect(bar?.isConnected).toBe(true);
    expect(document.activeElement).toBe(bar!.querySelector("input"));
  });
});

describe("PptxView annotation find", () => {
  it("highlights repeated notes in place and clears them on no results, close and editing", async () => {
    const view = mount(PptxView, 'sample.pptx');
    const file = await open(view, 'highlight-notes.pptx');
    const v = view as any;
    v.doc.setNotes(0, 'needle first\nsecond needle');
    v.renderCurrent();
    const before = v.doc.snapshot();
    v.openFind();
    const input = view.contentEl.querySelector<HTMLInputElement>('.pptx-find input[type=search]')!;
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    const mirror = view.contentEl.querySelector<HTMLElement>('.pptx-notes-highlight')!;
    expect(mirror.getAttribute('aria-hidden')).toBe('true');
    expect(mirror.textContent).toBe(v.notesEl.value + '\n');
    expect(mirror.querySelector('mark')!.previousSibling?.textContent ?? '').toBe('');
    expect(document.activeElement).toBe(input);
    v.renderCurrent();
    expect(mirror.querySelector('mark')!.textContent).toBe('needle');
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(mirror.querySelector('mark')!.previousSibling!.textContent).toBe('needle first\nsecond ');
    expect(v.doc.snapshot()).toEqual(before);
    input.value = 'absent'; input.dispatchEvent(new Event('input'));
    expect(mirror.querySelector('mark')).toBeNull();
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    v.notesEl.value = 'changed notes'; v.notesEl.dispatchEvent(new Event('input'));
    expect(mirror.querySelector('mark')).toBeNull();
    v.doc.setNotes(0, 'needle'); v.renderCurrent();
    input.dispatchEvent(new Event('input'));
    v.closeFind();
    expect(mirror.querySelector('mark')).toBeNull();
    await view.onUnloadFile(file);
  });

  it("searches speaker-note edits before their delayed save", async () => {
    const view = mount(PptxView, 'sample.pptx');
    const file = await open(view, 'notes.pptx');
    const v = view as any;
    v.notesEl.value = 'pending needle';
    v.notesEl.dispatchEvent(new Event('input'));
    v.openFind();
    const input = view.contentEl.querySelector<HTMLInputElement>('.pptx-find input[type=search]')!;
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    expect(view.contentEl.querySelector('.pptx-notes-highlight mark')!.textContent).toBe('needle');
    expect(view.contentEl.querySelector('.dw-find-preview')!.textContent).toBe('');
    expect(v.doc.notesText(0)).toBe('pending needle');
    await view.onUnloadFile(file);
  });
  it("includes notes and replies by default and keeps keyboard navigation in the find box", async () => {
    const view = mount(PptxView, 'sample.pptx');
    const file = await open(view, 'annotations.pptx');
    const v = view as any;
    v.doc.setNotes(0, 'speaker needle');
    v.doc.setNotes(1, 'second needle');
    const path = v.doc.slides()[0].path;
    const id = addSlideComment(v.doc, path, 'comment needle', 'Author');
    addSlideComment(v.doc, path, 'reply needle', 'Reply author', id);
    v.renderCurrent();
    v.openFind();
    const bar = view.contentEl.querySelector<HTMLElement>('.pptx-find')!;
    const input = bar.querySelector<HTMLInputElement>('input[type=search]')!;
    input.value = 'needle'; input.dispatchEvent(new Event('input'));
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('1 of 4');
    expect(view.contentEl.querySelector('.pptx-notes-highlight mark')!.textContent).toBe('needle');
    expect(bar.querySelector('.dw-find-preview')!.textContent).toBe('');
    expect(Array.from(bar.querySelectorAll<HTMLButtonElement>('button')).find(b => b.textContent === 'Replace')!.disabled).toBe(true);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(bar.querySelector('.dw-find-source')!.textContent).toContain('Comment');
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(bar.querySelector('.dw-find-source')!.textContent).toContain('Reply author');
    Array.from(bar.querySelectorAll<HTMLButtonElement>('button')).find(b => b.textContent === 'Open comment')!.click();
    expect(document.querySelector('.xlsx-comment-text mark')!.textContent).toBe('needle');
    Array.from(document.querySelectorAll<HTMLButtonElement>('.test-modal button')).find(b => b.textContent === 'Close')!.click();
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(v.index).toBe(1);
    expect(view.contentEl.querySelector('.pptx-notes-highlight mark')!.textContent).toBe('needle');
    expect(bar.querySelector('.dw-find-preview')!.textContent).toBe('');
    expect(document.activeElement).toBe(input);
    bar.querySelector<HTMLInputElement>('[aria-label="Search speaker notes"]')!.click();
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('1 of 2');
    bar.querySelector<HTMLInputElement>('[aria-label="Search comments"]')!.click();
    expect(bar.querySelector('.pptx-find-count')!.textContent).toBe('No results');
    await view.onUnloadFile(file);
  });
});

describe("editing keys Obsidian hotkeys could take", () => {
  const shielded = (view: any, mods: string[], key: string) => view.scope.keys.find((k: any) => k.key === key && k.mods.join() === mods.join());
  it("keeps Shift+Tab and Mod+B/I/K... for the editors when one has focus", async () => {
    for (const [View, file] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
      const view = mount(View as any, file);
      for (const [mods, key] of [[["Shift"], "Tab"], [["Mod"], "b"], [["Mod"], "i"]] as const) expect(shielded(view, [...mods], key), `${View.name} ${mods}+${key}`).toBeTruthy();
      const editable = document.body.createDiv({ attr: { contenteditable: "true" } });
      editable.focus();
      const k = shielded(view, ["Shift"], "Tab");
      // Not handled by the view (true): the key goes on to the focused editor, and no app hotkey takes it.
      expect(k.fn(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true }))).toBe(true);
      editable.remove();
    }
  });
});

describe("DocxView page size setting", () => {
  it("lays out an A4 document on Letter pages when the setting says so, and switches live", async () => {
    const docx = await DocxDocument.load(fixture("sample.docx"));
    // Every section on A4 (the sample has several).
    for (const pgSz of Array.from(docx.body.getElementsByTagNameNS(W, "pgSz"))) {
      pgSz.setAttributeNS(W, "w:w", "11906");
      pgSz.setAttributeNS(W, "w:h", "16838");
    }
    const bytes = await docx.save();
    const settings: DocWorksSettings = { ...DEFAULT_SETTINGS, pageSize: "letter" as const, author: "" };
    const app = new App();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
    document.body.appendChild(view.containerEl);
    await open(view, "a4.docx");
    const host = () => view.contentEl.querySelector<HTMLElement>(".docx-editor-host")!;
    const v = view as any;
    expect(v.setup.width).toBe(816);
    expect(host().style.getPropertyValue("--docx-page-width")).toBe(`${816 - v.setup.left - v.setup.right}px`);
    settings.pageSize = "document";
    view.applyPageSize();
    expect(v.setup.width).toBe(794);
    expect(host().style.getPropertyValue("--docx-page-width")).toBe(`${794 - v.setup.left - v.setup.right}px`);
  });
});

describe("Toolbar groups", () => {
  it("puts the controls between gaps into one boxed group", async () => {
    const { Toolbar } = await import("../src/presentation/common/toolbar");
    const tb = new Toolbar(document.body.createDiv());
    tb.button("B", "Bold", () => {});
    tb.button("I", "Italic", () => {});
    tb.gap();
    tb.select("Font", "Font", [["Arial", "Arial"]], () => {});
    tb.gap();
    tb.end();
    const groups = Array.from(tb.el.querySelectorAll(".dw-group")).filter((g) => g.childElementCount);
    expect(groups.map((g) => g.childElementCount)).toEqual([2, 1]);
    expect(tb.el.lastElementChild!.classList.contains("dw-end")).toBe(true);
    expect(tb.el.querySelector(".dw-segmented")).toBeNull();
  });

  it("offers an icon dropdown that still uses the native select", async () => {
    const { Toolbar } = await import("../src/presentation/common/toolbar");
    const tb = new Toolbar(document.body.createDiv());
    const picked: string[] = [];
    const s = tb.iconSelect("unfold-vertical", "Line spacing", [["1.5", "1.5"]], (v) => picked.push(v));
    expect(s.closest(".dw-icon-select")?.getAttribute("aria-label")).toBe("Line spacing");
    s.value = "1.5";
    s.dispatchEvent(new Event("change"));
    expect(picked).toEqual(["1.5"]);
  });

  it("shows the current value: ticks a matching option, or names a value that isn't listed", async () => {
    const { Toolbar, showValue } = await import("../src/presentation/common/toolbar");
    const tb = new Toolbar(document.body.createDiv());
    const s = tb.select("Font", "Font", [["Arial", "Arial"], ["Calibri", "Calibri"]], () => {});
    showValue(s, "Calibri");
    expect(s.value).toBe("Calibri");
    showValue(s, "Garamond");
    expect([s.value, s.options[s.selectedIndex].text]).toEqual(["", "Garamond"]);
    showValue(s, null);
    expect([s.value, s.options[s.selectedIndex].text]).toEqual(["", "Font"]);
  });

  it("shows a select's placeholder only on the closed control, not as a pickable option", async () => {
    const { Toolbar } = await import("../src/presentation/common/toolbar");
    const tb = new Toolbar(document.body.createDiv());
    const s = tb.select("Font", "Font", [["Arial", "Arial"]], () => {});
    expect(s.value).toBe("");
    expect(s.options[0].text).toBe("Font");
    expect(s.options[0].hidden).toBe(true);
  });
});

describe("DocxView drawings", () => {
  it("inserts native shapes, edits their text and retains them after saving", async () => {
    const view = mount(DocxView, "sample.docx");
    const saved = vi.spyOn(view.app.vault, "modifyBinary");
    const file = await open(view, "drawing.docx");
    const opened = vi.spyOn(Modal.prototype, "open");
    const button = (host: HTMLElement, label: string) => Array.from(host.querySelectorAll("button")).find(b => b.textContent === label)!;
    runMenu(view, "Insert", "Drawing");
    let modal = opened.mock.contexts.at(-1) as Modal;
    button(modal.contentEl, "Add shape").click();
    const text = modal.contentEl.querySelector<HTMLTextAreaElement>('[aria-label="Shape text"]')!;
    text.value = "Native drawing"; text.dispatchEvent(new Event("input"));
    button(modal.contentEl, "Insert").click();
    const drawing = view.contentEl.querySelector(".docx-drawing")!;
    expect(drawing.textContent).toContain("Native drawing");
    drawing.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    modal = opened.mock.contexts.at(-1) as Modal;
    const edit = modal.contentEl.querySelector<HTMLTextAreaElement>('[aria-label="Shape text"]')!;
    edit.value = "Edited drawing"; edit.dispatchEvent(new Event("input"));
    button(modal.contentEl, "Save").click();
    await view.requestSave();
    const back = await DocxDocument.load(saved.mock.calls.at(-1)![1]);
    expect(back.body.getElementsByTagNameNS(W, "txbxContent")[0].textContent).toBe("Edited drawing");
    await view.onUnloadFile(file);
  });
});

describe("DocxView header and footer", () => {
  it("opens the header and footer editor on a double-click of the header or footer band", async () => {
    const view = mount(DocxView, "sample.docx");
    await open(view, "a.docx");
    const dialog = vi.spyOn(view as any, "headerFooterDialog").mockImplementation(() => {});
    const host = view.contentEl.querySelector<HTMLElement>(".docx-editor-host")!;
    const band = host.querySelector(".ProseMirror")!.appendChild(createDiv({ cls: "docx-page-head" }));
    band.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(dialog).toHaveBeenCalledTimes(1);
    host.querySelector(".ProseMirror p")?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    // The blank rest of the last page isn't the footer.
    const last = host.querySelector(".ProseMirror")!.appendChild(createDiv({ cls: "docx-page-last" }));
    last.createDiv({ cls: "docx-page-fill" }).dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(dialog).toHaveBeenCalledTimes(1);
  });
});

describe("DocxView header and footer suggestions", () => {
  it("keeps a header made in the header dialog when the document is saved", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
    const saved = vi.spyOn(app.vault, "modifyBinary");
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, pageSize: "document" as const, author: "Ada" });
    document.body.appendChild(view.containerEl);
    await open(view, "hf.docx");
    const v = view as any;
    expect(v.docx.headerFooter("header", "default")).toBeNull();
    const opened = vi.spyOn(Modal.prototype, "open");
    v.headerFooterDialog();
    const modal = opened.mock.contexts[0] as Modal;
    const header = (modal as any).editors.get("header:default").view;
    header.dispatch(header.state.tr.insertText("Title", 1));
    modal.close();
    v.editor.dispatch(v.editor.state.tr.insertText("x", 1));
    await view.requestSave();
    const back = await DocxDocument.load(saved.mock.calls.at(-1)![1]);
    expect(back.headerFooter("header", "default")?.textContent).toBe("Title");
  });

  it("records header edits as suggestions while suggesting, lists them for review and accepts or rejects them", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, pageSize: "document" as const, author: "Ada" });
    document.body.appendChild(view.containerEl);
    await open(view, "hf.docx");
    const v = view as any;
    v.suggesting = true;
    const opened = vi.spyOn(Modal.prototype, "open");
    v.headerFooterDialog();
    const modal = opened.mock.contexts[0] as Modal;
    const header = (modal as any).editors.get("header:default").view;
    header.dispatch(header.state.tr.insertText("Draft ", 1));
    modal.close();
    const root = v.docx.headerFooter("header", "default") as Element;
    expect(root.getElementsByTagNameNS(W, "ins")[0]?.textContent).toBe("Draft ");
    v.reviewOpen = true;
    v.renderReview();
    const review = view.contentEl.querySelector(".docx-review")!;
    const card = Array.from(review.querySelectorAll(".docx-review-card")).find(c => c.textContent!.includes("Draft"))!;
    expect(card.textContent).toContain("Header");
    Array.from(card.querySelectorAll("button")).find(b => b.textContent === "Reject")!.click();
    expect(v.docx.headerFooter("header", "default").getElementsByTagNameNS(W, "ins")).toHaveLength(0);
    expect(v.docx.headerFooter("header", "default").textContent).not.toContain("Draft");
  });
});

describe("File menu (⋮ More options)", () => {
  it("has Obsidian's file items (rename, delete, other plugins' file-menu entries) plus Export…, in every editor", async () => {
    for (const [View, fixtureName, name] of [[DocxView, "sample.docx", "a.docx"], [PptxView, "sample.pptx", "a.pptx"]] as const) {
      const view = mount(View as any, fixtureName) as any;
      const file = await open(view, name);
      const app = view.app;
      app.fileManager = { promptForFileRename: vi.fn(), promptForDeletion: vi.fn() };
      app.workspace.trigger = vi.fn();
      const menu = new Menu();
      view.onPaneMenu(menu, "more-options");
      const titles = menu.items.map((i) => i.title);
      expect(titles).toEqual(expect.arrayContaining(["Rename…", "Delete file", "Export…"]));
      expect(app.workspace.trigger).toHaveBeenCalledWith("file-menu", menu, file, "more-options", view.leaf);
      menu.items.find((i) => i.title === "Rename…")!.run!();
      expect(app.fileManager.promptForFileRename).toHaveBeenCalledWith(file);
      await view.onUnloadFile(file);
    }
  });
});

describe("Slides text: superscript, subscript, highlight", () => {
  it("toggles superscript / subscript from Format › Text and highlights from the toolbar", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "sup.pptx");
    const v = view as any;
    const shape = v.rendered.find((r: any) => r.textEl && !r.shape.inherited && r.textEl.textContent.trim());
    v.setSelection([shape.shape.el], shape.shape.el);
    const textEl = () => v.rendered.find((r: any) => r.shape.el === shape.shape.el).textEl as HTMLElement;
    expect(menuTitles(view, "Format")).toContain("Text");
    runMenu(view, "Format", "Text", "Superscript");
    expect(textEl().querySelector('[data-baseline="30"]')).not.toBeNull();
    runMenu(view, "Format", "Text", "Subscript");
    expect(textEl().querySelector('[data-baseline="-25"]')).not.toBeNull();
    runMenu(view, "Format", "Text", "Subscript");
    expect(textEl().querySelector('[data-baseline="-25"]')).toBeNull();
    const highlight = view.contentEl.querySelector<HTMLInputElement>('input[aria-label="Highlight color"]')!;
    highlight.value = "#ffff00";
    highlight.dispatchEvent(new Event("change"));
    expect(textEl().querySelector('[data-highlight="#FFFF00"]')).not.toBeNull();
    await view.onUnloadFile(file);
  });
});

describe("Edit › Select all", () => {
  it("selects every object on the slide", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "select-all.pptx");
    const v = view as any;
    runMenu(view, "Edit", "Select all");
    expect(v.items.length).toBe(v.slideItems().length);
    expect(v.items.length).toBeGreaterThan(0);
    await view.onUnloadFile(file);
  });
});

describe("DocxView zoom and panels", () => {
  it("zooms the page from View › Zoom and the toolbar, as Slides and Sheets do; Fit fits the page to the pane", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "zoom.docx");
    const v = view as any;
    expect(menuTitles(view, "View")).toContain("Zoom");
    runMenu(view, "View", "Zoom", "150%");
    expect(v.editor.dom.style.zoom).toBe("1.5");
    const select = view.contentEl.querySelector<HTMLSelectElement>('select[aria-label="Zoom"]')!;
    expect(select.value).toBe("1.5");
    select.value = "0.75";
    select.dispatchEvent(new Event("change"));
    expect(v.editor.dom.style.zoom).toBe("0.75");
    runMenu(view, "View", "Zoom", "Fit");
    expect(v.editor.dom.style.zoom).toBe("1");
    await view.onUnloadFile(file);
  });

  it("puts Paint format just left of Clear formatting in Docs and Slides", async () => {
    for (const [View, sample] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
      const view = mount(View as any, sample);
      const file = await open(view, `order.${sample.split(".")[1]}`);
      const labels = Array.from(view.contentEl.querySelectorAll(".dw-toolbar button")).map((b) => b.getAttribute("aria-label"));
      expect(labels[labels.indexOf("Clear formatting") - 1]).toBe("Paint format (double-click to keep painting)");
      await view.onUnloadFile(file);
    }
  });

  it("names toolbar buttons as Google does in Docs and Slides, with link and comment right after the colours", async () => {
    for (const [View, sample] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
      const view = mount(View as any, sample);
      const file = await open(view, `names.${sample.split(".")[1]}`);
      const labels = Array.from(view.contentEl.querySelectorAll(".dw-toolbar button, .dw-toolbar label.dw-color")).map((b) => b.getAttribute("aria-label"));
      for (const l of ["Insert link (Ctrl+K)", "Add comment (Ctrl+Alt+M)", "Insert image", "Bulleted list (Ctrl+Shift+8)", "Numbered list (Ctrl+Shift+7)", "Decrease indent (Shift+Tab)", "Increase indent (Tab)"]) expect(labels, `${sample}: ${l}`).toContain(l);
      expect(labels.slice(labels.indexOf("Highlight color") + 1, labels.indexOf("Highlight color") + 4)).toEqual(["Insert link (Ctrl+K)", "Add comment (Ctrl+Alt+M)", "Insert image"]);
      if (View === PptxView) for (const l of ["Insert text box", "Insert a shape or line"]) expect(view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`)!.classList.contains("dw-icon"), l).toBe(true);
      await view.onUnloadFile(file);
    }
  });

  it("Docs: Ctrl+Alt+M adds a comment, Ctrl+Shift+8 / 7 make bulleted / numbered lists", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "keys.docx");
    const v = view as any;
    const key = (key: string, keyCode: number, extra: KeyboardEventInit) =>
      v.editor.dom.dispatchEvent(new KeyboardEvent("keydown", { key, keyCode, ctrlKey: true, bubbles: true, cancelable: true, ...extra }));
    const para = () => v.editor.state.selection.$from.parent;
    key("*", 56, { shiftKey: true });
    expect(para().attrs.numId).not.toBeNull();
    const bullet = para().attrs.numId;
    key("&", 55, { shiftKey: true });
    expect(para().attrs.numId).not.toBeNull();
    expect(para().attrs.numId).not.toBe(bullet);
    const comment = vi.spyOn(v, "addComment").mockResolvedValue(undefined);
    key("m", 77, { altKey: true });
    expect(comment).toHaveBeenCalled();
    await view.onUnloadFile(file);
  });

  it("Slides: Ctrl+Alt+M opens the slide's comments", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "keys.pptx");
    const scope = view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => boolean }[] };
    scope.keys.find((k) => k.key === "m" && k.mods.join() === "Mod,Alt")!.fn();
    expect(document.querySelector(".test-modal")?.textContent).toContain("Comments on slide 1");
    await view.onUnloadFile(file);
  });

  it("Slides zoom lists Fit then 50%–200% like Docs; 100% draws the slide at its real size", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "zoom.pptx");
    const v = view as any;
    menuTitles(view, "View");
    const zoom = Menu.last!.items.find((i) => i.title === "Zoom")!.submenu!.items.map((i) => i.title);
    expect(zoom).toEqual(["Fit", "50%", "75%", "90%", "100%", "125%", "150%", "200%"]);
    runMenu(view, "View", "Zoom", "100%");
    expect(v.scale * v.doc.slideSize.cx).toBeCloseTo(v.doc.slideSize.cx / 9525, 3);
    runMenu(view, "View", "Zoom", "200%");
    expect(v.scale * v.doc.slideSize.cx).toBeCloseTo((2 * v.doc.slideSize.cx) / 9525, 3);
    await view.onUnloadFile(file);
  });

  it("Insert › Special characters puts the picked character at the cursor (Docs and Slides)", async () => {
    const docs = mount(DocxView, "sample.docx");
    const f1 = await open(docs, "chars.docx");
    const d = docs as any;
    const { TextSelection } = await import("prosemirror-state");
    d.editor.dispatch(d.editor.state.tr.setSelection(TextSelection.create(d.editor.state.doc, 1)));
    runMenu(docs, "Insert", "Special characters");
    document.querySelector<HTMLButtonElement>('.test-modal .dw-char[aria-label="Section sign"]')!.click();
    expect(d.editor.state.doc.firstChild.textContent.startsWith("§")).toBe(true);
    await docs.onUnloadFile(f1);
    document.querySelectorAll(".test-modal").forEach((m) => m.remove());

    const slides = mount(PptxView, "sample.pptx");
    const f2 = await open(slides, "chars.pptx");
    const s = slides as any;
    const rs = s.rendered.find((r: any) => r.textEl && r.textEl.textContent);
    s.activeText = rs;
    const range = document.createRange();
    range.setStart(rs.textEl.querySelector("span")!.firstChild, 0);
    range.collapse(true);
    s.savedRange = range;
    runMenu(slides, "Insert", "Special characters");
    document.querySelector<HTMLButtonElement>('.test-modal .dw-char[aria-label="Section sign"]')!.click();
    expect(s.rendered.find((r: any) => r.shape.el === rs.shape.el).textEl.textContent.startsWith("§")).toBe(true);
    await slides.onUnloadFile(f2);
  });

  it("Docs: Insert › Chart adds a native chart drawn on the page; double-click edits it; undo brings the old one back", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "chart.docx");
    const v = view as any;
    const set = (modal: HTMLElement, label: string, value: string) => {
      const el = modal.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
      el.value = value;
      el.dispatchEvent(new Event(el.tagName === "SELECT" ? "change" : "input"));
    };
    runMenu(view, "Insert", "Chart", "Column");
    let modal = document.querySelector<HTMLElement>(".test-modal")!;
    set(modal, "Chart title", "Sales");
    set(modal, "Series 1, row 1", "42");
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Insert chart")!.click();
    await vi.waitFor(() => expect(v.editor.dom.querySelector(".docx-chart svg")).not.toBeNull());
    expect(v.editor.dom.querySelector(".docx-chart .chart-title")?.textContent).toBe("Sales");
    modal.remove();
    v.editor.dom.querySelector(".docx-chart").dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    modal = document.querySelector<HTMLElement>(".test-modal")!;
    expect(modal.querySelector<HTMLInputElement>('[aria-label="Series 1, row 1"]')!.value).toBe("42");
    set(modal, "Chart title", "Revenue");
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Apply")!.click();
    await vi.waitFor(() => expect(v.editor.dom.querySelector(".docx-chart .chart-title")?.textContent).toBe("Revenue"));
    const { undo } = await import("prosemirror-history");
    undo(v.editor.state, v.editor.dispatch);
    expect(v.editor.dom.querySelector(".docx-chart .chart-title")?.textContent).toBe("Sales");
    await view.onUnloadFile(file);
  });

  it("menu bar items show icons, as Google's menus do (Docs and Slides)", async () => {
    for (const [View, sample] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
      const view = mount(View as any, sample);
      const file = await open(view, `icons.${sample.split(".")[1]}`);
      for (const [menu, title, icon] of [["Edit", "Cut", "scissors"], ["Edit", "Paste", "clipboard-paste"], ["Insert", "Image", "image"], ["Insert", "Comment", "message-square-plus"], ["Format", "Clear formatting", "remove-formatting"], ["Tools", "Find and replace", "replace"]]) {
        menuTitles(view, menu);
        expect(Menu.last!.items.find((i) => i.title === title)?.icon, `${sample} ${menu} › ${title}`).toBe(icon);
      }
      await view.onUnloadFile(file);
    }
  });

  it("Insert › Chart lists Google's chart types (Docs and Slides); the editor opens on the one picked", async () => {
    for (const [View, sample] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
      const view = mount(View as any, sample);
      const file = await open(view, `types.${sample.split(".")[1]}`);
      menuTitles(view, "Insert");
      expect(Menu.last!.items.find((i) => i.title === "Chart")!.submenu!.items.map((i) => i.title)).toEqual(["Bar", "Column", "Line", "Pie"]);
      for (const [type, value] of [["Pie", "pie:none"], ["Bar", "bar:none"], ["Line", "line:none"]]) {
        runMenu(view, "Insert", "Chart", type);
        const modal = Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
        expect(modal.querySelector<HTMLSelectElement>('[aria-label="Chart type"]')!.value, `${sample} ${type}`).toBe(value);
        modal.remove();
      }
      await view.onUnloadFile(file);
    }
  });

  it("Help › Keyboard shortcuts (Ctrl+/) lists the app's shortcuts, searchable (Docs and Slides)", async () => {
    for (const [View, sample, own] of [[DocxView, "sample.docx", "Page break"], [PptxView, "sample.pptx", "Duplicate"]] as const) {
      const view = mount(View as any, sample);
      const file = await open(view, `keys.${sample.split(".")[1]}`);
      const scope = view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => boolean }[] };
      expect(scope.keys.some((k) => k.key === "/" && k.mods.join() === "Mod"), sample).toBe(true);
      runMenu(view, "Help", "Keyboard shortcuts");
      const modal = Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
      const rows = () => Array.from(modal.querySelectorAll(".dw-shortcut")).map((r) => r.textContent);
      expect(rows().some((t) => t?.includes(own)), `${sample} ${own}`).toBe(true);
      expect(rows().some((t) => t?.includes("Undo"))).toBe(true);
      const search = modal.querySelector<HTMLInputElement>('input[type="search"]')!;
      search.value = "bold";
      search.dispatchEvent(new Event("input"));
      expect(rows().length).toBe(1);
      expect(rows()[0]).toContain("Ctrl+B");
      modal.remove();
      await view.onUnloadFile(file);
    }
  });

  it("Slides: Ctrl+K edits the link, as its button says", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "link-key.pptx");
    const link = vi.spyOn(view as any, "editLink").mockResolvedValue(undefined);
    const scope = view.scope as unknown as { keys: { mods: string[]; key: string; fn: () => boolean }[] };
    scope.keys.find((k) => k.key === "k" && k.mods.join() === "Mod")!.fn();
    expect(link).toHaveBeenCalled();
    await view.onUnloadFile(file);
  });

  it("closes the comments sidebar with its × button", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "review-x.docx");
    const v = view as any;
    v.setReview(true);
    expect(v.reviewEl.style.display).not.toBe("none");
    v.reviewEl.querySelector('[aria-label="Close comments"]').click();
    expect(v.reviewEl.style.display).toBe("none");
    expect(v.reviewOpen).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("DocxView outline", () => {
  it("nests sub-headings under their section, counting from the top level the document uses", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "outline-depth.docx");
    const v = view as any;
    if (!v.outlineOpen) runMenu(view, "View", "Outline");
    const ed = v.editor, s = ed.state.schema;
    const paras = [s.nodes.paragraph.create({ style: "Heading2" }, s.text("Part")), s.nodes.paragraph.create({ style: "Heading3" }, s.text("Sub")), s.nodes.paragraph.create({ style: "Heading2" }, s.text("Next"))];
    ed.dispatch(ed.state.tr.replaceWith(0, ed.state.doc.content.size, paras));
    v.renderOutline();
    const rows = Array.from(v.outlineEl.querySelectorAll(".docx-outline-entry") as NodeListOf<HTMLElement>).map((b) => [b.textContent, b.dataset.depth]);
    expect(rows).toEqual([["Part", "0"], ["Sub", "1"], ["Next", "0"]]);
    await view.onUnloadFile(file);
  });

  it("closes the document outline with its × button", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "outline-x.docx");
    const v = view as any;
    if (!v.outlineOpen) runMenu(view, "View", "Outline");
    const outline = v.outlineEl as HTMLElement;
    expect(outline.style.display).not.toBe("none");
    outline.querySelector<HTMLButtonElement>('[aria-label="Close outline"]')!.click();
    expect(outline.style.display).toBe("none");
    expect(v.outlineOpen).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("DocxView word count", () => {
  it("Tools › Word count shows pages, words and characters like Google Docs, and turns the count while typing on and off", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "count.docx");
    const v = view as any;
    runMenu(view, "Tools", "Word count");
    const modal = document.body.querySelector<HTMLElement>(".test-modal")!;
    const rows = Array.from(modal.querySelectorAll(".dw-count-row")).map((r) => r.firstElementChild!.textContent);
    expect(rows).toEqual(["Pages", "Words", "Characters", "Characters excluding spaces"]);
    const box = modal.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    expect(box.checked).toBe(true);
    box.checked = false;
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "OK")!.click();
    expect(v.wordsEl.hidden).toBe(true);
    expect(document.body.querySelector(".test-modal")).toBeNull();
    // A selection shows "n of total".
    const { TextSelection } = await import("prosemirror-state");
    v.editor.dispatch(v.editor.state.tr.setSelection(TextSelection.create(v.editor.state.doc, 1, 4)));
    runMenu(view, "Tools", "Word count");
    expect(document.body.querySelector(".test-modal .dw-count-row:nth-child(3)")!.lastElementChild!.textContent).toMatch(/^3 of /);
    await view.onUnloadFile(file);
  });
});

describe("DocxView accepting suggestions", () => {
  it("undoes and redoes Accept and Reject with Ctrl+Z / Ctrl+Y, as Google Docs does", async () => {
    const { undo, redo } = await import("prosemirror-history");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "accept-undo.docx");
    const v = view as any;
    v.settings.author = "Ada";
    v.setReview(true);
    v.suggesting = true;
    v.editor.dispatch(v.editor.state.tr.insertText("Proposed ", 1));
    const review = () => view.contentEl.querySelector(".docx-review")!.textContent!;
    for (const button of ["Accept all", "Reject all"]) {
      v.renderReview();
      expect(review()).toContain("Inserted by Ada");
      Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".docx-review button")).find((b) => b.textContent === button)!.click();
      expect(review()).not.toContain("Inserted by Ada");
      expect(undo(v.editor.state, v.editor.dispatch)).toBe(true);
      v.renderReview();
      expect(review()).toContain("Inserted by Ada");
      expect(v.editor.state.doc.textContent).toContain("Proposed ");
    }
    expect(redo(v.editor.state, v.editor.dispatch)).toBe(true);
    v.renderReview();
    expect(review()).not.toContain("Inserted by Ada");
    expect(v.editor.state.doc.textContent).not.toContain("Proposed ");
    await view.onUnloadFile(file);
  });
});

describe("DocxView comment author", () => {
  it("records editor suggestions and persists them, and displays navigable headings and a ruler", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
    const saved = vi.spyOn(app.vault, "modifyBinary");
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, pageSize: "document" as const, author: "Ada" });
    document.body.appendChild(view.containerEl);
    const file = await open(view, "suggestions.docx");
    const v = view as any;
    const mode = view.contentEl.querySelector<HTMLSelectElement>('select[aria-label="Editing mode"]')!;
    expect(mode.value).toBe("editing");
    expect(saved).not.toHaveBeenCalled();
    mode.value = "suggesting";
    mode.dispatchEvent(new Event("change"));
    await Promise.resolve();
    await Promise.resolve();
    v.editor.dispatch(v.editor.state.tr.insertText("Proposed ", 1));
    v.renderReview();
    expect(view.contentEl.querySelector(".docx-review")!.textContent).toContain("Inserted by Ada");
    expect(view.contentEl.querySelector('[aria-label="Document ruler"]')).not.toBeNull();
    const host = view.contentEl.querySelector(".docx-editor-host")!;
    expect(host.querySelector(".docx-ruler")).toBeNull();
    expect(host.previousElementSibling!.className).toBe("docx-ruler-viewport");
    runMenu(view, "View", "Outline");
    expect(view.contentEl.querySelectorAll(".docx-outline-entry").length).toBeGreaterThan(0);
    const heading = v.editor.dom.querySelector("p");
    vi.spyOn(v.editor, "nodeDOM").mockReturnValue(heading);
    vi.spyOn(heading, "getBoundingClientRect").mockReturnValue({ top: 1000 } as DOMRect);
    vi.spyOn(host, "getBoundingClientRect").mockReturnValue({ top: 100 } as DOMRect);
    host.scrollTop = 50;
    view.contentEl.querySelector<HTMLButtonElement>(".docx-outline-entry")!.click();
    expect(host.scrollTop).toBe(934);
    await view.requestSave();
    const back = await DocxDocument.load(saved.mock.calls.at(-1)![1]);
    expect(back.body.getElementsByTagNameNS(W, "ins")[0].textContent).toBe("Proposed ");
    await view.onUnloadFile(file);
  });
  it("signs comments with the author name from settings", async () => {
    const settings: DocWorksSettings = { ...DEFAULT_SETTINGS, pageSize: "document" as const, author: "Ada Lovelace" };
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, settings);
    document.body.appendChild(view.containerEl);
    await open(view, "a.docx");
    expect(await (view as any).author()).toBe("Ada Lovelace");
  });
});

describe("DocxView header pictures", () => {
  it("shows a picture that lives in the header part (a logo), from the header's own relationships", async () => {
    const { blankDocx } = await import("../src/infrastructure/blank");
    const { writeBody } = await import("../src/domain/docx/docx-writer");
    const { docxSchema: schema } = await import("../src/domain/docx/schema");
    const docx = await DocxDocument.load(await blankDocx());
    const root = docx.ensureHeaderFooter("header");
    const path = docx.headerFooterPath("header")!;
    const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
    const rId = docx.addImage(PNG, "png", path);
    expect(docx.rels(docx.mainPath).some((r) => r.id === rId && r.type === "image")).toBe(false);
    const logo = schema.nodes.image.create({ rId, width: 40, height: 20, alt: "Logo" });
    writeBody(root, schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, logo)), new Map(), docx);
    const bytes = await docx.save();

    const app = new App();
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    await open(view, "logo.docx");
    const band = (view as any).renderHF("header", 1, 1) as HTMLElement;
    expect(band.querySelector("img")?.getAttribute("src")).toMatch(/^data:image\/png;base64,/);
  });
});

describe("current font and size in the toolbar", () => {
  it("Word: shows the font and size at the cursor, from the text's own formatting or its style", async () => {
    const view = mount(DocxView, "sample.docx");
    await open(view, "a.docx");
    const v = view as any;
    const { docxSchema: schema } = await import("../src/domain/docx/schema");
    const { TextSelection } = await import("prosemirror-state");
    const editor = v.editor;
    const marked = schema.text("Hi", [schema.marks.font.create({ font: "Georgia" }), schema.marks.size.create({ pt: 18 })]);
    editor.dispatch(editor.state.tr.insert(1, marked));
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2)));
    v.updateFormatControls();
    const [font, size] = ["Font", "Font size"].map((t) => view.contentEl.querySelector<HTMLSelectElement>(`select[aria-label="${t}"]`)!);
    expect(font.value).toBe("Georgia");
    expect(size.value).toBe("18");
    // The style select follows the paragraph (the sample's first paragraph is a styled one).
    const style = view.contentEl.querySelector<HTMLSelectElement>('select[aria-label="Paragraph style"]')!;
    expect(style.options[style.selectedIndex].text).not.toBe("Style");
  });

  it("Slides: shows the font and size of the text at the cursor", async () => {
    const view = mount(PptxView, "sample.pptx");
    await open(view, "a.pptx");
    const v = view as any;
    const title = v.rendered.find((r: any) => r.shape.name === "Title 1");
    const textNode = title.textEl.querySelector("span")?.firstChild ?? title.textEl.querySelector("p").firstChild;
    getSelection()!.collapse(textNode, 1);
    v.onSelectionChange();
    const [font, size] = ["Font", "Font size"].map((t) => view.contentEl.querySelector<HTMLSelectElement>(`select[aria-label="${t}"]`)!);
    expect(font.value).toBe("Calibri");
    expect(size.value).toBe("44");
  });
});

describe("slide authoring additions",()=>{
 it("edits the master canvas and theme without saving on open",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  const saved=vi.spyOn(view.app.vault,"modifyBinary");
  runMenu(view, "Slide", "Theme…");expect(saved).not.toHaveBeenCalled();
  const select=view.contentEl.querySelector<HTMLSelectElement>('[aria-label="Edit master or layout"]')!;
  select.value=select.options[1].value;select.dispatchEvent(new Event("change"));
  const before=v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length;
  runMenu(view, "Insert", "Text box");
  expect(v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length).toBe(before+1);
  v.undo();expect(v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length).toBe(before);
  await view.onUnloadFile(file);
 });
 it("inserts a formatted chart: stacked bars with a title and no legend, then changes it to a pie",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  runMenu(view, "Insert", "Chart", "Column");
  let modal=document.querySelector<HTMLElement>(".test-modal")!;
  const set=(label:string,value:string|boolean)=>{const el=modal.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;if(typeof value==="boolean")el.checked=value;else el.value=value;el.dispatchEvent(new Event(el.tagName==="SELECT"||el.type==="checkbox"?"change":"input"));};
  set("Chart type","bar:stacked");set("Chart title","Sales");set("Legend","none");set("Data labels",true);
  Array.from(modal.querySelectorAll("button")).find(b=>b.textContent==="Insert chart")!.click();
  await vi.waitFor(()=>expect(v.model.shapes.find((s:any)=>s.chart)).toBeTruthy());
  expect(v.model.shapes.find((s:any)=>s.chart).chart.format).toMatchObject({kind:"bar",stacking:"stacked",title:"Sales",legend:"none",labels:true,category:{reversed:true}});
  expect(view.contentEl.querySelector(".chart-title")?.textContent).toBe("Sales");
  modal.remove();
  v.editChart(v.model.shapes.find((s:any)=>s.chart).chart.path);
  modal=document.querySelector<HTMLElement>(".test-modal")!;
  expect(modal.querySelector<HTMLSelectElement>('[aria-label="Chart type"]')!.value).toBe("bar:stacked");
  set("Chart type","pie:none");
  Array.from(modal.querySelectorAll("button")).find(b=>b.textContent==="Apply")!.click();
  await vi.waitFor(()=>expect(v.model.shapes.find((s:any)=>s.chart).chart.format.kind).toBe("pie"));
  await view.onUnloadFile(file);
 });
 it("keeps the caret in the text box after Ctrl+Z / Ctrl+Y there, so typing carries on",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  const shape=v.rendered.find((r:any)=>r.textEl&&!r.shape.inherited);
  const ids=shape.el.dataset.spids;
  v.mutate((d:any)=>d.addSlide(v.index));v.index=0;v.renderAll();
  const textEl=v.rendered.find((r:any)=>r.el.dataset.spids===ids).textEl as HTMLElement;
  // jsdom can't focus contenteditable elements: record what gets focused instead.
  const focused:HTMLElement[]=[];
  const spy=vi.spyOn(HTMLElement.prototype,"focus").mockImplementation(function(this:HTMLElement){focused.push(this);});
  textEl.dispatchEvent(new KeyboardEvent("keydown",{key:"z",ctrlKey:true,bubbles:true}));
  const again=v.rendered.find((r:any)=>r.el.dataset.spids===ids).textEl as HTMLElement;
  expect(again).not.toBe(textEl);
  expect(focused.at(-1)).toBe(again);
  again.dispatchEvent(new KeyboardEvent("keydown",{key:"y",ctrlKey:true,bubbles:true}));
  expect(focused.at(-1)).toBe(v.rendered.find((r:any)=>r.el.dataset.spids===ids).textEl);
  spy.mockRestore();
  await view.onUnloadFile(file);
 });
 it("undoes with Ctrl+Z after a dialog closes, when focus is on the page rather than the slide",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  runMenu(view, "Insert", "Chart", "Column");
  const modal=document.querySelector<HTMLElement>(".test-modal")!;
  Array.from(modal.querySelectorAll("button")).find(b=>b.textContent==="Insert chart")!.click();
  await vi.waitFor(()=>expect(v.model.shapes.some((s:any)=>s.chart)).toBe(true));
  (document.activeElement as HTMLElement|null)?.blur();
  const scope=view.scope as unknown as {keys:{mods:string[];key:string;fn:()=>boolean}[]};
  const undo=scope.keys.find(k=>k.key==="z"&&k.mods.length===1)!;
  expect(undo.fn()).toBe(false);
  expect(v.model.shapes.some((s:any)=>s.chart)).toBe(false);
  scope.keys.find(k=>k.key==="y")!.fn();
  expect(v.model.shapes.some((s:any)=>s.chart)).toBe(true);
  // Typing in a field keeps the field's own undo.
  const input=view.contentEl.appendChild(createEl("input"));input.focus();
  expect(undo.fn()).toBe(true);
  await view.onUnloadFile(file);
 });
 it("inserts an editable chart from its data grid and supports undo",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  runMenu(view, "Insert", "Chart", "Column");
  const modal=document.querySelector<HTMLElement>(".test-modal")!;
  const field=modal.querySelector<HTMLInputElement>('input[aria-label="Series 1, row 1"]')!;field.value="42";field.dispatchEvent(new Event("input"));
  Array.from(modal.querySelectorAll("button")).find(b=>b.textContent==="Insert chart")!.click();
  await vi.waitFor(()=>expect(v.model.shapes.find((s:any)=>s.chart)?.chart.data.series[0].values[0]).toBe(42));
  expect(view.contentEl.querySelector('svg[aria-label="Chart"]')).not.toBeNull();
  const shape=v.model.shapes.find((s:any)=>s.chart);v.setSelection([shape.el]);
  expect(view.contentEl.textContent).toContain("Edit chart");
  v.undo();expect(v.model.shapes.some((s:any)=>s.chart)).toBe(false);
  await view.onUnloadFile(file);
 });
});


describe("editor accessibility", () => {
  it("navigates slide thumbnails with arrows/Home/End and exposes selection without saving", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "accessibility.pptx");
    const saved = vi.spyOn(view.app.vault, "modifyBinary");
    const list = view.contentEl.querySelector<HTMLElement>('[role="listbox"]')!;
    expect(document.getElementById(list!.getAttribute("aria-labelledby")!)?.textContent).toBe("Slides");
    list.focus();
    list.dispatchEvent(new KeyboardEvent("keydown", {key: "End", bubbles: true}));
    const options = list.querySelectorAll('[role="option"]');
    expect(options[options.length-1].getAttribute("aria-selected")).toBe("true");
    expect(list.getAttribute("aria-activedescendant")).toBe(options[options.length-1].id);
    list.dispatchEvent(new KeyboardEvent("keydown", {key: "Home", bubbles: true}));
    expect(options[0].getAttribute("aria-selected")).toBe("true");
    list.dispatchEvent(new KeyboardEvent("keydown", {key: "ArrowDown", bubbles: true}));
    expect(options[1].getAttribute("aria-selected")).toBe("true");
    expect(saved).not.toHaveBeenCalled();
    await view.onUnloadFile(file);
  });
  it("names the document editor and theme/color controls", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "accessibility.docx");
    // Named by a hidden label, not aria-label (Obsidian turns aria-label into a hover tooltip over the whole page).
    const pm = view.contentEl.querySelector('.ProseMirror')!;
    expect(pm.getAttribute("aria-label")).toBeNull();
    expect(document.getElementById(pm.getAttribute("aria-labelledby")!)?.textContent).toBe("Document content");
    for (const input of Array.from(view.contentEl.querySelectorAll('input[type="color"]'))) expect(input.getAttribute("aria-label")).toBeTruthy();
    expect(menuTitles(view, "View")).toContain("Light");
    expect(Menu.last!.items.find((i) => i.title === "Light")!.checked).toBe(true);
    await view.onUnloadFile(file);
  });
});

describe("PptxView links", () => {
  it("links the selected shape to a slide from the link dialog, and follows it with Ctrl+click", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "links.pptx");
    const v = view as any;
    const shape = v.rendered.find((r: any) => !r.shape.inherited && r.shape.kind === "sp").shape.el;
    v.setSelection([shape]);
    v.activeText = null;
    const done = v.editLink();
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    expect(modal.textContent).toContain("Slides in this presentation");
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Last slide")!.click();
    await done;
    const link = v.rendered.find((r: any) => r.shape.el.isEqualNode?.(shape) || r.shape.link)?.shape.link;
    expect(link).toBe("#jump:lastslide");
    const el = v.rendered.find((r: any) => r.shape.link).el as HTMLElement;
    el.dispatchEvent(new MouseEvent("click", { bubbles: true, ctrlKey: true }));
    expect(v.index).toBe(v.doc.slides().length - 1);
    await view.onUnloadFile(file);
  });
});

describe("PptxView menus", () => {
  it("has Google's menus with every command", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "menus.pptx");
    expect(menuNames(view)).toEqual(["File", "Edit", "View", "Insert", "Format", "Slide", "Arrange", "Tools", "Help"]);
    expect(menuTitles(view, "Edit")).toEqual(["Undo", "Redo", "Cut", "Copy", "Copy as Markdown outline", "Copy link to slide", "Paste", "Paste without formatting", "Select all", "Delete", "Find and replace"]);
    expect(menuTitles(view, "Tools")).toEqual(["Paint format", "Find and replace"]);
    expect(view.contentEl.querySelector('.dw-toolbar [aria-label^="Insert link"]')).not.toBeNull();
    expect(menuTitles(view, "View")).toEqual(["Present", "Present from start", "Grid view", "Format panel", "Comments", "Zoom", "Light", "Dark"]);
    expect(menuTitles(view, "Insert")).toEqual(["Text box", "Image", "Audio or video", "Shape", "Table", "Chart", "Link", "Field", "Special characters", "Comment", "New slide"]);
    expect(menuTitles(view, "Format")).toEqual(["Text", "Align", "Line spacing", "Character spacing", "Bulleted list", "Numbered list", "Clear formatting"]);
    expect(menuTitles(view, "Arrange")).toEqual(["Order", "Align", "Distribute", "Group", "Ungroup"]);
    const slide = menuTitles(view, "Slide");
    expect(slide).toContain("New slide");
    expect(slide.slice(-2)).toEqual(["Theme…", "Motion…"]);
    await view.onUnloadFile(file);
  });

  it("shows shape tools for a shape, and sets its fill from the row", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "menus.pptx");
    const v = view as any;
    const group = (l: string) => view.contentEl.querySelector<HTMLElement>(`.dw-toolbar [aria-label="${l}"]`)!.closest<HTMLElement>(".dw-group")!;
    expect(group("Shape fill").hidden).toBe(true);
    expect(group("Replace image").hidden).toBe(true);
    const shape = v.rendered.find((r: any) => r.shape.kind === "sp" && !r.shape.inherited && r.shape.el.localName === "sp");
    v.setSelection([shape.shape.el]);
    expect(group("Shape fill").hidden).toBe(false);
    expect(group("Replace image").hidden).toBe(true);
    const input = view.contentEl.querySelector<HTMLInputElement>('.dw-toolbar [aria-label="Shape fill"] input')!;
    input.value = "#ff0000";
    input.dispatchEvent(new Event("change"));
    const el = shape.shape.el;
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const spPr = Array.from(el.children as HTMLCollection).find((c: any) => c.localName === "spPr") as Element;
    expect(spPr.getElementsByTagNameNS(A, "srgbClr")[0].getAttribute("val")).toBe("FF0000");
    v.setSelection([]);
    expect(group("Shape fill").hidden).toBe(true);
    await view.onUnloadFile(file);
  });

  it("toggles grid view and the theme from View", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "menus.pptx");
    runMenu(view, "View", "Grid view");
    expect(view.contentEl.classList.contains("pptx-grid-view")).toBe(true);
    runMenu(view, "View", "Dark");
    expect(view.contentEl.classList.contains("pptx-theme-dark")).toBe(true);
    expect(view.contentEl.querySelector(".dw-segmented")).toBeNull();
    await view.onUnloadFile(file);
  });

  it("starts with the format panel closed; View › Format panel opens it", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "menus.pptx");
    const props = view.contentEl.querySelector<HTMLElement>(".pptx-props")!;
    expect(props.style.display).toBe("none");
    runMenu(view, "View", "Format panel");
    expect(props.style.display).not.toBe("none");
    await view.onUnloadFile(file);
  });

  it("opens the format panel as a drawer in a narrow pane", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "menus.pptx");
    const v = view as any;
    Object.defineProperty(view.contentEl, "clientWidth", { configurable: true, value: 500 });
    v.narrowWatch.check();
    const props = view.contentEl.querySelector<HTMLElement>(".pptx-props")!;
    const shown = v.showProps;
    runMenu(view, "View", "Format panel");
    expect(props.classList.contains("is-drawer")).toBe(true);
    expect(v.showProps).toBe(shown);
    props.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(props.classList.contains("is-drawer")).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("narrow panes (review fixes)", () => {
  const narrow = (view: { contentEl: HTMLElement }) => {
    Object.defineProperty(view.contentEl, "clientWidth", { configurable: true, value: 500 });
    (view as any).narrowWatch.check();
  };

  it("keeps the slide format drawer open when the selection is redrawn, even if the docked panel is hidden", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "narrow.pptx");
    const v = view as any;
    v.showProps = false;
    narrow(view);
    runMenu(view, "View", "Format panel");
    const shape = v.rendered.find((r: any) => r.shape.kind === "sp" && !r.shape.inherited);
    v.setSelection([shape.shape.el]);
    const props = view.contentEl.querySelector<HTMLElement>(".pptx-props")!;
    expect(props.classList.contains("is-drawer")).toBe(true);
    expect(props.style.display).not.toBe("none");
    await view.onUnloadFile(file);
  });

  it("opens Theme… and Motion… in the drawer", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "narrow.pptx");
    narrow(view);
    runMenu(view, "Slide", "Theme…");
    const props = view.contentEl.querySelector<HTMLElement>(".pptx-props")!;
    expect(props.classList.contains("is-drawer")).toBe(true);
    props.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    runMenu(view, "Slide", "Motion…");
    expect(props.classList.contains("is-drawer")).toBe(true);
    await view.onUnloadFile(file);
  });

  it("shows the review panel as a drawer when a comment is added, without changing the remembered docked state", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "narrow.docx");
    const v = view as any;
    const before = view.app.loadLocalStorage("doc-works-review-panel");
    const wasOpen = v.reviewOpen;
    narrow(view);
    v.setReview(true);
    const review = view.contentEl.querySelector<HTMLElement>(".docx-review")!;
    expect(review.classList.contains("is-drawer")).toBe(true);
    expect(review.childElementCount).toBeGreaterThan(0);
    expect(v.reviewOpen).toBe(wasOpen);
    expect(view.app.loadLocalStorage("doc-works-review-panel")).toBe(before);
    await view.onUnloadFile(file);
  });
});

describe("PptxView grid view", () => {
  it("toggles a grid of slides from the toolbar and Ctrl+Alt+1; double-clicking a slide opens it", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "grid.pptx");
    const v = view as any;
    runMenu(view, "View", "Grid view");
    expect(view.contentEl.classList.contains("pptx-grid-view")).toBe(true);
    const key = v.scope.keys.find((k: any) => k.key === "1" && k.mods.join() === "Mod,Alt");
    key.fn();
    expect(view.contentEl.classList.contains("pptx-grid-view")).toBe(false);
    key.fn();
    const thumbs = view.contentEl.querySelectorAll<HTMLElement>(".pptx-thumb");
    thumbs[1].dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(v.index).toBe(1);
    expect(view.contentEl.classList.contains("pptx-grid-view")).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("PptxView paint format", () => {
  it("copies the selected shape's format and paints it onto the next shape clicked", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "paint.pptx");
    const v = view as any;
    const own = () => v.rendered.filter((r: any) => !r.shape.inherited && r.shape.kind === "sp" && r.shape.el.localName === "sp");
    const [a, b] = own();
    v.mutate(() => {
      writeFill(a.shape.el.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spPr")[0], { type: "solid", color: "#123456" });
    }, "slide");
    const src = own()[0];
    v.setSelection([src.shape.el]);
    const brush = view.contentEl.querySelector<HTMLElement>('[aria-label="Paint format (double-click to keep painting)"]')!;
    brush.click();
    expect(brush.classList.contains("is-active")).toBe(true);
    const target = own().find((r: any) => r.shape.el !== src.shape.el && r.shape.el.isEqualNode(b.shape.el)) ?? own()[1];
    target.el.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
    expect(own().find((r: any) => r.shape.fill?.color === "#123456" && r !== own()[0])).toBeTruthy();
    expect(brush.classList.contains("is-active")).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("DocxView watermark", () => {
  it("adds a text watermark from the dialog, saves it in the header and draws it on the pages", async () => {
    const { readWatermark } = await import("../src/domain/docx/watermark");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "wm.docx");
    const v = view as any;
    v.watermarkDialog();
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    const text = modal.querySelector<HTMLInputElement>('input[aria-label="Watermark text"]')!;
    text.value = "DRAFT";
    text.dispatchEvent(new Event("input"));
    Array.from(modal.querySelectorAll("button")).find((b) => b.textContent === "Apply")!.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(readWatermark(v.docx)).toMatchObject({ kind: "text", text: "DRAFT", diagonal: true });
    const head = v.renderHF("header", 1, 1, 0) as HTMLElement;
    expect(head.querySelector(".docx-watermark")!.textContent).toBe("DRAFT");
    v.watermarkDialog();
    Array.from(document.querySelectorAll<HTMLButtonElement>(".test-modal button")).find((b) => b.textContent === "Remove watermark")!.click();
    expect(readWatermark(v.docx)).toBeNull();
    await view.onUnloadFile(file);
  });
});

describe("DocxView update style to match", () => {
  it("moves the paragraph's own formatting into its style and clears it from the paragraph", async () => {
    const { TextSelection } = await import("prosemirror-state");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "style.docx");
    const v = view as any;
    const ed = v.editor;
    let at = -1;
    ed.state.doc.descendants((n: any, pos: number) => {
      if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 3) at = pos;
      return at < 0;
    });
    const para = ed.state.doc.nodeAt(at);
    const style = para.attrs.style ?? "Normal";
    const tr = ed.state.tr.setNodeMarkup(at, undefined, { ...para.attrs, align: "center" }).addMark(at + 1, at + 1 + para.content.size, ed.state.schema.marks.italic.create());
    ed.dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 2)));
    const select = view.contentEl.querySelector<HTMLSelectElement>('select[aria-label="Paragraph style"]')!;
    expect(Array.from(select.options).map((o) => o.text)).toContainEqual(expect.stringMatching(/^Update '.+' to match$/));
    select.value = "update";
    select.dispatchEvent(new Event("change"));
    const after = ed.state.doc.nodeAt(at);
    expect(after.attrs.align).toBeNull();
    let italic = false;
    after.descendants((n: any) => void (italic ||= n.marks.some((m: any) => m.type.name === "italic")));
    expect(italic).toBe(false);
    const st = Array.from(v.docx.stylesXml().getElementsByTagNameNS(W, "style") as HTMLCollectionOf<Element>).find((s) => s.getAttributeNS(W, "styleId") === style)!;
    expect(st.getElementsByTagNameNS(W, "jc")[0].getAttributeNS(W, "val")).toBe("center");
    expect(st.getElementsByTagNameNS(W, "i")).toHaveLength(1);
    await view.onUnloadFile(file);
  });
});

describe("DocxView paint format", () => {
  it("copies with the brush and paints the next selection made with the mouse", async () => {
    const { TextSelection } = await import("prosemirror-state");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "paint.docx");
    const v = view as any;
    const ed = v.editor;
    const texts: number[] = [];
    ed.state.doc.descendants((n: any, pos: number) => {
      if (n.type.name === "paragraph" && n.textContent.length > 4 && texts.length < 2) texts.push(pos);
      return texts.length < 2;
    });
    const [a, b] = texts;
    let tr = ed.state.tr.addMark(a + 1, a + 4, ed.state.schema.marks.italic.create());
    ed.dispatch(tr.setSelection(TextSelection.create(tr.doc, a + 2)));
    const brush = view.contentEl.querySelector<HTMLElement>('[aria-label="Paint format (double-click to keep painting)"]')!;
    brush.click();
    expect(brush.classList.contains("is-active")).toBe(true);
    tr = ed.state.tr.setSelection(TextSelection.create(ed.state.doc, b + 1, b + 4));
    ed.dispatch(tr);
    ed.dom.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    await new Promise((r) => setTimeout(r, 0));
    expect(ed.state.doc.resolve(b + 2).marks().some((m: any) => m.type.name === "italic")).toBe(true);
    expect(brush.classList.contains("is-active")).toBe(false);
    await view.onUnloadFile(file);
  });
});

describe("save status in documents and slides", () => {
  for (const [View, name] of [[DocxView, "sample.docx"], [PptxView, "sample.pptx"]] as const) {
    it(`${View.name}: shows Couldn't save with Retry, then Saved`, async () => {
      const view = mount(View as any, name) as any;
      const file = await open(view, `status-${name}`);
      let fail = true;
      view.app.vault.modifyBinary = async () => {
        if (fail) throw new Error("read-only file");
      };
      view.autosave.dirty = true;
      await view.requestSave();
      const status = view.contentEl.querySelector(".dw-status-text")!.parentElement as HTMLElement;
      expect(status.textContent).toContain("Couldn't save");
      fail = false;
      status.querySelector("button")!.click();
      await view.autosave.idle();
      expect(status.textContent).toBe("Saved");
      await view.onUnloadFile(file);
    });
  }
});

describe("keyboard and focus", () => {
  const tabStops = (root: HTMLElement) => Array.from(root.querySelectorAll<HTMLElement>(".dw-toolbar button, .dw-toolbar select, .dw-toolbar input")).filter((el) => el.tabIndex === 0);
  const f6 = (view: any, shift = false) => view.scope.keys.find((k: any) => k.key === "F6" && k.mods.join() === (shift ? "Shift" : "")).fn();

  it("documents: one toolbar Tab stop, F6 between document and toolbar, Escape in the outline goes back", async () => {
    const view = mount(DocxView, "sample.docx") as any;
    const file = await open(view, "keys.docx");
    expect(tabStops(view.contentEl)).toHaveLength(1);
    expect(view.contentEl.querySelector(".dw-toolbar")!.getAttribute("role")).toBe("toolbar");
    view.editor.focus();
    f6(view, true);
    expect(view.contentEl.querySelector(".dw-toolbar")!.contains(document.activeElement)).toBe(true);
    f6(view);
    expect(view.editor.dom.contains(document.activeElement)).toBe(true);
    view.outlineEl.show();
    const entry = view.outlineEl.createEl("button", { text: "x" });
    entry.focus();
    entry.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(view.editor.dom.contains(document.activeElement)).toBe(true);
    await view.onUnloadFile(file);
  });

  it("slides: one toolbar Tab stop, F6 visits thumbnails, slide and notes; Escape in the format panel goes back", async () => {
    const view = mount(PptxView, "sample.pptx") as any;
    const file = await open(view, "keys.pptx");
    expect(tabStops(view.contentEl)).toHaveLength(1);
    view.stageEl.focus();
    f6(view);
    expect(view.propsEl.contains(document.activeElement) || document.activeElement === view.notesEl).toBe(true);
    f6(view, true);
    expect(document.activeElement).toBe(view.stageEl);
    f6(view, true);
    expect(document.activeElement).toBe(view.thumbsEl);
    view.showProps = true;
    view.propsEl.show();
    const field = view.propsEl.createEl("button", { text: "x" });
    field.focus();
    field.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(document.activeElement).toBe(view.stageEl);
    await view.onUnloadFile(file);
  });

  it("returns focus to the document when a dialog closes", async () => {
    const view = mount(DocxView, "sample.docx") as any;
    (view.app as any).workspace = { getActiveViewOfType: () => view };
    const file = await open(view, "modal.docx");
    const modal = new Modal(view.app as any);
    modal.open();
    (document.activeElement as HTMLElement | null)?.blur();
    document.body.appendChild(createDiv({ cls: "modal-container" })).remove();
    await new Promise((r) => setTimeout(r, 0));
    await new Promise((r) => setTimeout(r, 0));
    expect(view.editor.dom.contains(document.activeElement)).toBe(true);
    modal.close();
    await view.onUnloadFile(file);
  });
});

describe("DocxView menus", () => {
  it("has Google's menus with every command", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "menus.docx");
    expect(menuNames(view)).toEqual(["File", "Edit", "View", "Insert", "Format", "Tools", "Help"]);
    expect(menuTitles(view, "File")).toEqual(["Export…"]);
    expect(menuTitles(view, "Edit")).toEqual(["Undo", "Redo", "Cut", "Copy", "Paste", "Paste without formatting", "Select all", "Copy link to heading", "Delete", "Find and replace"]);
    expect(menuTitles(view, "View")).toEqual(["Editing", "Suggesting", "Page view", "Ruler", "Outline", "Review panel", "Zoom", "Light", "Dark"]);
    expect(menuTitles(view, "Insert")).toEqual(["Image", "Table", "Drawing", "Chart", "Equation", "Special characters", "Horizontal line", "Break", "Footnote", "Comment", "Link", "Table of contents", "Header & footer", "Watermark"]);
    expect(menuTitles(view, "Format")).toEqual(["Text", "Paragraph style", "Align", "Line spacing", "Paragraph…", "Columns…", "Bulleted list", "Numbered list", "Clear formatting"]);
    expect(menuTitles(view, "Tools")).toEqual(["Paint format", "Word count", "Find and replace"]);
    await view.onUnloadFile(file);
  });

  it("runs commands from the menus, with checks showing the current state", async () => {
    const { TextSelection } = await import("prosemirror-state");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "menus.docx");
    const editor = (view as any).editor;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 4)));
    runMenu(view, "Format", "Text", "Bold");
    expect(editor.state.doc.rangeHasMark(1, 4, editor.state.schema.marks.bold)).toBe(true);
    menuTitles(view, "View");
    expect(Menu.last!.items.find((i) => i.title === "Light")!.checked).toBe(true);
    runMenu(view, "View", "Dark");
    expect(view.contentEl.hasClass("docx-theme-dark")).toBe(true);
    menuTitles(view, "View");
    expect(Menu.last!.items.find((i) => i.title === "Dark")!.checked).toBe(true);
    await view.onUnloadFile(file);
  });

  it("keeps everyday controls in the row and moves the rest to menus", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "menus.docx");
    const label = (l: string) => view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`);
    for (const l of ["Undo (Ctrl+Z)", "Paragraph style", "Font", "Font size", "Bold (Ctrl+B)", "Text color", "Highlight color", "Insert image", "Align", "Line spacing", "Bulleted list (Ctrl+Shift+8)", "Clear formatting"]) expect(label(l), l).not.toBeNull();
    for (const l of ["Document outline", "Show ruler", "Subscript", "Table", "Find and replace (Ctrl+F)"]) expect(label(l), l).toBeNull();
    expect(view.contentEl.querySelector(".dw-segmented")).toBeNull();
    // The mode select is the right-most group (it folds into ⋮ first); the word count sits with the save status.
    expect(view.contentEl.querySelector('.dw-push [aria-label="Editing mode"]')).not.toBeNull();
    expect(view.contentEl.querySelector(".dw-end .docx-words")).not.toBeNull();
    await view.onUnloadFile(file);
  });

  it("shows picture tools while a picture is selected, and table tools in a table", async () => {
    const { NodeSelection, TextSelection } = await import("prosemirror-state");
    const { cmd } = await import("../src/presentation/docx/docx-editor");
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "menus.docx");
    const editor = (view as any).editor;
    const s = editor.state.schema;
    editor.dispatch(editor.state.tr.insert(1, s.nodes.image.create({ src: "data:,", width: 40, height: 20 })));
    const group = (l: string) => view.contentEl.querySelector<HTMLElement>(`.dw-toolbar [aria-label="${l}"]`)!.closest<HTMLElement>(".dw-group")!;
    expect(group("Crop image").hidden).toBe(true);
    expect(group("Cell fill").hidden).toBe(true);
    editor.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 1)));
    expect(group("Crop image").hidden).toBe(false);
    const wrap = view.contentEl.querySelector<HTMLSelectElement>('.dw-toolbar select[aria-label="Text wrapping"]')!;
    wrap.value = "square";
    wrap.dispatchEvent(new Event("change"));
    expect(editor.state.doc.nodeAt(1).attrs.wrap).toBe("square");
    const end = editor.state.doc.content.size;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(end - 1))));
    cmd.insertTable(2, 2)(editor.state, editor.dispatch);
    expect(group("Crop image").hidden).toBe(true);
    expect(group("Cell fill").hidden).toBe(false);
    await view.onUnloadFile(file);
  });

  it("opens the review panel as a drawer in a narrow pane, leaving its docked state alone", async () => {
    const view = mount(DocxView, "sample.docx");
    const file = await open(view, "menus.docx");
    const v = view as any;
    const wasOpen = v.reviewOpen;
    Object.defineProperty(view.contentEl, "clientWidth", { configurable: true, value: 500 });
    v.narrowWatch.check();
    expect(view.contentEl.hasClass("is-narrow")).toBe(true);
    runMenu(view, "View", "Review panel");
    const review = view.contentEl.querySelector<HTMLElement>(".docx-review")!;
    expect(review.hasClass("is-drawer")).toBe(true);
    expect(review.childElementCount).toBeGreaterThan(0);
    expect(v.reviewOpen).toBe(wasOpen);
    review.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(review.hasClass("is-drawer")).toBe(false);
    expect(review.style.display === "none").toBe(!wasOpen);
    runMenu(view, "View", "Outline");
    const outline = view.contentEl.querySelector<HTMLElement>(".docx-outline")!;
    expect(outline.hasClass("is-drawer")).toBe(true);
    expect(outline.textContent).toContain("Document outline");
    await view.onUnloadFile(file);
  });
});

describe("DocxView context menus", () => {
  async function setup(name: string) {
    const { Menu } = await import("./obsidian-stub");
    const { TextSelection } = await import("prosemirror-state");
    const view = mount(DocxView, "sample.docx") as any;
    const file = await open(view, name);
    const ed = view.editor;
    // A fresh paragraph at the end, so the tests control its content.
    const end = ed.state.doc.content.size;
    const rightClick = (pos: number) => {
      const dom = ed.domAtPos(pos);
      const el = (dom.node.nodeType === 3 ? dom.node.parentElement : dom.node) as HTMLElement;
      Menu.last = null;
      const e = new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 1, clientY: 1 });
      el.dispatchEvent(e);
      return { prevented: e.defaultPrevented, items: Menu.last ? (Menu.last as any).items.map((i: any) => i.title) : [], run: (t: string) => (Menu.last as any).items.find((i: any) => i.title === t).run() };
    };
    const select = (from: number, to = from) => ed.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, from, to)));
    return { view, file, ed, end, rightClick, select };
  }

  it("leaves plain text to the native menu (spelling included)", async () => {
    const { view, file, ed, rightClick } = await setup("ctx-plain.docx");
    let at = -1;
    ed.state.doc.descendants((n: any, pos: number) => (at < 0 && n.isText ? ((at = pos + 1), false) : at < 0));
    const r = rightClick(at);
    expect(r.prevented).toBe(false);
    expect(r.items).toEqual([]);
    await view.onUnloadFile(file);
  });

  it("offers table actions in a table, with spelling fixes for a misspelled word first", async () => {
    const { view, file, ed, end, rightClick, select } = await setup("ctx-table.docx");
    const { cmd } = await import("../src/presentation/docx/docx-editor");
    select(end - 1);
    cmd.insertTable(2, 2)(ed.state, ed.dispatch);
    let cell = -1;
    ed.state.doc.descendants((n: any, pos: number) => (n.type.name === "table_cell" && pos > end - 2 && cell < 0 ? ((cell = pos), false) : cell < 0));
    ed.dispatch(ed.state.tr.insertText("teh cat", cell + 2));
    view.speller = { misspelled: (w: string) => w === "teh", suggestions: () => ["the", "ten"], addToDictionary: vi.fn() };
    select(cell + 3);
    const r = rightClick(cell + 3);
    expect(r.prevented).toBe(true);
    expect(r.items.slice(0, 3)).toEqual(["the", "ten", "Add to dictionary"]);
    expect(r.items).toEqual(expect.arrayContaining(["Cut", "Copy", "Paste", "Insert row above", "Insert row below", "Insert column left", "Insert column right", "Delete row", "Delete column", "Split cell", "Delete table"]));
    r.run("the");
    expect(ed.state.doc.nodeAt(cell)!.textContent).toBe("the cat");
    rightClick(cell + 3).run("Insert row below");
    let rows = 0;
    ed.state.doc.descendants((n: any, pos: number) => void (n.type.name === "table_row" && pos > end - 2 && rows++));
    expect(rows).toBe(3);
    await view.onUnloadFile(file);
  });

  it("shades a selection of cells from the right-click menu, with a custom colour", async () => {
    const { view, file, ed, end, rightClick, select } = await setup("ctx-shade.docx");
    const { cmd } = await import("../src/presentation/docx/docx-editor");
    select(end - 1);
    cmd.insertTable(2, 2)(ed.state, ed.dispatch);
    const cells: number[] = [];
    ed.state.doc.descendants((n: any, pos: number) => void (n.type.name === "table_cell" && pos > end - 2 && cells.push(pos)));
    const fills = () => cells.map((p) => ed.state.doc.nodeAt(p)!.attrs.background);
    // All four, as dragging from the first cell to the last selects; right-clicking the last keeps them.
    const { CellSelection } = await import("../src/presentation/docx/cell-selection");
    ed.dispatch(ed.state.tr.setSelection(CellSelection.create(ed.state.doc, cells[0], cells[3])));
    const r = rightClick(cells[3] + 2);
    expect(ed.state.selection).toBeInstanceOf(CellSelection);
    const shading = (Menu.last as any).items.find((i: any) => i.title === "Cell shading").submenu;
    expect(shading.items.map((i: any) => i.title)).toEqual(["None", "Blue", "Gray", "Yellow", "Green", "Custom…"]);
    expect(r.prevented).toBe(true);
    shading.items.find((i: any) => i.title === "Custom…").run();
    const picker = document.body.querySelector<HTMLInputElement>("input.dw-color-temp")!;
    picker.value = "#ff8800";
    picker.dispatchEvent(new Event("change"));
    expect(fills()).toEqual(["#FF8800", "#FF8800", "#FF8800", "#FF8800"]);
    await view.onUnloadFile(file);
  });

  it("offers link actions on a link", async () => {
    const { view, file, ed, end, rightClick, select } = await setup("ctx-link.docx");
    ed.dispatch(ed.state.tr.insertText("site", end - 1).addMark(end - 1, end + 3, ed.state.schema.marks.link.create({ href: "https://example.com/" })));
    select(end + 1);
    const r = rightClick(end + 1);
    expect(r.items).toEqual(expect.arrayContaining(["Open link", "Edit link…", "Remove link"]));
    r.run("Remove link");
    expect(ed.state.doc.rangeHasMark(end - 1, end + 3, ed.state.schema.marks.link)).toBe(false);
    await view.onUnloadFile(file);
  });

  it("puts a comment card's actions on its right-click menu", async () => {
    const { Menu } = await import("./obsidian-stub");
    const { view, file, ed, end, select, rightClick } = await setup("ctx-comment.docx");
    const { cmd } = await import("../src/presentation/docx/docx-editor");
    ed.dispatch(ed.state.tr.insertText("noted", end - 1));
    const id = view.docx.addComment("Check this", "Tester");
    select(end - 1, end + 4);
    cmd.addCommentAnchors(id)(ed.state, ed.dispatch);
    view.setReview(true);
    const card = Array.from(view.reviewEl.querySelectorAll(".docx-review-card") as NodeListOf<HTMLElement>).find((c) => c.dataset.commentId === id)!;
    const buttons = Array.from(card.querySelectorAll(".docx-review-actions button")).map((b) => b.textContent);
    card.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect((Menu.last as any).items.map((i: any) => i.title)).toEqual(buttons);
    // Right-clicking the commented text offers the same actions.
    const mark = view.editor.dom.querySelector(`[data-comment="${id}"]`) as HTMLElement;
    expect(mark).not.toBeNull();
    Menu.last = null;
    mark.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect((Menu.last as any).items.map((i: any) => i.title)).toEqual(expect.arrayContaining(["Resolve", "Reply…", "Delete comment"]));
    void rightClick;
    await view.onUnloadFile(file);
  });
});

describe("PptxView context menus", () => {
  const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
  const rightClick = async (el: HTMLElement) => {
    const { Menu } = await import("./obsidian-stub");
    Menu.last = null;
    el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    const items = (Menu.last as any).items as { title: string; run: () => void }[];
    return { titles: items.map((i) => i.title), run: (t: string) => items.find((i) => i.title === t)!.run() };
  };

  it("adds table actions when right-clicking a table cell", async () => {
    const { addTable } = await import("../src/domain/pptx/table");
    const view = mount(PptxView, "sample.pptx") as any;
    const file = await open(view, "ctx.pptx");
    view.mutate((doc: any) => addTable(doc.xml(doc.slides()[view.index].path), 2, 2, { x: 0, y: 0, w: 3000000, h: 1000000 }, null), "slide");
    const cell = view.rendered.find((r: any) => r.shape.kind === "cell");
    const menu = await rightClick(cell.el);
    expect(menu.titles).toEqual(expect.arrayContaining(["Cut", "Copy", "Insert row above", "Insert row below", "Insert column left", "Insert column right", "Delete row", "Delete column"]));
    menu.run("Insert row below");
    const table = view.rendered.find((r: any) => r.shape.table)?.shape.table;
    expect(table.rows).toHaveLength(3);
    await view.onUnloadFile(file);
  });

  it("adds picture actions when right-clicking a picture", async () => {
    const { addPicture } = await import("../src/domain/pptx/slide-edit");
    const { setPictureAdjust, NO_ADJUST } = await import("../src/domain/pptx/picture-adjust");
    const view = mount(PptxView, "sample.pptx") as any;
    const file = await open(view, "ctx-pic.pptx");
    view.mutate((doc: any) => {
      const path = doc.slides()[view.index].path;
      const pic = addPicture(doc.xml(path), { x: 0, y: 0, w: 914400, h: 914400 }, doc.addImage(path, PNG, "png"), "Photo");
      setPictureAdjust(pic, { ...NO_ADJUST, recolor: "gray" });
    }, "slide");
    const pic = view.rendered.find((r: any) => r.shape.kind === "pic");
    const menu = await rightClick(pic.el);
    expect(menu.titles).toEqual(expect.arrayContaining(["Replace image…", "Reset crop", "Reset colour adjustments"]));
    menu.run("Reset colour adjustments");
    expect(view.rendered.find((r: any) => r.shape.kind === "pic").shape.adjust.recolor).toBe("none");
    await view.onUnloadFile(file);
  });
});
