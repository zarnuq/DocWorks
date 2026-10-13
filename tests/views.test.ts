import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, localData, Modal, TFile, WorkspaceLeaf } from "./obsidian-stub";
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

describe("PptxView motion editor", () => {
  it("keeps slide creation in the Slide dropdown without a duplicate toolbar button", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "motion.pptx");
    const labels = Array.from(view.contentEl.querySelectorAll(".dw-toolbar button")).map((b) => b.textContent);
    expect(labels).toContain("Slide ▾");
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
    button("Motion").click();
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
    tb.end("light", () => {}, "paper");
    const groups = Array.from(tb.el.querySelectorAll(".dw-group")).filter((g) => g.childElementCount);
    expect(groups.map((g) => g.childElementCount)).toEqual([2, 1]);
    expect(tb.el.lastElementChild!.classList.contains("dw-end")).toBe(true);
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
    button(view.contentEl, "Drawing").click();
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

describe("DocxView comment author", () => {
  it("records editor suggestions and persists them, and displays navigable headings and a ruler", async () => {
    const app = new App();
    app.vault.readBinary = async () => fixture("sample.docx").buffer as ArrayBuffer;
    const saved = vi.spyOn(app.vault, "modifyBinary");
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, pageSize: "document" as const, author: "Ada" });
    document.body.appendChild(view.containerEl);
    const file = await open(view, "suggestions.docx");
    const v = view as any;
    const mode = view.contentEl.querySelector<HTMLSelectElement>('[aria-label="Editing mode"]')!;
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
    const outline = view.contentEl.querySelector<HTMLButtonElement>('[aria-label="Document outline"]');
    // Toolbar icons use title rather than aria-label.
    (outline ?? view.contentEl.querySelector<HTMLButtonElement>('button[title="Document outline"]'))!.click();
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
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Theme")!.click();expect(saved).not.toHaveBeenCalled();
  const select=view.contentEl.querySelector<HTMLSelectElement>('[aria-label="Edit master or layout"]')!;
  select.value=select.options[1].value;select.dispatchEvent(new Event("change"));
  const before=v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length;
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Text box")!.click();
  expect(v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length).toBe(before+1);
  v.undo();expect(v.currentSlideXml().getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main","sp").length).toBe(before);
  await view.onUnloadFile(file);
 });
 it("inserts a formatted chart: stacked bars with a title and no legend, then changes it to a pie",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Chart")!.click();
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
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Chart")!.click();
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
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Chart")!.click();
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
    expect(view.contentEl.querySelector('[aria-label="Light paper"]')?.getAttribute("aria-pressed")).toBe("true");
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

describe("PptxView grid view", () => {
  it("toggles a grid of slides from the toolbar and Ctrl+Alt+1; double-clicking a slide opens it", async () => {
    const view = mount(PptxView, "sample.pptx");
    const file = await open(view, "grid.pptx");
    const v = view as any;
    const toggle = view.contentEl.querySelector<HTMLElement>('[aria-label="Grid view (Ctrl+Alt+1)"]')!;
    toggle.click();
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
