import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, localData, Modal, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DocxView, DocWorksSettings } from "../src/presentation/docx/docx-view";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { readAnimations, readTransition } from "../src/domain/pptx/animation";

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

describe("DocxView page size setting", () => {
  it("lays out an A4 document on Letter pages when the setting says so, and switches live", async () => {
    const docx = await DocxDocument.load(fixture("sample.docx"));
    const sectPr = Array.from(docx.body.children).filter((c) => c.localName === "sectPr").pop()!;
    const pgSz = sectPr.getElementsByTagNameNS(W, "pgSz")[0];
    pgSz.setAttributeNS(W, "w:w", "11906");
    pgSz.setAttributeNS(W, "w:h", "16838");
    const bytes = await docx.save();
    const settings: DocWorksSettings = { pageSize: "letter", author: "" };
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
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { pageSize: "document", author: "Ada" });
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
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { pageSize: "document", author: "Ada" });
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
    const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { pageSize: "document", author: "Ada" });
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
    const settings: DocWorksSettings = { pageSize: "document", author: "Ada Lovelace" };
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
 it("inserts an editable chart from its data grid and supports undo",async()=>{
  const view=mount(PptxView,"sample.pptx");const file=await open(view,"authoring.pptx");const v=view as any;
  Array.from(view.contentEl.querySelectorAll("button")).find(b=>b.textContent==="Chart")!.click();
  const modal=document.querySelector<HTMLElement>(".test-modal")!;
  const field=modal.querySelector<HTMLInputElement>('input[aria-label="Series 1, row 1"]')!;field.value="42";field.dispatchEvent(new Event("input"));
  Array.from(modal.querySelectorAll("button")).find(b=>b.textContent==="Insert chart")!.click();
  await vi.waitFor(()=>expect(v.model.shapes.find((s:any)=>s.chart)?.chart.data.series[0].values[0]).toBe(42));
  expect(view.contentEl.querySelector('svg[aria-label="Chart"]')).not.toBeNull();
  const shape=v.model.shapes.find((s:any)=>s.chart);v.setSelection([shape.el]);
  expect(view.contentEl.textContent).toContain("Edit chart data");
  v.undo();expect(v.model.shapes.some((s:any)=>s.chart)).toBe(false);
  await view.onUnloadFile(file);
 });
});
