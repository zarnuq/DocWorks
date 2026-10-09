import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DocxView, DocWorksSettings } from "../src/presentation/docx/docx-view";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";

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

describe("DocxView switching files", () => {
  it("pending word-count and review timers don't fire after unload", async () => {
    window.localStorage.setItem("doc-works-review-panel", "on");
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
    window.localStorage.removeItem("doc-works-review-panel");
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
    const settings: DocWorksSettings = { pageSize: "letter" };
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
    expect(dialog).toHaveBeenCalledTimes(1);
  });
});
