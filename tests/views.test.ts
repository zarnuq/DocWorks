import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DocxView } from "../src/presentation/docx/docx-view";
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
