import { FileView, Notice, TFile, WorkspaceLeaf } from "obsidian";
import { PptxDocument, PptxSnapshot } from "../../infrastructure/pptx/pptx-document";
import { Align, Rect, RunProps, SlideModel, parseSlide, resolveLevel } from "../../domain/pptx/slide-model";
import { ParaInput, addPicture, addTextBox, setRect, writeParagraphs } from "../../domain/pptx/slide-edit";
import { History } from "../../core/utils/history";
import { FONTS, PaperTheme, Toolbar } from "../common/toolbar";
import { nextVisible, Slideshow } from "./slideshow";
import { RenderedShape, renderSlide } from "./slide-renderer";
import { applyRunStyle, insertPlainText, markEmptyParagraphs, readParagraphs, renderParagraphs, selectionRunProps, updateParagraphs } from "./text-html";

export const VIEW_TYPE_PPTX = "doc-works-pptx-view";

const THUMB_WIDTH = 168;
const FONT_SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 96];
const IMAGE_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
};
const MIN_SIZE_EMU = 91440;

type Snapshot = { snap: PptxSnapshot; index: number };

const isMod = (e: KeyboardEvent) => e.ctrlKey || e.metaKey;
const extOf = (name: string) => name.slice(name.lastIndexOf(".") + 1).toLowerCase();

export class PptxView extends FileView {
  allowNoFile = false;
  paperTheme: "light" | "dark" = "light";

  private doc: PptxDocument | null = null;
  private index = 0;
  private model: SlideModel | null = null;
  private rendered: RenderedShape[] = [];
  private scale = 1;
  private selected: RenderedShape | null = null;
  private activeText: RenderedShape | null = null;
  private savedRange: Range | null = null;
  private history = new History<Snapshot>();
  private imageUrls = new Map<string, string>();
  private dirty = false;
  private saveTimer: number | null = null;
  private saving: Promise<void> = Promise.resolve();
  private commitTimer: number | null = null;
  private pendingCommit: RenderedShape | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private lastStageWidth = 0;

  private thumbsEl!: HTMLElement;
  private stageEl!: HTMLElement;
  private slideHostEl!: HTMLElement;
  private notesEl!: HTMLTextAreaElement;
  private notesTimer: number | null = null;
  private hideBtn: HTMLButtonElement | null = null;
  private saveIndicatorEl: HTMLElement | null = null;
  private syncTheme: ((t: PaperTheme) => void) | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_PPTX;
  }

  getIcon(): string {
    return "presentation";
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Pptx";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "pptx";
  }

  // ---------------------------------------------------------------------------
  // Lifecycle
  // ---------------------------------------------------------------------------

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("pptx-editor-root");
    const status = this.contentEl.createDiv({ cls: "dw-loading", text: "Loading presentation..." });

    try {
      this.doc = await PptxDocument.load(await this.app.vault.readBinary(file));
      await this.loadMedia();
    } catch (err) {
      status.setText(`Could not open this .pptx file: ${(err as Error).message}`);
      console.error(err);
      this.doc = null;
      return;
    }
    status.remove();

    this.buildToolbar(this.contentEl.createDiv({ cls: "dw-toolbar pptx-toolbar" }));
    const body = this.contentEl.createDiv({ cls: "pptx-body" });
    this.thumbsEl = body.createDiv({ cls: "pptx-thumbs" });
    this.stageEl = body.createDiv({ cls: "pptx-stage" });
    this.stageEl.tabIndex = 0;
    this.slideHostEl = this.stageEl.createDiv({ cls: "pptx-slide-host" });
    this.notesEl = this.stageEl.createEl("textarea", { cls: "pptx-notes", attr: { "aria-label": "Speaker notes" } });
    this.notesEl.addEventListener("input", () => {
      if (this.notesTimer) window.clearTimeout(this.notesTimer);
      this.notesTimer = window.setTimeout(() => this.flushNotes(), 400);
    });
    this.notesEl.addEventListener("blur", () => this.flushNotes());

    this.stageEl.addEventListener("keydown", this.onStageKey);
    this.stageEl.addEventListener("pointerdown", (e) => {
      if (e.target === this.stageEl || e.target === this.slideHostEl) this.select(null);
    });
    this.stageEl.addEventListener("paste", this.onPaste);
    this.stageEl.addEventListener("dragover", (e) => e.preventDefault());
    this.stageEl.addEventListener("drop", (e) => {
      const image = Array.from(e.dataTransfer?.files ?? []).find((f) => f.type.startsWith("image/"));
      if (!image) return;
      e.preventDefault();
      void this.insertImageFile(image);
    });
    document.addEventListener("selectionchange", this.onSelectionChange);
    window.addEventListener("blur", this.onWindowBlur);
    this.resizeObserver = new ResizeObserver(() => {
      // Re-render only on real width changes, never while the user is typing.
      const w = this.stageEl.clientWidth;
      if (Math.abs(w - this.lastStageWidth) > 8 && !this.pendingCommit && !this.activeTextFocused()) this.renderCurrent();
    });
    this.resizeObserver.observe(this.stageEl);

    this.index = 0;
    this.setPaperTheme(this.paperTheme);
    this.renderAll();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    document.removeEventListener("selectionchange", this.onSelectionChange);
    window.removeEventListener("blur", this.onWindowBlur);
    this.resizeObserver?.disconnect();
    this.flushCommit();
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    await this.saving;
    for (const url of this.imageUrls.values()) URL.revokeObjectURL(url);
    this.imageUrls.clear();
    this.doc = null;
    this.rendered = [];
    this.history = new History<Snapshot>();
  }

  /** Force an immediate save, e.g. before Obsidian quits. */
  async requestSave(): Promise<void> {
    this.flushCommit();
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
  }

  private onWindowBlur = () => {
    this.flushCommit();
    if (this.dirty) void this.saveNow();
  };

  private async loadMedia(): Promise<void> {
    for (const path of this.doc!.mediaPaths()) {
      const mime = IMAGE_MIME[extOf(path)];
      if (!mime || this.imageUrls.has(path)) continue;
      const bytes = await this.doc!.media(path);
      this.imageUrls.set(path, URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime })));
    }
  }

  private imageUrl = (path: string) => this.imageUrls.get(path) ?? null;

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  private renderAll(): void {
    this.renderThumbs();
    this.renderCurrent();
  }

  private renderThumbs(): void {
    if (!this.doc) return;
    this.thumbsEl.empty();
    this.doc.slides().forEach((slide, i) => {
      const item = this.thumbsEl.createDiv({ cls: "pptx-thumb" });
      item.toggleClass("is-active", i === this.index);
      item.toggleClass("is-hidden", this.doc!.isHidden(i));
      item.createDiv({ cls: "pptx-thumb-num", text: String(i + 1) });
      const host = item.createDiv({ cls: "pptx-thumb-slide" });
      renderSlide(host, parseSlide(this.doc!, slide.path), this.doc!.slideSize, {
        width: THUMB_WIDTH,
        editable: false,
        imageUrl: this.imageUrl,
      });
      item.onclick = () => this.goTo(i);
    });
  }

  private refreshThumb(i: number): void {
    const host = this.thumbsEl.children[i]?.querySelector<HTMLElement>(".pptx-thumb-slide");
    const slide = this.doc?.slides()[i];
    if (!host || !slide || !this.doc) return;
    renderSlide(host, parseSlide(this.doc, slide.path), this.doc.slideSize, { width: THUMB_WIDTH, editable: false, imageUrl: this.imageUrl });
  }

  private renderCurrent(): void {
    if (!this.doc) return;
    const slides = this.doc.slides();
    this.index = Math.max(0, Math.min(this.index, slides.length - 1));
    this.model = parseSlide(this.doc, slides[this.index].path);

    this.lastStageWidth = this.stageEl.clientWidth;
    const width = Math.max(320, Math.min((this.lastStageWidth || 1008) - 48, 1400));
    this.scale = width / this.doc.slideSize.cx;
    this.rendered = renderSlide(this.slideHostEl, this.model, this.doc.slideSize, { width, editable: true, imageUrl: this.imageUrl });
    this.rendered.forEach((rs) => this.wireShape(rs));

    const canEdit = this.doc.canEditNotes(this.index);
    this.notesEl.value = this.doc.notesText(this.index);
    this.notesEl.disabled = !canEdit;
    this.notesEl.placeholder = canEdit ? "Click to add speaker notes" : "This presentation has no notes master, so notes can't be added.";
    this.hideBtn?.setText(this.doc.isHidden(this.index) ? "Show slide" : "Hide slide");
    this.selected = null;
    this.activeText = null;
    this.savedRange = null;
    Array.from(this.thumbsEl.children).forEach((c, i) => c.toggleClass("is-active", i === this.index));
    this.thumbsEl.children[this.index]?.scrollIntoView({ block: "nearest" });
  }

  /** Re-parse the current slide after an XML edit without touching the live DOM. */
  private refreshModel(): void {
    if (!this.doc || !this.model) return;
    this.model = parseSlide(this.doc, this.model.path);
    this.rendered.forEach((rs, i) => {
      if (this.model!.shapes[i]) rs.shape = this.model!.shapes[i];
    });
  }

  private wireShape(rs: RenderedShape): void {
    const { el, textEl, shape } = rs;
    if (shape.inherited) return;

    el.addEventListener("pointerdown", (e) => {
      this.select(rs);
      if (textEl?.contains(e.target as Node)) return;
      this.stageEl.focus({ preventScroll: true });
      if (!shape.grouped) this.startDrag(e, rs, "move");
    });

    if (!shape.grouped) {
      for (const mode of ["move", "resize"] as const) {
        const handle = el.createDiv({ cls: `pptx-handle pptx-handle-${mode}` });
        handle.setAttribute("title", mode === "move" ? "Drag to move" : "Drag to resize");
        handle.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
          this.select(rs);
          this.stageEl.focus({ preventScroll: true });
          this.startDrag(e, rs, mode);
        });
      }
    }

    if (textEl) {
      textEl.addEventListener("input", () => {
        textEl.toggleClass("is-empty", !textEl.textContent);
        markEmptyParagraphs(textEl);
        this.scheduleCommit(rs);
      });
      textEl.addEventListener("focus", () => {
        this.activeText = rs;
        this.select(rs);
      });
      textEl.addEventListener("blur", () => this.flushCommit());
      textEl.addEventListener("keydown", (e) => this.onTextKey(e, rs));
    }
  }

  private select(rs: RenderedShape | null): void {
    this.selected?.el.removeClass("is-selected");
    this.selected = rs;
    rs?.el.addClass("is-selected");
  }

  private activeTextFocused(): boolean {
    return !!this.activeText?.textEl && this.activeText.textEl.contains(document.activeElement);
  }

  private startDrag(e: PointerEvent, rs: RenderedShape, mode: "move" | "resize"): void {
    e.preventDefault();
    const start = { x: e.clientX, y: e.clientY };
    const r0 = { ...rs.shape.rect };
    let next: Rect | null = null;

    const onMove = (ev: PointerEvent) => {
      const dx = (ev.clientX - start.x) / this.scale;
      const dy = (ev.clientY - start.y) / this.scale;
      next =
        mode === "move"
          ? { ...r0, x: r0.x + dx, y: r0.y + dy }
          : { ...r0, w: Math.max(MIN_SIZE_EMU, r0.w + dx), h: Math.max(MIN_SIZE_EMU, r0.h + dy) };
      Object.assign(rs.el.style, {
        left: `${next.x * this.scale}px`,
        top: `${next.y * this.scale}px`,
        width: `${next.w * this.scale}px`,
        height: `${next.h * this.scale}px`,
      });
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      if (!next || !this.doc) return;
      this.flushCommit();
      this.pushHistory();
      setRect(rs.shape.el, next);
      this.refreshModel();
      this.markDirty();
      this.refreshThumb(this.index);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  // ---------------------------------------------------------------------------
  // Text editing
  // ---------------------------------------------------------------------------

  private defaultsFor(rs: RenderedShape): (lvl: number) => RunProps {
    return (lvl) => resolveLevel(rs.shape.text!.style, lvl).defaults;
  }

  private onSelectionChange = () => {
    const sel = window.getSelection();
    if (!sel?.rangeCount) return;
    const range = sel.getRangeAt(0);
    const rs = this.rendered.find((r) => r.textEl?.contains(range.commonAncestorContainer));
    if (rs) {
      this.savedRange = range.cloneRange();
      this.activeText = rs;
    }
  };

  private scheduleCommit(rs: RenderedShape): void {
    if (this.pendingCommit && this.pendingCommit !== rs) this.flushCommit();
    this.pendingCommit = rs;
    this.updateSaveStatus(true);
    if (this.commitTimer) window.clearTimeout(this.commitTimer);
    this.commitTimer = window.setTimeout(() => this.flushCommit(), 400);
  }

  /** Write pending speaker-notes edits for the current slide. */
  private flushNotes(): void {
    if (this.notesTimer) window.clearTimeout(this.notesTimer);
    this.notesTimer = null;
    if (!this.doc || this.notesEl.disabled || this.notesEl.value === this.doc.notesText(this.index)) return;
    this.pushHistory();
    this.doc.setNotes(this.index, this.notesEl.value);
    this.markDirty();
  }

  private flushCommit(): void {
    this.flushNotes();
    const rs = this.pendingCommit;
    if (this.commitTimer) window.clearTimeout(this.commitTimer);
    this.commitTimer = null;
    this.pendingCommit = null;
    if (rs) this.commitText(rs);
  }

  private commitText(rs: RenderedShape, paras?: ParaInput[]): void {
    if (!this.doc || !rs.textEl || !rs.shape.text) return;
    this.pushHistory();
    writeParagraphs(rs.shape.text, paras ?? readParagraphs(rs.textEl, rs.pxPerPt, this.defaultsFor(rs)));
    this.refreshModel();
    this.markDirty();
    this.refreshThumb(this.index);
  }

  /**
   * Run a formatting command against the active text box. Restores the
   * selection the toolbar click stole; with only a caret, formats the whole box.
   */
  private withText(fn: (rs: RenderedShape, textEl: HTMLElement) => void, wholeIfCollapsed = true): RenderedShape | null {
    const rs = this.activeText ?? (this.selected?.textEl ? this.selected : null);
    const textEl = rs?.textEl;
    if (!rs || !textEl || !this.rendered.includes(rs)) {
      new Notice("Click into a text box first.");
      return null;
    }
    const sel = window.getSelection()!;
    let range = this.savedRange && textEl.contains(this.savedRange.commonAncestorContainer) ? this.savedRange : null;
    if (!range || (wholeIfCollapsed && range.collapsed)) {
      range = document.createRange();
      range.selectNodeContents(textEl);
    }
    sel.removeAllRanges();
    sel.addRange(range);

    fn(rs, textEl);
    if (this.pendingCommit === rs) this.pendingCommit = null;
    this.commitText(rs);
    return rs;
  }

  private toggle(prop: "b" | "i" | "u" | "strike"): void {
    this.withText((rs, el) => {
      const current = selectionRunProps(el, rs.pxPerPt, this.defaultsFor(rs));
      applyRunStyle(el, { [prop]: !current?.[prop] }, rs.pxPerPt);
    });
  }

  private setRunStyle(patch: Partial<RunProps>): void {
    this.withText((rs, el) => applyRunStyle(el, patch, rs.pxPerPt));
  }

  private updateParas(fn: Parameters<typeof updateParagraphs>[1]): void {
    const rs = this.withText((_, el) => updateParagraphs(el, fn), false);
    // Levels change indents and default sizes: redraw from the saved model.
    if (rs?.textEl && rs.shape.text) renderParagraphs(rs.textEl, rs.shape.text.paras, rs.pxPerPt, this.scale, rs.shape.text.lnSpcReduction);
  }

  private clearFormatting(): void {
    const rs = this.activeText ?? this.selected;
    if (!rs?.textEl || !rs.shape.text) return;
    const defaults = this.defaultsFor(rs);
    const paras = readParagraphs(rs.textEl, rs.pxPerPt, defaults).map((p) => ({
      ...p,
      runs: p.runs.map((r) => ({ ...defaults(p.lvl), text: r.text })),
    }));
    this.flushCommit();
    this.commitText(rs, paras);
    renderParagraphs(rs.textEl, rs.shape.text.paras, rs.pxPerPt, this.scale, rs.shape.text.lnSpcReduction);
  }

  private onTextKey(e: KeyboardEvent, rs: RenderedShape): void {
    const key = e.key.toLowerCase();
    if (isMod(e) && (key === "z" || key === "y")) {
      e.preventDefault();
      if (key === "y" || e.shiftKey) this.redo();
      else this.undo();
    } else if (isMod(e) && (key === "b" || key === "i" || key === "u")) {
      e.preventDefault();
      this.toggle(key);
    } else if (e.key === "Tab") {
      e.preventDefault();
      this.updateParas((p) => ({ ...p, lvl: p.lvl + (e.shiftKey ? -1 : 1) }));
    } else if (e.key === "Escape") {
      rs.textEl?.blur();
      this.stageEl.focus({ preventScroll: true });
    }
  }

  private onPaste = (e: ClipboardEvent) => {
    const image = Array.from(e.clipboardData?.files ?? []).find((f) => f.type.startsWith("image/"));
    if (image) {
      e.preventDefault();
      void this.insertImageFile(image);
      return;
    }
    // Only plain text goes into slides; foreign HTML would bring styles we can't save.
    const target = this.rendered.find((r) => r.textEl?.contains(e.target as Node))?.textEl;
    if (target) {
      e.preventDefault();
      insertPlainText(e.clipboardData?.getData("text/plain") ?? "");
      // Programmatic edits fire no input event; send one so the change is committed.
      target.dispatchEvent(new Event("input", { bubbles: true }));
    }
  };

  // ---------------------------------------------------------------------------
  // Slide & shape commands
  // ---------------------------------------------------------------------------

  private onStageKey = (e: KeyboardEvent) => {
    const target = e.target as HTMLElement;
    if (target.isContentEditable || target === this.notesEl) return;
    const key = e.key.toLowerCase();
    if (e.key === "F5") {
      e.preventDefault();
      this.startPresentation(!e.shiftKey);
      return;
    }
    if (isMod(e) && (key === "z" || key === "y")) {
      e.preventDefault();
      if (key === "y" || e.shiftKey) this.redo();
      else this.undo();
    } else if ((e.key === "Delete" || e.key === "Backspace") && this.selected) {
      e.preventDefault();
      this.deleteSelected();
    } else if (e.key === "PageDown" || e.key === "ArrowDown" || e.key === "ArrowRight") {
      e.preventDefault();
      this.goTo(this.index + 1);
    } else if (e.key === "PageUp" || e.key === "ArrowUp" || e.key === "ArrowLeft") {
      e.preventDefault();
      this.goTo(this.index - 1);
    } else if (e.key === "Escape") {
      this.select(null);
    }
  };

  private goTo(i: number): void {
    if (!this.doc || i < 0 || i >= this.doc.slides().length) return;
    this.flushCommit();
    this.index = i;
    this.renderCurrent();
    this.stageEl.focus({ preventScroll: true });
  }

  /** Run the slideshow from the first visible slide, or from the current one. */
  startPresentation(fromStart: boolean): void {
    const doc = this.doc;
    if (!doc) return;
    this.flushCommit();
    const count = doc.slides().length;
    const hidden = Array.from({ length: count }, (_, i) => doc.isHidden(i));
    const start = fromStart ? (nextVisible(hidden, -1, 1) ?? 0) : this.index;
    const deck = {
      size: doc.slideSize,
      count,
      isHidden: (i: number) => hidden[i],
      notes: (i: number) => doc.notesText(i),
      render: (host: HTMLElement, i: number, width: number) =>
        renderSlide(host, parseSlide(doc, doc.slides()[i].path), doc.slideSize, { width, editable: false, imageUrl: this.imageUrl }),
    };
    new Slideshow(deck, start, (i) => this.goTo(i)).open();
  }

  private pushHistory(): void {
    if (this.doc) this.history.push({ snap: this.doc.snapshot(), index: this.index });
  }

  /** Snapshot, apply a structural change, then redraw everything. */
  private mutate(fn: (doc: PptxDocument) => void): void {
    if (!this.doc) return;
    this.flushCommit();
    const before = this.doc.snapshot();
    const index = this.index;
    try {
      fn(this.doc);
    } catch (err) {
      this.doc.restore(before);
      this.index = index;
      new Notice((err as Error).message);
      return;
    }
    this.history.push({ snap: before, index });
    this.markDirty();
    this.renderAll();
  }

  private undo(): void {
    this.travel((cur) => this.history.undo(cur));
  }

  private redo(): void {
    this.travel((cur) => this.history.redo(cur));
  }

  private travel(step: (current: Snapshot) => Snapshot | null): void {
    if (!this.doc) return;
    this.flushCommit();
    const target = step({ snap: this.doc.snapshot(), index: this.index });
    if (!target) return;
    this.doc.restore(target.snap);
    this.index = target.index;
    this.markDirty();
    this.renderAll();
  }

  private currentSlideXml(): Document | null {
    const slide = this.doc?.slides()[this.index];
    return slide ? this.doc!.xml(slide.path) : null;
  }

  private addTextBox(): void {
    const { cx, cy } = this.doc?.slideSize ?? { cx: 0, cy: 0 };
    this.mutate(() => addTextBox(this.currentSlideXml()!, { x: cx * 0.25, y: cy * 0.42, w: cx * 0.5, h: cy * 0.16 }, "Text"));
    const rs = this.rendered[this.rendered.length - 1];
    if (rs?.textEl) {
      rs.textEl.focus();
      const range = document.createRange();
      range.selectNodeContents(rs.textEl);
      window.getSelection()?.removeAllRanges();
      window.getSelection()?.addRange(range);
    }
  }

  private deleteSelected(): void {
    const rs = this.selected;
    if (!rs || rs.shape.inherited) return;
    this.mutate(() => rs.shape.el.remove());
  }

  private promptInsertImage(): void {
    const input = document.body.createEl("input", { type: "file", attr: { accept: "image/*" } });
    input.setCssStyles({ display: "none" });
    input.onchange = () => {
      const file = input.files?.[0];
      if (file) void this.insertImageFile(file);
      input.remove();
    };
    input.click();
  }

  private async insertImageFile(file: File): Promise<void> {
    if (!this.doc) return;
    const ext = IMAGE_MIME[extOf(file.name)] ? extOf(file.name) : file.type.split("/")[1]?.replace("jpeg", "jpg") ?? "png";
    const bytes = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(new Blob([bytes], { type: IMAGE_MIME[ext] ?? file.type }));
    const natural = await new Promise<{ w: number; h: number }>((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth || 1, h: img.naturalHeight || 1 });
      img.onerror = () => resolve({ w: 4, h: 3 });
      img.src = url;
    });

    const { cx, cy } = this.doc.slideSize;
    const fit = Math.min((cx * 0.6) / natural.w, (cy * 0.6) / natural.h);
    const w = natural.w * fit;
    const h = natural.h * fit;
    this.mutate((doc) => {
      const path = doc.slides()[this.index].path;
      const rId = doc.addImage(path, bytes, ext);
      this.imageUrls.set(doc.relTarget(path, rId)!, url);
      addPicture(doc.xml(path)!, { x: (cx - w) / 2, y: (cy - h) / 2, w, h }, rId, file.name);
    });
  }

  // ---------------------------------------------------------------------------
  // Toolbar
  // ---------------------------------------------------------------------------

  private buildToolbar(bar: HTMLElement): void {
    // Keep the text selection when clicking toolbar buttons.
    bar.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).closest("button")) e.preventDefault();
    });
    const tb = new Toolbar(bar);
    const btn = (label: string, title: string, run: () => void, cls = "") => tb.button(label, title, run, cls);
    const sep = () => tb.gap();

    btn("↶", "Undo (Ctrl+Z)", () => this.undo());
    btn("↷", "Redo (Ctrl+Y)", () => this.redo());
    sep();

    tb.select("Font", "Font", FONTS.map((f) => [f, f]), (font) => this.setRunStyle({ font }));
    tb.select("Font size", "Size", FONT_SIZES.map((sz) => [`${sz} pt`, String(sz)]), (sz) => this.setRunStyle({ sz: Number(sz) }));
    sep();

    btn("B", "Bold (Ctrl+B)", () => this.toggle("b"), "bold");
    btn("I", "Italic (Ctrl+I)", () => this.toggle("i"), "italic");
    btn("U", "Underline (Ctrl+U)", () => this.toggle("u"), "underline");
    btn("S", "Strikethrough", () => this.toggle("strike"), "strike");
    tb.color("A", "Text color", "#000000", (color) => this.setRunStyle({ color }));
    btn("Tx", "Clear formatting", () => this.clearFormatting());
    sep();

    for (const [label, align] of [
      ["Left", "l"],
      ["Center", "ctr"],
      ["Right", "r"],
      ["Justify", "just"],
    ] as [string, Align][]) {
      btn(label, `Align ${label.toLowerCase()}`, () => this.updateParas((p) => ({ ...p, align })));
    }
    btn("• List", "Toggle bullets", () => this.updateParas((p) => ({ ...p, bullet: p.bullet ? null : "•" })));
    btn("Outdent", "Decrease list level (Shift+Tab)", () => this.updateParas((p) => ({ ...p, lvl: p.lvl - 1 })));
    btn("Indent", "Increase list level (Tab)", () => this.updateParas((p) => ({ ...p, lvl: p.lvl + 1 })));
    sep();

    btn("Text box", "Insert text box", () => this.addTextBox());
    btn("Image", "Insert image from file", () => this.promptInsertImage());
    btn("Delete shape", "Delete selected shape (Del)", () => this.deleteSelected());
    sep();

    btn("New slide", "Insert slide after this one", () => this.mutate((d) => (this.index = d.addSlide(this.index))));
    btn("Duplicate", "Duplicate slide", () => this.mutate((d) => (this.index = d.duplicateSlide(this.index))));
    btn("Delete slide", "Delete slide", () => this.mutate((d) => d.deleteSlide(this.index)));
    this.hideBtn = btn("Hide slide", "Skip this slide in the slideshow", () => this.mutate((d) => d.setHidden(this.index, !d.isHidden(this.index))));
    btn("▲", "Move slide up", () =>
      this.mutate((d) => {
        if (this.index > 0) d.moveSlide(this.index, --this.index);
      })
    );
    btn("▼", "Move slide down", () =>
      this.mutate((d) => {
        if (this.index < d.slides().length - 1) d.moveSlide(this.index, ++this.index);
      })
    );
    sep();
    btn("▶ Present", "Present from this slide (Shift+F5). F5 presents from the beginning.", () => this.startPresentation(false), "mod-cta");

    const end = tb.end(this.paperTheme, (t) => this.setPaperTheme(t), "workspace");
    this.syncTheme = end.syncTheme;
    this.saveIndicatorEl = end.status;
  }

  setPaperTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    this.contentEl.toggleClass("pptx-theme-dark", theme === "dark");
    this.contentEl.toggleClass("pptx-theme-light", theme === "light");
    this.syncTheme?.(theme);
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------

  private markDirty(): void {
    this.dirty = true;
    this.updateSaveStatus(true);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => void this.saveNow(), 800);
  }

  private saveNow(): Promise<void> {
    // Serialize saves: JSZip must not be written while a previous save is generating.
    this.saving = this.saving.then(async () => {
      if (!this.doc || !this.file) return;
      this.flushCommit();
      if (this.saveTimer) window.clearTimeout(this.saveTimer);
      const file = this.file;
      try {
        this.dirty = false;
        const bytes = await this.doc.save();
        await this.app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
      } catch (err) {
        this.dirty = true;
        console.error("Failed to save .pptx:", err);
        new Notice(`Failed to save ${file.name}: ${(err as Error).message}`);
      } finally {
        this.updateSaveStatus(this.dirty || !!this.pendingCommit);
      }
    });
    return this.saving;
  }

  private updateSaveStatus(dirty: boolean): void {
    this.saveIndicatorEl?.setText(dirty ? "Saving…" : "Saved");
  }
}
