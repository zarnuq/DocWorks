import { addSlideComment, editSlideComment, removeSlideThread, resolveSlideThread, slideComments } from "../../domain/pptx/comments";
import { addMedia, linkedMediaCandidates, MEDIA_MIME, setMediaPlayback } from "../../domain/pptx/media";
import { addChart, chartWorkbook, prepareChartEdit } from "../../domain/pptx/chart";
import { ChartKind, chartPalette, setChartFormat } from "../../domain/pptx/chart-format";
import { chartDialog, NEW_CHART_TYPES, newChartFormat } from "../common/chart-form";
import { App, Modal, FileView, Menu, Modifier, Notice, Scope, setIcon, Setting, TFile, WorkspaceLeaf } from "obsidian";
import { dropIndex, PptxDocument, PptxSnapshot, SlideClip } from "../../infrastructure/pptx/pptx-document";
import { Align, Rect, RunProps, ShapeModel, SlideModel, parseSlide, plainText, resolveLevel, slideColors } from "../../domain/pptx/slide-model";
import { JUMPS, linkLabel, linkTarget, setShapeLink } from "../../domain/pptx/links";
import { NS } from "../../core/utils/xml";
import { writeFill, writeLine } from "../../domain/pptx/shape-style";
import { applyFormat, copyFormat, ShapeFormat } from "../../domain/pptx/paint-format";
import { NO_ADJUST, replacePicture, setPictureAdjust } from "../../domain/pptx/picture-adjust";
import {
  CONNECTORS,
  Glue,
  ParaInput,
  addPicture,
  addShape,
  addTextBox,
  connect,
  lineBox,
  rerouteConnectors,
  setAdjust,
  setCrop,
  setRect,
  setTransform,
  shapeId,
  sitePoint,
  reorder,
  writeBodyPr,
  writeParagraphs,
} from "../../domain/pptx/slide-edit";
import { presetGeometry, solveAdjust } from "../../domain/pptx/geometry";
import { buildShapeGallery, ShapeChoice } from "./shape-gallery";
import { ArrangeControl, buildArrangePanel, buildShapePanel, buildSlidePanel, DASHES, PanelControl, SlideControl, spPrOf, tableActions } from "./shape-panel";
import { alignRects, Alignment, Clip, copyItems, distributeRects, groupItems, itemRect, pasteItems, setItemRect, snapRect, topLevel, ungroupItem } from "../../domain/pptx/arrange";
import { History } from "../../core/utils/history";
import { IMAGE_MIME } from "../../core/utils/images";
import { pickImage } from "../common/images";
import { findSlideMatches, SlideSearchMatch, replaceAll, replaceMatch } from "../../domain/pptx/find";
import { showFindPreview } from "../common/find-preview";
import { findBar } from "../common/find-bar";
import { addTable, CellRange, insertRow, setTableStyleId, TABLE_STYLES } from "../../domain/pptx/table";
import { runEditCommand } from "../common/spelling";
import { specialCharacters } from "../common/special-chars";
import { COMMON_SHORTCUTS, registerShortcutsKey, shortcutsDialog, SLIDES_SHORTCUTS } from "../common/shortcuts";
import { EDITING_KEYS, FONTS, menuItem, promptText, shieldEditingKeys, showValue, subMenu, Toolbar, addFileItems } from "../common/toolbar";
import { openDrawer, watchNarrow } from "../common/narrow";
import { Autosave } from "../common/autosave";
import { escapeToEditor, focusRegions, returnFocusAfterOverlays } from "../common/keyboard";

import { addSmartArtItem, removeSmartArtItem, setSmartArtText, smartArtItems, smartArtParts } from "../../domain/pptx/smartart";
import { DEFAULT_SETTINGS, DocWorksSettings, paperFor } from "../common/settings";
import { loadEmbeddedFonts } from "./embedded-fonts";
import { eotToTtf } from "../../infrastructure/fonts/eot";
import { nextVisible, Slideshow } from "./slideshow";
import { readAnimations, readTransition, writeTransition } from "../../domain/pptx/animation";
import { updateAnimationPart, addAnimation, animationEntries, moveAnimation, removeAnimation, updateAnimation } from "../../domain/pptx/animation-edit";
import { buildMotionPanel } from "./motion-panel";
import { canExportPdf, renderPdf } from "../common/pdf";
import { renderPng } from "../common/png";
import { exportDialog, writeExport } from "../common/export-dialog";
import { RenderedShape, renderSlide } from "./slide-renderer";
import { rangeInParagraph } from "./text-html";
import { applyRunStyle, insertField, insertPlainText, refreshMarks, readParagraphs, renderParagraphs, selectionRunProps, updateParagraphs } from "./text-html";

export const VIEW_TYPE_PPTX = "doc-works-pptx-view";

const THUMB_WIDTH = 168;
/** Panel padding (2 × 8), slide number column (14) and gap (6), thumbnail border (2 × 2). */
const THUMB_CHROME = 40;
/** Slide width in grid view. */
const GRID_THUMB_WIDTH = 240;
const ALIGNS: [string, Align][] = [["Left", "l"], ["Center", "ctr"], ["Right", "r"], ["Justify", "just"]];
const LINE_SPACINGS = ["1.0", "1.15", "1.5", "2.0", "2.5", "3.0"];
const CHAR_SPACINGS: [string, string][] = [["Very tight", "-3"], ["Tight", "-1.5"], ["Normal", "0"], ["Loose", "3"], ["Very loose", "6"]];
const FONT_SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 96];
const MIN_SIZE_EMU = 91440;
const EMU_PER_PT = 12700;
const EMU_PER_PX = 9525;
const EMU_PER_INCH = 914400;
/** Snap distance for connector ends, in px. */
const SNAP_PX = 10;
const PROPS_KEY = "doc-works-format-panel";
/** CSS custom highlights (typed as a map; the DOM lib in use lacks its maplike methods). */
const highlights = () => CSS.highlights as unknown as Map<string, Highlight>;
/** Zoom steps; 1 = fit the slide to the stage. */
/** Zoom steps as Docs: "fit" fits the slide to the pane, a number is that scale of the slide's real size (96 dpi). */
const ZOOMS = ["fit", "0.5", "0.75", "0.9", "1", "1.25", "1.5", "2"];
const zoomLabel = (z: string) => (z === "fit" ? "Fit" : `${Math.round(Number(z) * 100)}%`);

type Snapshot = { snap: PptxSnapshot; index: number; editPart?: string | null };

const isMod = (e: KeyboardEvent | MouseEvent) => e.ctrlKey || e.metaKey;
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** Start / end of a range as character offsets into `root`'s text. */
function textOffsets(root: HTMLElement, range: Range): [number, number] {
  const at = (node: Node, offset: number) => {
    const r = document.createRange();
    r.setStart(root, 0);
    r.setEnd(node, offset);
    return r.toString().length;
  };
  return [at(range.startContainer, range.startOffset), at(range.endContainer, range.endOffset)];
}

/** Range covering character offsets [start, end) of `root`'s text. */
function rangeAt(root: HTMLElement, [start, end]: [number, number]): Range {
  const range = document.createRange();
  range.selectNodeContents(root);
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let seen = 0;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.textContent?.length ?? 0;
    if (start >= seen && start <= seen + len) range.setStart(n, start - seen);
    if (end >= seen && end <= seen + len) {
      range.setEnd(n, end - seen);
      break;
    }
    seen += len;
  }
  return range;
}

/**
 * In an empty text box the caret can end up on the box itself (the prompt text it was clicked on is gone) or
 * outside it (a click that first selects the shape), so it's drawn in the wrong place; put it in the empty first line.
 */
function caretIntoEmptyLine(textEl: HTMLElement): void {
  const sel = getSelection();
  const p = textEl.querySelector("p");
  if (!sel || !p || textEl.textContent) return;
  // Already inside the line: leave it.
  if (sel.anchorNode && sel.anchorNode !== textEl && textEl.contains(sel.anchorNode)) return;
  sel.collapse(p, 0);
}

/** Copied slide items, shared by every open presentation. */
let clipboard: { clip: Clip; slidePath: string; pastes: number } | null = null;
/** Slides copied from any open deck (Copy slide / Ctrl+C on the thumbnails). */
let slideClipboard: SlideClip | null = null;
const extOf = (name: string) => name.slice(name.lastIndexOf(".") + 1).toLowerCase();

export class PptxView extends FileView {
  /** Stops giving focus back to the document after dialogs close. */
  private stopFocusReturn: (() => void) | null = null;
  allowNoFile = false;
  paperTheme: "light" | "dark" = "dark";

  private doc: PptxDocument | null = null;
  private index = 0;
  private model: SlideModel | null = null;
  private rendered: RenderedShape[] = [];
  private scale = 1;
  /** Rendered shape that gets text formatting, handles and the shape panel. */
  private selected: RenderedShape | null = null;
  /** Selected slide items (direct children of the shape tree: shapes, pictures, groups). */
  private items: Element[] = [];
  private autofitting = false;
  /** Shapes whose text changed since they were last fitted. */
  private needsFit = new Set<Element>();
  private activeText: RenderedShape | null = null;
  private savedRange: Range | null = null;
  private history = new History<Snapshot>();
  private imageUrls = new Map<string, string>();
  private autosave = new Autosave({
    delay: 800,
    readOnly: () => this.settings.readOnly,
    canSave: () => !!this.doc && !!this.file,
    name: () => this.file?.name ?? "",
    before: () => this.flushCommit(),
    pending: () => !!this.pendingCommit,
    write: async () => {
      const file = this.file!;
      const bytes = await this.doc!.save();
      await this.app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    },
  });
  private commitTimer: number | null = null;
  private pendingCommit: RenderedShape | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private lastStageWidth = 0;

  private thumbsEl!: HTMLElement;
  private thumbObserver: ResizeObserver | null = null;
  private fitThumbs: (() => void) | null = null;
  private stageEl!: HTMLElement;
  private slideHostEl!: HTMLElement;
  private notesEl!: HTMLTextAreaElement;
  private notesHighlightEl!: HTMLElement;
  private notesObserver: ResizeObserver | null = null;
  private notesTimer: number | null = null;
  private propsEl!: HTMLElement;
  // Closed until View › Format panel opens it, as Google Slides' Format options.
  private showProps = this.app.loadLocalStorage(PROPS_KEY) === "shown";
  private showMotion = false;
  private showTheme = false;
  private editPart: string | null = null;
  private galleryEl: HTMLElement | null = null;
  private fontSelect: HTMLSelectElement | null = null;
  private sizeSelect: HTMLSelectElement | null = null;
  /** Shape picked in the gallery, waiting to be drawn on the slide. */
  private drawing: ShapeChoice | null = null;
  private lockAspect = false;
  /** Table whose cell has the caret, and the selected cell range (Shift+click extends it). */
  private tableCell: { frame: Element; range: CellRange } | null = null;
  /** The next cell focus comes from a Shift+click that extends the range. */
  private extendingRange = false;
  /** Slide index being dragged in the thumbnail pane. */
  private dragFrom: number | null = null;
  private toolbar: Toolbar | null = null;
  private narrowWatch: ReturnType<typeof watchNarrow> | null = null;
  private closeDrawer: (() => void) | null = null;

  /** Counts onLoadFile calls, so an older load still reading the file gives way. */
  private loads = 0;

  constructor(leaf: WorkspaceLeaf, private settings: DocWorksSettings = { ...DEFAULT_SETTINGS }) {
    super(leaf);
    // Shape shortcuts go through a view scope so they win over app hotkeys (Ctrl+G opens the graph otherwise).
    // They only act while the slide stage has focus; in text they fall through to normal editing.
    this.scope = new Scope(this.app.scope);
    focusRegions(this.scope, this.contentEl, () => [this.contentEl.querySelector<HTMLElement>(".dw-menubar"), this.contentEl.querySelector<HTMLElement>(".dw-toolbar"), this.thumbsEl, this.stageEl, this.propsEl, this.notesEl, this.findEl]);
    const keys: [Modifier[], string, () => void][] = [
      [["Mod"], "c", () => void this.copySelection()],
      [["Mod"], "x", () => void this.cutSelection()],
      [["Mod"], "v", () => this.paste()],
      [["Mod"], "d", () => void this.duplicateSelection()],
      [["Mod"], "a", () => this.setSelection(this.slideItems())],
      [["Mod"], "g", () => this.groupSelection()],
      [["Mod", "Shift"], "g", () => this.ungroupSelection()],
    ];
    this.scope.register(["Mod"], "f", () => {
      this.openFind();
      return false;
    });
    // Google's shortcuts for links, comments and lists. Shift+7 / Shift+8 arrive as "&" / "*" on most layouts.
    const google: [Modifier[], string, () => void][] = [
      [["Mod"], "k", () => void this.editLink()],
      [["Mod", "Alt"], "m", () => this.commentsDialog()],
      [["Mod", "Shift"], "8", () => this.toggleList(false)],
      [["Mod", "Shift"], "*", () => this.toggleList(false)],
      [["Mod", "Shift"], "7", () => this.toggleList(true)],
      [["Mod", "Shift"], "&", () => this.toggleList(true)],
    ];
    for (const [mods, key, run] of google) {
      this.scope.register(mods, key, () => {
        if (this.settings.readOnly) return true;
        run();
        return false;
      });
    }
    registerShortcutsKey(this.scope, () => shortcutsDialog(this.app, [...COMMON_SHORTCUTS, ...SLIDES_SHORTCUTS]));
    // Google's grid view shortcut.
    this.scope.register(["Mod", "Alt"], "1", () => {
      this.setGridView(!this.gridView);
      return false;
    });
    // On the thumbnails, copy / cut / paste act on whole slides.
    const slideKeys: Record<string, () => void> = { c: () => void this.copySlide(false), x: () => void this.copySlide(true), v: () => void this.pasteSlides() };
    for (const [mods, key, run] of keys) {
      this.scope.register(mods, key, () => {
        if (this.settings.readOnly && key !== "c" && key !== "a") return true;
        if (mods.length === 1 && slideKeys[key] && this.thumbsEl && document.activeElement === this.thumbsEl) {
          slideKeys[key]();
          return false;
        }
        if (!this.stageEl || document.activeElement !== this.stageEl) return true;
        run();
        return false;
      });
    }
    // Text boxes and fields keep these keys (Mod+A, the stage's select-all, is registered above).
    shieldEditingKeys(this.scope, EDITING_KEYS.filter(([, key]) => key !== "a"));
    // Undo / redo also when focus is on the page (after a dialog or a toolbar button); the stage and text boxes
    // handle their own keys, and fields keep their own undo.
    const history: [Modifier[], string, () => void][] = [
      [["Mod"], "z", () => this.undo()],
      [["Mod", "Shift"], "z", () => this.redo()],
      [["Mod"], "y", () => this.redo()],
    ];
    for (const [mods, key, run] of history) {
      this.scope.register(mods, key, () => {
        const a = document.activeElement as HTMLElement | null;
        if (a && (a === this.stageEl || a.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(a.tagName))) return true;
        run();
        return false;
      });
    }
  }

  onPaneMenu(menu: Menu, source: string): void {
    super.onPaneMenu(menu, source);
    addFileItems(this, menu, source);
    if (canExportPdf()) menu.addItem((i) => i.setTitle("Export…").setIcon("file-output").setSection("action").onClick(() => void this.exportAs()));
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
    const load = ++this.loads;
    this.stopFocusReturn?.();
    this.stopFocusReturn = returnFocusAfterOverlays(() => this.app.workspace.getActiveViewOfType(PptxView) === this, () => this.stageEl?.focus({ preventScroll: true }));
    this.contentEl.empty();
    this.contentEl.addClass("pptx-editor-root");
    this.contentEl.toggleClass("dw-read-only", this.settings.readOnly);
    if (this.settings.readOnly) this.showProps = false;
    const status = this.contentEl.createDiv({ cls: "dw-loading", text: "Loading presentation..." });

    try {
      const doc = await PptxDocument.load(await this.app.vault.readBinary(file));
      // Opened again meanwhile (a setting change while loading): only the newest load draws.
      if (load !== this.loads) return;
      this.doc = doc;
      await this.loadMedia();
      if (load !== this.loads) return;
    } catch (err) {
      status.setText(`Could not open this .pptx file: ${(err as Error).message}`);
      console.error(err);
      this.doc = null;
      return;
    }
    status.remove();

    this.buildToolbar(this.contentEl.createDiv({ cls: "dw-toolbar pptx-toolbar" }));
    const body = this.contentEl.createDiv({ cls: "pptx-body" });
    // Named through a hidden label rather than aria-label, which Obsidian would show as a hover tooltip.
    const thumbsName = body.createDiv({ cls: "dw-sr-only", text: "Slides", attr: { id: `pptx-thumbs-name-${Math.random().toString(36).slice(2)}` } });
    this.thumbsEl = body.createDiv({ cls: "pptx-thumbs", attr: { role: "listbox", "aria-labelledby": thumbsName.id } });
    this.thumbsEl.tabIndex = 0;
    this.thumbsEl.id = `dw-slides-${crypto.randomUUID()}`;
    // Delete / Backspace on the thumbnails deletes the current slide (a deck keeps at least one).
    this.thumbsEl.addEventListener("keydown", (e) => {
      if (e.target === this.thumbsEl && this.doc && ["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) {
        e.preventDefault();
        const last = this.doc.slides().length - 1;
        // In the grid, up / down move a whole row.
        const perRow = this.gridView ? Math.max(1, Math.floor(this.thumbsEl.clientWidth / (GRID_THUMB_WIDTH + 40))) : 1;
        const step = { ArrowDown: perRow, ArrowUp: -perRow, ArrowRight: 1, ArrowLeft: -1 }[e.key as "ArrowDown"] ?? 0;
        this.goTo(e.key === "Home" ? 0 : e.key === "End" ? last : Math.max(0, Math.min(last, this.index + step)));
        return;
      }
      if (e.target === this.thumbsEl && this.gridView && (e.key === "Enter" || e.key === "Escape")) {
        e.preventDefault();
        this.setGridView(false);
        return;
      }
      if ((e.key !== "Delete" && e.key !== "Backspace") || e.target !== this.thumbsEl || !this.doc || this.settings.readOnly) return;
      e.preventDefault();
      const count = this.doc.slides().length;
      if (count < 2) return;
      const i = this.index;
      this.mutate((d) => {
        d.deleteSlide(i);
        this.index = Math.min(i, count - 2);
      });
    });
    // Thumbnails are drawn THUMB_WIDTH wide; shrink them to the room the panel leaves (scroll bar, narrow windows).
    // Measured from the panel (its scroll bar space is always reserved), not the thumbnails: scaling them must not
    // change what's measured, or the scroll bar coming and going makes them flicker.
    const fitThumbs = (this.fitThumbs = () => {
      const width = this.thumbsEl.clientWidth;
      if (!width || this.gridView) return;
      const room = width - THUMB_CHROME;
      this.thumbsEl.setCssProps({ "--pptx-thumb-zoom": String(Math.min(1, Math.max(0.3, room / THUMB_WIDTH))) });
    });
    this.thumbObserver?.disconnect();
    this.thumbObserver = new ResizeObserver(fitThumbs);
    this.thumbObserver.observe(this.thumbsEl);
    this.stageEl = body.createDiv({ cls: "pptx-stage" });
    this.stageEl.tabIndex = 0;
    // Ctrl / Cmd + wheel zooms the slide.
    this.stageEl.addEventListener(
      "wheel",
      (e) => {
        if (!isMod(e)) return;
        e.preventDefault();
        this.setZoom(String((this.scale * EMU_PER_PX) * (e.deltaY < 0 ? 1.1 : 1 / 1.1)));
      },
      { passive: false }
    );
    this.slideHostEl = this.stageEl.createDiv({ cls: "pptx-slide-host" });
    this.propsEl = body.createDiv({ cls: "pptx-props" });
    escapeToEditor(this.propsEl, () => this.stageEl.focus({ preventScroll: true }));
    this.propsEl.hide();
    this.slideHostEl.addEventListener("pointerdown", (e) => this.drawing && this.startDraw(e), { capture: true });
    const notesBox = this.stageEl.createDiv({ cls: "pptx-notes-box" });
    this.notesHighlightEl = notesBox.createDiv({ cls: "pptx-notes-highlight", attr: { "aria-hidden": "true" } });
    this.notesEl = notesBox.createEl("textarea", { cls: "pptx-notes", attr: { "aria-label": "Speaker notes", placeholder: this.settings.readOnly ? "" : "Click to add speaker notes" } });
    this.notesEl.readOnly = this.settings.readOnly;
    this.notesEl.addEventListener("input", () => {
      this.clearNotesHighlight();
      if (this.notesTimer) window.clearTimeout(this.notesTimer);
      this.notesTimer = window.setTimeout(() => this.flushNotes(), 400);
    });
    this.notesEl.addEventListener("blur", () => this.flushNotes());
    const syncNotesHighlight = () => {
      const style = getComputedStyle(this.notesEl);
      for (const property of ["font-family", "font-size", "font-weight", "font-style", "line-height", "letter-spacing", "padding-top", "padding-right", "padding-bottom", "padding-left", "border-width", "word-break", "overflow-wrap", "tab-size", "text-align"]) {
        this.notesHighlightEl.style.setProperty(property, style.getPropertyValue(property));
      }
      this.notesHighlightEl.style.width = `${this.notesEl.offsetWidth}px`;
      this.notesHighlightEl.scrollTop = this.notesEl.scrollTop;
      this.notesHighlightEl.scrollLeft = this.notesEl.scrollLeft;
    };
    this.notesEl.addEventListener("scroll", syncNotesHighlight);
    this.notesObserver?.disconnect();
    this.notesObserver = new ResizeObserver(syncNotesHighlight);
    this.notesObserver.observe(this.notesEl);

    this.stageEl.addEventListener("keydown", this.onStageKey);
    this.stageEl.addEventListener("pointerdown", (e) => {
      const t = e.target as HTMLElement;
      if (this.settings.readOnly || this.drawing || e.button !== 0 || t.closest(".pptx-shape:not(.is-inherited)") || t.closest("textarea, .pptx-props")) return;
      if (t === this.stageEl || this.slideHostEl.contains(t)) this.startMarquee(e);
    });
    this.stageEl.addEventListener("contextmenu", (e) => !this.settings.readOnly && this.contextMenu(e));
    this.stageEl.addEventListener("paste", (e) => !this.settings.readOnly && this.onPaste(e));
    this.stageEl.addEventListener("dragover", (e) => e.preventDefault());
    this.stageEl.addEventListener("drop", (e) => {
      const image = Array.from(e.dataTransfer?.files ?? []).find((f) => f.type.startsWith("image/"));
      if (!image || this.settings.readOnly) return;
      e.preventDefault();
      void this.insertImageFile(image);
    });
    document.addEventListener("selectionchange", this.onSelectionChange);
    window.addEventListener("blur", this.onWindowBlur);
    this.resizeObserver = new ResizeObserver(() => {
      // Re-render only on real width changes, never while the user is typing.
      if (!this.pendingCommit && !this.activeTextFocused()) this.refit();
    });
    this.resizeObserver.observe(this.stageEl);
    this.narrowWatch = watchNarrow(this.contentEl, (narrow) => {
      if (!narrow) this.closeDrawer?.();
      this.toolbar?.fit();
      this.fitThumbs?.();
    });

    this.index = 0;
    this.setPaperTheme(paperFor(this.settings, "dark"));
    this.renderAll();
    void this.loadFonts();
  }

  /** Removes the deck's embedded fonts from the page. */
  private unloadFonts: (() => void) | null = null;

  /** Load the fonts the deck carries, then redraw with them (after the first render, so opening isn't held up). */
  private async loadFonts(): Promise<void> {
    const doc = this.doc;
    if (!doc) return;
    const { loaded, unload } = await loadEmbeddedFonts(await doc.embeddedFonts(), eotToTtf);
    if (this.doc !== doc) return unload(); // another file was opened meanwhile
    this.unloadFonts = unload;
    if (loaded) this.renderAll();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    this.unloadFonts?.();
    this.unloadFonts = null;
    this.thumbObserver?.disconnect();
    this.thumbObserver = null;
    this.notesObserver?.disconnect();
    this.notesObserver = null;
    document.removeEventListener("selectionchange", this.onSelectionChange);
    window.removeEventListener("blur", this.onWindowBlur);
    this.resizeObserver?.disconnect();
    this.narrowWatch?.();
    this.closeDrawer?.();
    // onLoadFile empties contentEl; without this the next file's Ctrl+F and shape gallery see stale elements.
    if (this.findEl) this.closeFind();
    this.closeGallery();
    await this.autosave.close();
    this.stopFocusReturn?.();
    this.stopFocusReturn = null;
    for (const url of this.imageUrls.values()) URL.revokeObjectURL(url);
    this.imageUrls.clear();
    this.doc = null;
    this.showMotion = false;
    this.showTheme = false; this.editPart = null;
    this.rendered = [];
    this.history = new History<Snapshot>();
  }

  /** Force an immediate save, e.g. before Obsidian quits. */
  async requestSave(): Promise<void> {
    await this.autosave.flush();
  }

  private onWindowBlur = () => void this.autosave.flush();

  private async loadMedia(): Promise<void> {
    for (const path of this.doc!.mediaPaths()) {
      const mime = IMAGE_MIME[extOf(path)] ?? MEDIA_MIME[extOf(path)];
      if (!mime || this.imageUrls.has(path)) continue;
      const bytes = await this.doc!.media(path);
      this.imageUrls.set(path, URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime })));
    }
  }

  private imageUrl = (path: string) => this.imageUrls.get(path) ?? this.linkedMediaUrl(path);

  /** Vault resource URL (app://, never the network) for media the deck links to rather than embeds. */
  private linkedMediaUrl(target: string): string | null {
    if (!this.file || this.doc?.mediaPaths().includes(target)) return null;
    for (const path of linkedMediaCandidates(target, this.file.path)) {
      const file = this.app.vault.getAbstractFileByPath(path);
      if (file instanceof TFile) return this.app.vault.getResourcePath(file);
    }
    return null;
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  private renderAll(): void {
    this.renderThumbs();
    this.renderCurrent();
  }

  /** Paint format: the copied format, and whether it stays on after painting once (double-click). */
  private painter: { format: ShapeFormat; sticky: boolean } | null = null;
  private brushBtn: HTMLElement | null = null;

  private startPainting(sticky: boolean): void {
    // A second click turns it off (the dblclick's first click already turned it on).
    if (this.painter && !sticky) return this.stopPainting();
    const source = this.selected?.shape ?? this.rendered.find((r) => r.shape.el === this.items[0])?.shape;
    if (!source) {
      new Notice("Select a shape to copy its format.");
      return;
    }
    this.painter = { format: copyFormat(source), sticky };
    this.brushBtn?.addClass("is-active");
    this.stageEl.addClass("is-painting");
  }

  private stopPainting(): void {
    this.painter = null;
    this.brushBtn?.removeClass("is-active");
    this.stageEl?.removeClass("is-painting");
  }

  /** Paint the copied format onto a shape (the whole top-level item for shapes in groups). */
  private paintOnto(shape: ShapeModel): void {
    const painter = this.painter;
    if (!painter) return;
    const el = shape.el;
    this.mutate((doc) => {
      const part = this.editPart ?? doc.slides()[this.index].path;
      const model = parseSlide(doc, part).shapes.find((s) => s.el === el);
      if (model) applyFormat(model, painter.format);
      doc.touch(part);
    }, "slide");
    if (!painter.sticky) this.stopPainting();
  }

  /** Google's grid view: the slides fill the window as a grid, for reordering; double-click opens one. */
  private gridView = false;

  private setGridView(on: boolean): void {
    this.gridView = on;
    this.contentEl.toggleClass("pptx-grid-view", on);
    if (!this.thumbsEl) return;
    this.thumbsEl.setCssProps({ "--pptx-thumb-zoom": on ? String(GRID_THUMB_WIDTH / THUMB_WIDTH) : "" });
    if (on) this.thumbsEl.focus();
    else {
      this.fitThumbs?.();
      this.stageEl?.focus({ preventScroll: true });
      this.renderCurrent();
    }
  }

  private renderThumbs(): void {
    if (!this.doc) return;
    this.thumbsEl.empty();
    const sections = this.doc.sections();
    const header = (parent: HTMLElement, k: number) => {
      const h = parent.createDiv({ cls: "pptx-section", text: sections[k].name || "Untitled section" });
      h.addEventListener("click", (e) => e.stopPropagation());
      if (this.settings.readOnly) return;
      h.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.sectionMenu(k).showAtMouseEvent(e);
      });
      h.addEventListener("dblclick", () => void this.renameSection(k));
    };
    const slides = this.doc.slides();
    slides.forEach((slide, i) => {
      const item = this.thumbsEl.createDiv({ cls: "pptx-thumb", attr: { role: "option", id: `${this.thumbsEl.id}-slide-${i}`, "aria-label": `Slide ${i + 1}${this.doc!.isHidden(i) ? ", hidden" : ""}`, "aria-selected": String(i === this.index) } });
      // Section headers sit inside the first slide's item, so thumbsEl.children[i] stays slide i.
      sections.forEach((sec, k) => sec.start === i && header(item, k));
      item.toggleClass("is-active", i === this.index);
      item.toggleClass("is-hidden", this.doc!.isHidden(i));
      item.createDiv({ cls: "pptx-thumb-num", text: String(i + 1) });
      const host = item.createDiv({ cls: "pptx-thumb-slide" });
      renderSlide(host, parseSlide(this.doc!, slide.path), this.doc!.slideSize, {
        width: THUMB_WIDTH,
        editable: false,
        imageUrl: this.imageUrl,
        slideNumber: i + 1,
      });
      item.onclick = () => {
        this.goTo(i);
        // Copy / paste on the thumbnails act on slides.
        this.thumbsEl.focus();
      };
      item.addEventListener("dblclick", () => {
        if (!this.gridView) return;
        this.setGridView(false);
        this.goTo(i);
      });
      if (this.settings.readOnly) return;
      item.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        this.goTo(i);
        this.slideMenu(i).showAtMouseEvent(e);
      });
      this.wireThumbDrag(item, i);
    });
    // Empty sections at the end have no slide to sit on.
    sections.forEach((sec, k) => sec.start >= slides.length && header(this.thumbsEl, k));
    this.thumbsEl.setAttribute("aria-activedescendant", `${this.thumbsEl.id}-slide-${this.index}`);
    if (this.settings.readOnly) return;
    const add = this.thumbsEl.createEl("button", { cls: "pptx-thumb-add", attr: { "aria-label": "New slide" } });
    setIcon(add, "plus");
    add.addEventListener("click", () => this.mutate((d) => (this.index = d.addSlide(d.slides().length - 1))));
  }

  /** Drag a thumbnail above or below another to move the slide. */
  private wireThumbDrag(item: HTMLElement, i: number): void {
    item.draggable = true;
    const clear = () => this.thumbsEl.querySelectorAll(".drop-before, .drop-after").forEach((t) => t.removeClass("drop-before", "drop-after"));
    // In the grid, slides sit side by side: before / after is left / right.
    const below = (e: DragEvent) => {
      const r = item.getBoundingClientRect();
      return this.gridView ? e.clientX > r.left + r.width / 2 : e.clientY > r.top + r.height / 2;
    };
    item.addEventListener("dragstart", (e) => {
      this.dragFrom = i;
      e.dataTransfer?.setData("text/plain", String(i + 1));
      if (e.dataTransfer) e.dataTransfer.effectAllowed = "move";
    });
    item.addEventListener("dragend", () => {
      this.dragFrom = null;
      clear();
    });
    item.addEventListener("dragover", (e) => {
      if (this.dragFrom === null) return;
      e.preventDefault();
      clear();
      item.addClass(below(e) ? "drop-after" : "drop-before");
    });
    item.addEventListener("drop", (e) => {
      const from = this.dragFrom;
      if (from === null) return;
      e.preventDefault();
      clear();
      const to = dropIndex(from, i, below(e));
      if (to !== from) this.mutate((d) => d.moveSlide(from, (this.index = to)));
    });
  }

  private slideMenu(i: number, menu = new Menu()): Menu {
    const item = (title: string, icon: string, run: () => void) => menu.addItem((m) => m.setTitle(title).setIcon(icon).onClick(run));
    const count = this.doc?.slides().length ?? 0;
    item("New slide", "plus", () => this.mutate((d) => (this.index = d.addSlide(i))));
    item("Duplicate slide", "copy", () => this.mutate((d) => (this.index = d.duplicateSlide(i))));
    item("Copy slide", "clipboard-copy", () => void this.copySlide(false));
    item("Cut slide", "scissors", () => void this.copySlide(true));
    if (slideClipboard) item("Paste slide", "clipboard-paste", () => void this.pasteSlides());
    item("Delete slide", "trash", () => this.mutate((d) => d.deleteSlide(i)));
    item(this.doc?.isHidden(i) ? "Show slide" : "Hide slide", "eye-off", () => this.mutate((d) => d.setHidden(i, !d.isHidden(i))));
    menu.addSeparator();
    if (i > 0) item("Move up", "arrow-up", () => this.mutate((d) => d.moveSlide(i, (this.index = i - 1))));
    if (i < count - 1) item("Move down", "arrow-down", () => this.mutate((d) => d.moveSlide(i, (this.index = i + 1))));
    item("Add section here", "list-plus", () => void this.addSection(i));
    return menu;
  }

  /** The pending copy of the current slide (for paste to wait on). */
  copying: Promise<void> | null = null;

  /** Copy (or cut) the current slide, for pasting into this or another open deck. */
  private copySlide(cut: boolean): Promise<void> {
    const doc = this.doc;
    if (!doc) return Promise.resolve();
    const index = this.index;
    this.copying = (async () => {
      slideClipboard = await doc.copySlides([index]);
      if (cut && doc === this.doc && doc.slides().length > 1) this.mutate((d) => d.deleteSlide(index));
      new Notice(cut ? "Slide cut" : "Slide copied");
    })();
    return this.copying;
  }

  /** Paste copied slides after the current one; the destination's layouts and theme apply. */
  private async pasteSlides(): Promise<void> {
    const clip = slideClipboard;
    if (!this.doc || !clip) return;
    this.mutate((d) => (this.index = d.pasteSlides(clip, this.index)));
    await this.loadMedia();
    this.renderAll();
  }

  private sectionMenu(k: number): Menu {
    const menu = new Menu();
    menu.addItem((m) => m.setTitle("Rename section").setIcon("pencil").onClick(() => void this.renameSection(k)));
    menu.addItem((m) => m.setTitle("Remove section").setIcon("trash").onClick(() => this.mutate((d) => d.removeSection(k))));
    return menu;
  }

  private async addSection(start: number): Promise<void> {
    const name = await promptText(this.app, "Section name", "New section");
    if (name?.trim()) this.mutate((d) => d.addSection(start, name.trim()));
  }

  private async renameSection(k: number): Promise<void> {
    const name = await promptText(this.app, "Section name", this.doc?.sections()[k]?.name ?? "");
    if (name?.trim()) this.mutate((d) => d.renameSection(k, name.trim()));
  }

  private slideControl(): SlideControl {
    const doc = this.doc!;
    const index = this.index;
    const layout = doc.layoutOf(doc.slides()[index].path) ?? "";
    return {
      layouts: doc.layouts(doc.masterOf(layout) ?? undefined),
      layout,
      setLayout: (p) => this.mutate((d) => d.setLayout(index, p), "slide"),
      background: this.model?.ownBackground ?? null,
      setBackground: (f) => this.mutate((d) => d.setBackground(index, f), "slide"),
      applyBackgroundToAll: () => {
        const f = this.model?.ownBackground ?? null;
        this.mutate((d) => d.slides().forEach((_, i) => d.setBackground(i, f)));
      },
      size: doc.slideSize,
      setSize: (cx, cy) => this.mutate((d) => d.setSlideSize(cx, cy)),
    };
  }

  /** Open the motion editor while keeping the object selection for animation targets. */
  private openMotion(): void {
    this.flushCommit();
    this.showMotion = true;
    this.showProps = true;
    this.openPropsDrawer();
    this.setSelection(this.items, this.selected?.shape.el ?? null);
  }

  private renderMotionPanel(): void {
    const xml = this.currentSlideXml();
    if (!xml) return;
    const selected = this.items.length === 1 ? this.items[0] : null;
    const selectedId = selected ? shapeId(selected) : null;
    const selectedName = selected?.getElementsByTagNameNS(P_NS, "cNvPr")[0]?.getAttribute("name") ?? "Object";
    buildMotionPanel(this.propsEl, {
      transition: readTransition(xml), entries: animationEntries(xml), selectedId, selectedName,
      setTransition: (patch) => this.mutate(() => writeTransition(xml, patch), "slide"),
      applyTransitionToAll: () => {
        const t = readTransition(xml);
        const patch = t && t.type !== "cut" ? { type: t.type, dir: t.dir, dur: t.dur } : { type: "cut" };
        this.mutate((d) => d.slides().forEach((s) => writeTransition(d.xml(s.path)!, patch)));
      },
      addAnimation: (settings) => { if (selectedId) this.mutate(() => addAnimation(xml, selectedId, settings), "slide"); },
      updateAnimation: (entry, patch) => this.mutate(() => updateAnimation(xml, entry, patch), "slide"),
      updatePart: (part, patch) => this.mutate(() => updateAnimationPart(xml, part, patch), "slide"),
      removeAnimation: (entry) => this.mutate(() => removeAnimation(xml, entry), "slide"),
      moveAnimation: (entry, direction) => this.mutate(() => moveAnimation(xml, entry, direction), "slide"),
      preview: () => this.startPresentation(false, true),
      back: () => {
        this.showMotion = false;
        this.setSelection(this.items, this.selected?.shape.el ?? null);
      },
    });
  }

  private refreshThumb(i: number): void {
    const host = this.thumbsEl.children[i]?.querySelector<HTMLElement>(".pptx-thumb-slide");
    const slide = this.doc?.slides()[i];
    if (!host || !slide || !this.doc) return;
    renderSlide(host, parseSlide(this.doc, slide.path), this.doc.slideSize, { width: THUMB_WIDTH, editable: false, imageUrl: this.imageUrl, slideNumber: i + 1 });
  }

  /** Re-render if the stage width really changed since the last render. */
  private refit(): void {
    if (Math.abs(this.stageEl.clientWidth - this.lastStageWidth) > 8) this.renderCurrent();
  }

  private renderCurrent(): void {
    if (!this.doc) return;
    const slides = this.doc.slides();
    this.index = Math.max(0, Math.min(this.index, slides.length - 1));
    if (this.editPart && !this.doc.xml(this.editPart)) this.editPart = null;
    this.model = parseSlide(this.doc, this.editPart ?? slides[this.index].path);

    this.lastStageWidth = this.stageEl.clientWidth;
    const width = this.zoom === "fit" ? Math.max(320, Math.min((this.lastStageWidth || 1008) - 48, 1400)) : (this.doc.slideSize.cx / EMU_PER_PX) * Number(this.zoom);
    this.scale = width / this.doc.slideSize.cx;
    // The notes box lines up with the slide.
    this.stageEl.setCssProps({ "--pptx-slide-width": `${width}px` });
    this.rendered = renderSlide(this.slideHostEl, this.model, this.doc.slideSize, { width, editable: !this.settings.readOnly, imageUrl: this.imageUrl, slideNumber: this.index + 1 });
    if (!this.settings.readOnly) this.rendered.forEach((rs) => this.wireShape(rs));

    this.notesEl.value = this.doc.notesText(this.index);
    this.clearNotesHighlight();
    const notesMatch = this.findMatches[this.findIndex];
    if (this.findEl && notesMatch?.kind === "notes" && notesMatch.slide === this.index && notesMatch.text === this.notesEl.value) this.highlightNotes(notesMatch.start, notesMatch.end);
    // Keep the selection across redraws of the same slide (resizes, edits).
    const xml = this.currentSlideXml();
    const keep = this.items.filter((i) => xml?.contains(i));
    const primary = this.selected?.shape.el ?? null;
    this.items = [];
    this.selected = null;
    this.setSelection(keep, primary);
    this.activeText = null;
    this.savedRange = null;
    Array.from(this.thumbsEl.querySelectorAll('[role="option"]')).forEach((c, i) => {
      c.toggleClass("is-active", i === this.index);
      c.setAttribute("aria-selected", String(i === this.index));
    });
    this.thumbsEl.setAttribute("aria-activedescendant", `${this.thumbsEl.id}-slide-${this.index}`);
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

    // Ctrl / Cmd + click follows a shape's own link (text links are followed by the text's handler).
    if (shape.link) el.addEventListener("click", (e) => {
      if (isMod(e) && shape.link) this.followLink(shape.link);
    });

    el.addEventListener("pointerdown", (e) => {
      if (e.button !== 0 || this.drawing) return;
      if (this.painter) {
        e.preventDefault();
        e.stopPropagation();
        this.paintOnto(shape);
        return;
      }
      const item = topLevel(shape.el);
      // Shift+click in another cell of the same table extends the cell range.
      if (shape.kind === "cell" && e.shiftKey && this.tableCell?.frame === item && shape.cell) {
        e.stopPropagation();
        this.extendingRange = true;
        const [r, c] = this.tableCell.range;
        this.tableCell = { frame: item, range: [r, c, shape.cell.r, shape.cell.c] };
        this.setSelection([item], shape.el);
        return;
      }
      if (e.shiftKey || isMod(e)) {
        e.preventDefault();
        this.setSelection(this.items.includes(item) ? this.items.filter((i) => i !== item) : [...this.items, item]);
        return;
      }
      if (!this.items.includes(item)) this.setSelection([item], shape.el);
      // Pressing in text selects text; a shape without text (not a placeholder) moves by its body, as in Google
      // Slides; a plain click still puts the caret in it.
      const emptyShape = !shape.placeholder && !!textEl?.classList.contains("is-empty");
      if (!emptyShape && (textEl?.contains(e.target as Node) || (e.target as HTMLElement).closest(".pptx-cell-text"))) return;
      if (!emptyShape) this.stageEl.focus({ preventScroll: true });
      this.startMoveItems(e);
    });

    if (!shape.grouped) {
      for (const mode of ["move", "resize"] as const) {
        const handle = el.createDiv({ cls: `pptx-handle pptx-handle-${mode}` });
        handle.setAttribute("title", mode === "move" ? "Drag to move" : "Drag to resize");
        handle.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
          this.select(rs);
          this.stageEl.focus({ preventScroll: true });
          if (mode === "move") this.startMoveItems(e);
          else this.startDrag(e, rs, mode);
        });
      }
      if (shape.el.localName === "sp" || shape.el.localName === "pic") {
        const rot = el.createDiv({ cls: "pptx-handle pptx-handle-rotate", attr: { "aria-label": "Drag to rotate (Shift snaps to 15°)" } });
        rot.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
          this.select(rs);
          this.startRotate(e, rs);
        });
      }
    }

    // Text fills the shape, so its edges are the grab area for moving (as in PowerPoint).
    if (textEl && shape.kind !== "cell") {
      for (const side of ["t", "r", "b", "l"]) {
        el.createDiv({ cls: `pptx-edge pptx-edge-${side}` }).addEventListener("pointerdown", (e) => {
          if (e.button !== 0) return;
          e.stopPropagation();
          const item = topLevel(shape.el);
          if (e.shiftKey || isMod(e)) {
            this.setSelection(this.items.includes(item) ? this.items.filter((i) => i !== item) : [...this.items, item]);
            return;
          }
          if (!this.items.includes(item)) this.setSelection([item], shape.el);
          this.stageEl.focus({ preventScroll: true });
          this.startMoveItems(e);
        });
      }
    }

    if (textEl) {
      textEl.addEventListener("click", (e) => {
        caretIntoEmptyLine(textEl);
        // Ctrl / Cmd + click follows a link.
        const link = (e.target as HTMLElement).closest<HTMLElement>("[data-link]")?.dataset.link;
        if (link && isMod(e)) {
          e.stopPropagation();
          this.followLink(link);
        }
      });
      textEl.addEventListener("input", () => {
        textEl.toggleClass("is-empty", !textEl.textContent);
        refreshMarks(textEl);
        this.scheduleCommit(rs);
      });
      textEl.addEventListener("focus", () => {
        this.activeText = rs;
        window.setTimeout(() => caretIntoEmptyLine(textEl));
        if (shape.kind === "cell" && shape.cell) {
          const frame = topLevel(shape.el);
          if (!this.extendingRange) this.tableCell = { frame, range: [shape.cell.r, shape.cell.c, shape.cell.r, shape.cell.c] };
          this.extendingRange = false;
          this.setSelection([frame], shape.el);
          return;
        }
        if (this.items.includes(topLevel(shape.el))) this.selected = rs;
        else this.select(rs);
      });
      textEl.addEventListener("blur", () => {
        this.flushCommit();
        // Refit only text that was edited: opening and clicking around must not change the file.
        if (this.needsFit.delete(shape.el)) this.autofit(rs);
        // The stage may have resized while typing (e.g. the format panel opened).
        if (!rs.textEl?.contains(document.activeElement)) this.refit();
      });
      textEl.addEventListener("keydown", (e) => this.onTextKey(e, rs));
    }
  }

  /** Select one rendered shape's slide item (null clears the selection). */
  private select(rs: RenderedShape | null): void {
    this.setSelection(rs ? [topLevel(rs.shape.el)] : [], rs?.shape.el ?? null);
  }

  /** Select the rendered shape for `el` again after a redraw. */
  /** The selected cell range of table `frame`, if it is still inside the grid. */
  private cellRange(frame: Element, table: SlideModel["shapes"][number]["table"]): CellRange | null {
    if (!table || this.tableCell?.frame !== frame) return null;
    const [r1, c1, r2, c2] = this.tableCell.range;
    const ok = (r: number, c: number) => r >= 0 && c >= 0 && r < table.rows.length && c < table.cols.length;
    return ok(r1, c1) && ok(r2, c2) ? [Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2)] : null;
  }

  /** Mark the cells of a multi-cell range. */
  private highlightCells(): void {
    const t = this.tableCell;
    const [r1, c1, r2, c2] = t ? this.cellRange(t.frame, this.rendered.find((r) => r.shape.el === t.frame)?.shape.table ?? null) ?? [0, 0, -1, -1] : [0, 0, -1, -1];
    const many = r1 !== r2 || c1 !== c2;
    for (const r of this.rendered) {
      const cell = r.shape.cell;
      if (!cell) continue;
      const inside = many && t && topLevel(r.shape.el) === t.frame && cell.r >= r1 && cell.r <= r2 && cell.c >= c1 && cell.c <= c2;
      r.el.toggleClass("is-cell-selected", !!inside);
    }
  }

  /** Tab / Shift+Tab: move to the next or previous cell; Tab in the last cell adds a row. */
  private nextCell(rs: RenderedShape, dir: number): void {
    const frame = topLevel(rs.shape.el);
    const cells = this.rendered.filter((r) => r.shape.kind === "cell" && r.textEl && topLevel(r.shape.el) === frame);
    const target = cells[cells.indexOf(rs) + dir];
    const focus = (r: RenderedShape | undefined) => {
      if (!r?.textEl) return;
      r.textEl.focus();
      const range = document.createRange();
      range.selectNodeContents(r.textEl);
      window.getSelection()?.removeAllRanges();
      window.getSelection()?.addRange(range);
    };
    if (target) return focus(target);
    if (dir < 0) return;
    const rows = rs.shape.cell ? this.rendered.find((r) => r.shape.el === frame)?.shape.table?.rows.length ?? 0 : 0;
    this.editShape(frame, (f) => insertRow(f.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "tbl")[0], rows - 1, true));
    focus(this.rendered.find((r) => r.shape.kind === "cell" && r.textEl && topLevel(r.shape.el) === frame && r.shape.cell!.r === rows && r.shape.cell!.c === 0));
  }

  private async insertTable(): Promise<void> {
    const answer = await promptText(this.app, "Insert table: rows × columns", "3 x 3");
    const m = answer && /(\d+)\s*[x×*, ]\s*(\d+)/.exec(answer);
    if (!m || !this.doc) return;
    const rows = Math.min(50, Math.max(1, Number(m[1])));
    const cols = Math.min(30, Math.max(1, Number(m[2])));
    const { cx, cy } = this.doc.slideSize;
    const w = cx * 0.8;
    const h = Math.min(cy * 0.7, rows * 370840);
    let frame: Element | null = null;
    const style = TABLE_STYLES[1].id;
    this.mutate((d) => {
      d.ensureTableStyle(style);
      frame = addTable(this.currentSlideXml()!, rows, cols, { x: (cx - w) / 2, y: (cy - h) / 2, w, h }, style);
    }, "slide");
    this.reselect(frame);
  }

  private reselect(el: Element | null): void {
    if (el) this.setSelection([topLevel(el)], el);
  }

  /**
   * Set the selected items. A single plain shape gets handles and the shape
   * panel; groups and multi-selections get a selection box and the arrange panel.
   */
  private setSelection(items: Element[], primaryEl: Element | null = null): void {
    // Nothing on a read-only slide can be selected (moved, deleted, cut).
    if (this.settings.readOnly) [items, primaryEl] = [[], null];
    this.applySelection(items, primaryEl);
    this.toolbar?.sync();
  }

  private applySelection(items: Element[], primaryEl: Element | null): void {
    for (const r of this.rendered) {
      r.el.removeClass("is-selected", "is-in-selection");
      r.el.querySelectorAll(".pptx-handle-adjust").forEach((h) => h.remove());
    }
    this.slideHostEl.querySelectorAll(".pptx-selbox").forEach((b) => b.remove());
    this.items = items;
    if (this.tableCell && !items.includes(this.tableCell.frame)) this.tableCell = null;
    this.highlightCells();
    const single = items.length === 1 && items[0].localName !== "grpSp" ? items[0] : null;
    this.selected = this.rendered.find((r) => r.shape.el === (primaryEl ?? single)) ?? null;
    for (const r of this.rendered) if (items.includes(topLevel(r.shape.el))) r.el.addClass("is-in-selection");

    if (this.showMotion) {
      if (this.selected && !this.selected.shape.inherited) this.selected.el.addClass("is-selected");
      if (items.length && !single) this.drawSelectionBox();
      this.renderMotionPanel();
      this.showPropsEl();
      return;
    }

    if (!items.length && this.showTheme) { this.renderThemePanel(); return; }
    if (!items.length) {
      if (this.doc && this.model) buildSlidePanel(this.propsEl, this.slideControl());
      this.propsEl.toggle((this.showProps || this.propsEl.hasClass("is-drawer")) && !!this.model);
      return;
    }
    const rs = single ? this.rendered.find((r) => r.shape.el === single) : null;
    if (rs && !rs.shape.inherited) {
      this.selected = rs;
      rs.el.addClass("is-selected");
      this.addAdjustHandles(rs);
      const ctl: PanelControl = {
        lockAspect: this.lockAspect,
        edit: (fn: (el: Element) => void) => this.editShape(rs.shape.el, fn, true),
        arrange: this.arrangeControl(),
        cells: this.cellRange(rs.shape.el, rs.shape.table),
        tableStyles: TABLE_STYLES,
        setTableStyle: (id: string) =>
          this.editShape(rs.shape.el, (frame) => {
            this.doc?.ensureTableStyle(id);
            setTableStyleId(frame.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "tbl")[0], id);
          }),
      };
      Object.defineProperty(ctl, "lockAspect", { get: () => this.lockAspect, set: (v: boolean) => (this.lockAspect = v) });
      buildShapePanel(this.propsEl, rs.shape, ctl);
      if (rs.shape.media) {
        for (const [text,key] of [["Play automatically","autoPlay"],["Loop","loop"]] as const) {
          const label=this.propsEl.createEl("label",{cls:"pptx-props-row"});label.createSpan({text});
          const field=label.createEl("input",{type:"checkbox"});field.checked=rs.shape.media[key];
          field.addEventListener("change",()=>this.mutate(()=>setMediaPlayback(this.currentSlideXml()!,rs.shape.el,{[key]:field.checked})));
        }
      }
      if (rs.shape.chart) {
        const button = this.propsEl.createEl("button", { text: "Edit chart" });
        button.addEventListener("click", () => this.editChart(rs.shape.chart!.path));
      }
      if (rs.shape.diagram) {
        const button = this.propsEl.createEl("button", { text: "Edit diagram text" });
        button.addEventListener("click", () => this.editSmartArt(rs.shape.el));
      }
    } else {
      this.drawSelectionBox();
      buildArrangePanel(this.propsEl, this.arrangeControl());
    }
    this.showPropsEl();
  }

  /** Dashed box around a group or multi-selection; a single group can be resized from its corner. */
  private drawSelectionBox(): void {
    const slide = this.slideHostEl.querySelector<HTMLElement>(".pptx-slide");
    if (!slide || !this.items.length) return;
    const b = this.union(this.items.map((i) => this.itemBounds(i)));
    const box = slide.createDiv({ cls: "pptx-selbox" });
    Object.assign(box.style, { left: `${b.x * this.scale}px`, top: `${b.y * this.scale}px`, width: `${b.w * this.scale}px`, height: `${b.h * this.scale}px` });
    if (this.items.length === 1) {
      const group = this.items[0];
      box.createDiv({ cls: "pptx-handle pptx-handle-resize" }).addEventListener("pointerdown", (e) => {
        e.stopPropagation();
        this.startResizeItem(e, group, box);
      });
    }
  }

  private union(rects: Rect[]): Rect {
    const x = Math.min(...rects.map((r) => r.x));
    const y = Math.min(...rects.map((r) => r.y));
    return { x, y, w: Math.max(...rects.map((r) => r.x + r.w)) - x, h: Math.max(...rects.map((r) => r.y + r.h)) - y };
  }

  /** Bounds of a slide item: a group's own box, else the shape as parsed (placeholders may inherit theirs). */
  private itemBounds(item: Element): Rect {
    if (item.localName === "grpSp") return itemRect(item);
    return this.rendered.find((r) => r.shape.el === item)?.shape.rect ?? itemRect(item);
  }

  /** Slide-owned items on the current slide, in drawing order. */
  private slideItems(): Element[] {
    return [...new Set(this.rendered.filter((r) => !r.shape.inherited).map((r) => topLevel(r.shape.el)))];
  }

  /** One undoable XML change to a shape; keeps glued connectors attached and the shape selected. `refit` re-runs autofit. */
  private editShape(el: Element, fn: (el: Element) => void, refit = false): void {
    this.mutate(() => {
      fn(el);
      this.reroute();
    }, "slide");
    this.reselect(el);
    const rs = refit ? this.rendered.find((r) => r.shape.el === el) : null;
    if (rs) this.autofit(rs);
  }

  /** One undoable change to several items, which stay selected. */
  private editItems(items: Element[], fn: () => void): void {
    this.mutate(() => {
      fn();
      this.reroute();
    }, "slide");
    this.setSelection(items.filter((i) => i.isConnected));
  }

  private arrangeControl(): ArrangeControl {
    const items = this.items;
    const rects = () => items.map((i) => this.itemBounds(i));
    const slide = { x: 0, y: 0, w: this.doc?.slideSize.cx ?? 0, h: this.doc?.slideSize.cy ?? 0 };
    const place = (out: Rect[]) => this.editItems(items, () => items.forEach((it, k) => setItemRect(it, out[k])));
    return {
      count: items.length,
      groups: items.filter((i) => i.localName === "grpSp").length,
      align: (mode: Alignment) => place(alignRects(rects(), mode, items.length === 1 ? slide : undefined)),
      distribute: (axis) => place(distributeRects(rects(), axis)),
      group: () => this.groupSelection(),
      ungroup: () => this.ungroupSelection(),
      order: (to) => {
        // Keep the items' relative order: "front" and "backward" go in document order, the others in reverse.
        const sorted = [...items].sort((a, b) => (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1));
        const seq = to === "front" || to === "backward" ? sorted : sorted.reverse();
        this.editItems(items, () => seq.forEach((it) => reorder(it, to)));
      },
    };
  }

  private groupSelection(): void {
    const xml = this.currentSlideXml();
    if (!xml || this.items.length < 2) return;
    if (this.items.some((i) => i.getElementsByTagNameNS(P_NS, "ph").length)) {
      new Notice("Placeholders can't be grouped.");
      return;
    }
    const items = this.items;
    const rects = items.map((i) => this.itemBounds(i));
    let group: Element | null = null;
    this.mutate(() => (group = groupItems(xml, items, rects)), "slide");
    if (group) this.setSelection([group]);
  }

  private ungroupSelection(): void {
    const items = this.items;
    const out: Element[] = [];
    this.mutate(() => {
      for (const it of items) out.push(...(it.localName === "grpSp" ? ungroupItem(it) : [it]));
    }, "slide");
    this.setSelection(out.filter((i) => i.isConnected));
  }

  // --- Copy / paste -----------------------------------------------------------

  private async copySelection(): Promise<boolean> {
    if (!this.doc || !this.items.length) return false;
    const slidePath = this.editPart ?? this.doc.slides()[this.index].path;
    clipboard = { clip: await copyItems(this.doc, slidePath, this.items), slidePath, pastes: 0 };
    return true;
  }

  private paste(): void {
    if (!this.doc || !clipboard) return;
    const slidePath = this.editPart ?? this.doc.slides()[this.index].path;
    // Pasting where the items came from offsets each copy so it doesn't hide the original.
    const same = clipboard.clip.doc === this.doc && clipboard.slidePath === slidePath;
    const step = same ? ++clipboard.pastes * (EMU_PER_INCH / 10) : 0;
    const clip = clipboard.clip;
    let result: { pasted: Element[]; skipped: number } = { pasted: [], skipped: 0 };
    this.mutate((doc) => (result = pasteItems(doc, slidePath, clip, { dx: step, dy: step })), "slide");
    if (result.skipped) new Notice(`${result.skipped} item(s) have missing or unsupported linked resources.`);
    this.setSelection(result.pasted);
    const doc = this.doc;
    void this.loadMedia().then(() => { if (this.doc === doc) this.renderCurrent(); });
  }

  private async cutSelection(): Promise<void> {
    if (await this.copySelection()) this.deleteSelected();
  }

  private async duplicateSelection(): Promise<void> {
    if (await this.copySelection()) this.paste();
  }

  // --- Moving, marquee, snapping ----------------------------------------------

  /** Drag every selected item together, snapping the selection to other shapes and the slide (Alt: no snapping). */
  private startMoveItems(e: PointerEvent): void {
    e.preventDefault();
    const items = this.items.slice();
    if (!items.length || !this.doc) return;
    const start = { x: e.clientX, y: e.clientY };
    const rects0 = items.map((i) => this.itemBounds(i));
    const box0 = this.union(rects0);
    const others = this.slideItems().filter((i) => !items.includes(i)).map((i) => this.itemBounds(i));
    // Table cells ride along inside their table's element; moving them too would double the offset.
    const moving = this.rendered
      .filter((r) => r.shape.kind !== "cell" && items.includes(topLevel(r.shape.el)))
      .map((r) => ({ el: r.el, left: r.el.offsetLeft, top: r.el.offsetTop }));
    const selbox = this.slideHostEl.querySelector<HTMLElement>(".pptx-selbox");
    const box = selbox ? { left: selbox.offsetLeft, top: selbox.offsetTop } : null;
    const size = this.doc.slideSize;
    let d = { dx: 0, dy: 0 };

    const onMove = (ev: PointerEvent) => {
      d = { dx: (ev.clientX - start.x) / this.scale, dy: (ev.clientY - start.y) / this.scale };
      let guides: ({ x: number } | { y: number })[] = [];
      if (!ev.altKey) {
        const snap = snapRect({ ...box0, x: box0.x + d.dx, y: box0.y + d.dy }, others, size, SNAP_PX / this.scale);
        d = { dx: d.dx + snap.dx, dy: d.dy + snap.dy };
        guides = snap.guides;
      }
      for (const m of moving) Object.assign(m.el.style, { left: `${m.left + d.dx * this.scale}px`, top: `${m.top + d.dy * this.scale}px` });
      if (selbox && box) Object.assign(selbox.style, { left: `${box.left + d.dx * this.scale}px`, top: `${box.top + d.dy * this.scale}px` });
      this.showGuides(guides);
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      this.showGuides([]);
      if (Math.hypot(d.dx, d.dy) * this.scale < 1) return;
      this.editItems(items, () =>
        items.forEach((it, k) => {
          setItemRect(it, { ...rects0[k], x: rects0[k].x + d.dx, y: rects0[k].y + d.dy });
          // Dragging a connector away unglues it, as in PowerPoint.
          if (it.localName === "cxnSp") {
            connect(it, "st", null);
            connect(it, "end", null);
          }
        })
      );
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  /** Scale a group from its selection box corner (Shift keeps the aspect ratio). */
  private startResizeItem(e: PointerEvent, item: Element, box: HTMLElement): void {
    e.preventDefault();
    const start = { x: e.clientX, y: e.clientY };
    const r0 = this.itemBounds(item);
    let next: Rect | null = null;
    const onMove = (ev: PointerEvent) => {
      let w = Math.max(MIN_SIZE_EMU, r0.w + (ev.clientX - start.x) / this.scale);
      let h = Math.max(MIN_SIZE_EMU, r0.h + (ev.clientY - start.y) / this.scale);
      if ((ev.shiftKey || this.lockAspect) && r0.w && r0.h) {
        const k = Math.max(w / r0.w, h / r0.h);
        [w, h] = [r0.w * k, r0.h * k];
      }
      next = { ...r0, w, h };
      Object.assign(box.style, { width: `${w * this.scale}px`, height: `${h * this.scale}px` });
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const rect = next;
      if (rect) this.editItems([item], () => setItemRect(item, rect));
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  private showGuides(guides: ({ x: number } | { y: number })[]): void {
    const slide = this.slideHostEl.querySelector<HTMLElement>(".pptx-slide");
    slide?.querySelectorAll(".pptx-guide").forEach((g) => g.remove());
    for (const g of guides) {
      const line = slide?.createDiv({ cls: `pptx-guide ${"x" in g ? "is-vertical" : "is-horizontal"}` });
      if (line) line.setCssProps({ "--pptx-guide-at": `${("x" in g ? g.x : g.y) * this.scale}px` });
    }
  }

  /** Rubber-band selection: items entirely inside the dragged box (Shift adds to the selection). */
  private startMarquee(e: PointerEvent): void {
    const slide = this.slideHostEl.querySelector<HTMLElement>(".pptx-slide");
    if (!slide) return;
    const origin = slide.getBoundingClientRect();
    const at = (ev: PointerEvent) => ({ x: (ev.clientX - origin.left) / this.scale, y: (ev.clientY - origin.top) / this.scale });
    const a = at(e);
    let b = a;
    const add = e.shiftKey;
    const band = slide.createDiv({ cls: "pptx-marquee" });
    const onMove = (ev: PointerEvent) => {
      b = at(ev);
      const r = lineBox(a, b).rect;
      Object.assign(band.style, { left: `${r.x * this.scale}px`, top: `${r.y * this.scale}px`, width: `${r.w * this.scale}px`, height: `${r.h * this.scale}px` });
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      band.remove();
      const r = lineBox(a, b).rect;
      if (r.w * this.scale < 3 && r.h * this.scale < 3) {
        if (!add) this.setSelection([]);
        return;
      }
      const inside = this.slideItems().filter((i) => {
        const q = this.itemBounds(i);
        return q.x >= r.x && q.y >= r.y && q.x + q.w <= r.x + r.w && q.y + q.h <= r.y + r.h;
      });
      this.setSelection(add ? [...new Set([...this.items, ...inside])] : inside);
      this.stageEl.focus({ preventScroll: true });
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  private contextMenu(e: MouseEvent): void {
    const target = e.target as HTMLElement;
    if (target.closest("textarea, .pptx-text[contenteditable='true']:focus")) return;
    const shapeEl = target.closest<HTMLElement>(".pptx-shape:not(.is-inherited)");
    const rs = shapeEl ? this.rendered.find((r) => r.el === shapeEl) : null;
    if (rs && !this.items.includes(topLevel(rs.shape.el))) this.select(rs);
    e.preventDefault();
    const menu = new Menu();
    const item = (title: string, fn: () => void, enabled = true) => menu.addItem((i) => i.setTitle(title).setDisabled(!enabled).onClick(fn));
    const has = this.items.length > 0;
    item("Cut", () => void this.cutSelection(), has);
    item("Copy", () => void this.copySelection(), has);
    item("Paste", () => this.paste(), !!clipboard);
    item("Duplicate", () => void this.duplicateSelection(), has);
    item("Delete", () => this.deleteSelected(), has);
    menu.addSeparator();
    item("Group", () => this.groupSelection(), this.items.length > 1);
    item("Ungroup", () => this.ungroupSelection(), this.items.some((i) => i.localName === "grpSp"));
    item("Bring to front", () => this.arrangeControl().order("front"), has);
    item("Send to back", () => this.arrangeControl().order("back"), has);
    // A table cell: its row / column / merge actions (on the cells selected with Shift+click, else this one).
    const cellRs = this.rendered.find((r) => r.shape.kind === "cell" && r.el.contains(target));
    if (cellRs?.shape.cell) {
      const frame = topLevel(cellRs.shape.el);
      const table = this.rendered.find((r) => r.shape.el === frame)?.shape.table;
      const here = cellRs.shape.cell;
      const sel = this.tableCell?.frame === frame ? this.tableCell.range : null;
      const inSel = sel && here.r >= Math.min(sel[0], sel[2]) && here.r <= Math.max(sel[0], sel[2]) && here.c >= Math.min(sel[1], sel[3]) && here.c <= Math.max(sel[1], sel[3]);
      const range: CellRange = inSel ? sel : [here.r, here.c, here.r, here.c];
      if (table) {
        menu.addSeparator();
        for (const [title, fn] of tableActions(table, range)) item(title, () => this.editShape(frame, (f) => fn(f.getElementsByTagNameNS(NS.a, "tbl")[0])));
      }
    }
    // A picture: replace it, or undo a crop or colour changes.
    if (rs?.shape.kind === "pic" && !rs.shape.media) {
      const el = rs.shape.el;
      menu.addSeparator();
      item("Replace image…", () => this.replaceImage(el));
      item("Reset crop", () => this.editShape(el, (p) => setCrop(p, { l: 0, t: 0, r: 0, b: 0 })), !!rs.shape.crop);
      item("Reset colour adjustments", () => this.editShape(el, (p) => setPictureAdjust(p, NO_ADJUST)), JSON.stringify(rs.shape.adjust) !== JSON.stringify(NO_ADJUST));
    }
    menu.addSeparator();
    item("Select all", () => this.setSelection(this.slideItems()));
    menu.showAtMouseEvent(e);
  }

  // --- Autofit ----------------------------------------------------------------

  /**
   * PowerPoint's autofit, applied when editing ends: "shrink" finds the largest
   * font scale (with 20% less line spacing) that fits the box; "resize" sets
   * the shape height to the text height.
   */
  private autofit(rs: RenderedShape): void {
    const t = rs.shape.text;
    const el = rs.textEl;
    if (!t || !el || t.autofit === "none" || rs.shape.inherited || this.autofitting || !el.isConnected) return;
    // Measure on a hidden copy so the live text (and the caret / selection in it) is left alone.
    const probe = el.cloneNode(false) as HTMLElement;
    probe.removeAttribute("contenteditable");
    probe.addClass("pptx-measure");
    el.parentElement!.appendChild(probe);
    const base = EMU_PER_PT * this.scale;
    const render = (scale: number, reduction: number) =>
      renderParagraphs(probe, t.paras, base * scale, this.scale, reduction, { slideNumber: this.index + 1 });
    let apply: (() => void) | null = null;
    try {
      if (t.autofit === "resize") {
        render(t.fontScale, t.lnSpcReduction);
        Object.assign(probe.style, { bottom: "auto", height: "auto" });
        const needed = probe.offsetHeight / this.scale;
        if (Math.abs(needed - rs.shape.rect.h) > EMU_PER_PT * 2) apply = () => this.editShape(rs.shape.el, (x) => setRect(x, { ...rs.shape.rect, h: needed }));
      } else {
        const fits = () => probe.scrollHeight <= probe.clientHeight + 1;
        let scale = 1;
        let reduction = 0;
        render(1, 0);
        if (!fits()) {
          reduction = 0.2;
          let [lo, hi] = [0.25, 1];
          for (let i = 0; i < 10; i++) {
            const mid = (lo + hi) / 2;
            render(mid, reduction);
            if (fits()) lo = mid;
            else hi = mid;
          }
          // PowerPoint steps the font scale in 2.5% increments.
          scale = Math.floor(lo * 40) / 40;
        }
        if (Math.abs(scale - t.fontScale) > 0.001 || Math.abs(reduction - t.lnSpcReduction) > 0.001) {
          apply = () => this.editShape(rs.shape.el, () => writeBodyPr(t.el, { autofit: { type: "shrink", fontScale: scale, lnSpcReduction: reduction } }));
        }
      }
    } finally {
      probe.remove();
    }
    if (!apply) return;
    this.autofitting = true;
    try {
      apply();
    } finally {
      this.autofitting = false;
    }
  }


  /** Move glued connector ends after shapes moved. Re-parses the current slide. */
  private reroute(): boolean {
    const xml = this.currentSlideXml();
    if (!xml || !this.doc) return false;
    return rerouteConnectors(xml, parseSlide(this.doc, this.doc.slides()[this.index].path).shapes);
  }

  private activeTextFocused(): boolean {
    return !!this.activeText?.textEl && this.activeText.textEl.contains(document.activeElement);
  }

  private startDrag(e: PointerEvent, rs: RenderedShape, mode: "move" | "resize"): void {
    e.preventDefault();
    const start = { x: e.clientX, y: e.clientY };
    const r0 = { ...rs.shape.rect };
    const el = rs.shape.el;
    const isConnector = el.localName === "cxnSp";
    const min = isConnector ? 0 : MIN_SIZE_EMU;
    let next: Rect | null = null;

    const onMove = (ev: PointerEvent) => {
      const dx = (ev.clientX - start.x) / this.scale;
      const dy = (ev.clientY - start.y) / this.scale;
      if (mode === "move") next = { ...r0, x: r0.x + dx, y: r0.y + dy };
      else {
        let w = Math.max(min, r0.w + dx);
        let h = Math.max(min, r0.h + dy);
        // Shift, the panel's lock, or a picture: keep the aspect ratio.
        if ((ev.shiftKey || this.lockAspect || rs.shape.kind === "pic") && r0.w && r0.h) {
          const k = Math.max(w / r0.w, h / r0.h);
          [w, h] = [r0.w * k, r0.h * k];
        }
        next = { ...r0, w, h };
      }
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
      const rect = next;
      if (!rect || !this.doc) return;
      this.editShape(el, (x) => {
        setRect(x, rect);
        // Dragging a connector away unglues it, as in PowerPoint.
        if (isConnector && mode === "move") {
          connect(x, "st", null);
          connect(x, "end", null);
        }
      }, mode === "resize");
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  private startRotate(e: PointerEvent, rs: RenderedShape): void {
    e.preventDefault();
    const box = rs.el.getBoundingClientRect();
    const cx = box.left + box.width / 2;
    const cy = box.top + box.height / 2;
    let angle: number | null = null;
    const onMove = (ev: PointerEvent) => {
      angle = (Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey) angle = Math.round(angle / 15) * 15;
      angle = ((angle % 360) + 360) % 360;
      rs.el.style.transform = `rotate(${angle}deg)`;
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const rot = angle;
      if (rot !== null) this.editShape(rs.shape.el, (x) => setTransform(x, { rot }));
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  /** Yellow handles that change a preset's adjust values (arrow head size, corner rounding...). */
  private addAdjustHandles(rs: RenderedShape): void {
    const { shape } = rs;
    if (shape.inherited || shape.grouped || shape.geom === "custom" || !shape.handles.length) return;
    const place = (h: HTMLElement, x: number, y: number) =>
      Object.assign(h.style, {
        left: `${(shape.flipH ? shape.rect.w - x : x) * this.scale}px`,
        top: `${(shape.flipV ? shape.rect.h - y : y) * this.scale}px`,
      });
    shape.handles.forEach((handle, i) => {
      const h = rs.el.createDiv({ cls: "pptx-handle pptx-handle-adjust", attr: { "aria-label": "Drag to adjust the shape" } });
      place(h, handle.x, handle.y);
      h.addEventListener("pointerdown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        const box = rs.el.getBoundingClientRect();
        const [cx, cy] = [box.left + box.width / 2, box.top + box.height / 2];
        const a = (-shape.rot * Math.PI) / 180;
        const { w, h: ht } = shape.rect;
        let adj = shape.adj;
        const onMove = (ev: PointerEvent) => {
          // Pointer in the shape's own (unrotated, unflipped) coordinates.
          const [dx, dy] = [ev.clientX - cx, ev.clientY - cy];
          let x = (dx * Math.cos(a) - dy * Math.sin(a)) / this.scale + w / 2;
          let y = (dx * Math.sin(a) + dy * Math.cos(a)) / this.scale + ht / 2;
          if (shape.flipH) x = w - x;
          if (shape.flipV) y = ht - y;
          adj = solveAdjust(shape.geom, w, ht, shape.adj, i, { x, y });
          const moved = presetGeometry(shape.geom, w, ht, adj)?.handles[i];
          if (moved) place(h, moved.x, moved.y);
        };
        const onUp = () => {
          window.removeEventListener("pointermove", onMove);
          window.removeEventListener("pointerup", onUp);
          const changed = Object.entries(adj).filter(([k, v]) => shape.adj[k] !== v);
          if (changed.length) this.editShape(shape.el, (x) => changed.forEach(([k, v]) => setAdjust(x, k, v)));
        };
        window.addEventListener("pointermove", onMove);
        window.addEventListener("pointerup", onUp);
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Drawing new shapes
  // ---------------------------------------------------------------------------

  private toggleGallery(anchor: HTMLElement): void {
    if (this.galleryEl) return this.closeGallery();
    const g = (this.galleryEl = this.contentEl.createDiv({ cls: "pptx-gallery" }));
    buildShapeGallery(g, (choice) => {
      this.closeGallery();
      this.drawing = choice;
      this.stageEl.addClass("is-drawing");
      new Notice("Drag on the slide to draw the shape, or click for a default size. Esc cancels.");
    });
    const a = anchor.getBoundingClientRect();
    const c = this.contentEl.getBoundingClientRect();
    const left = Math.max(0, Math.min(a.left - c.left, c.width - g.offsetWidth - 8));
    g.setCssProps({ "--pptx-gallery-top": `${a.bottom - c.top + 4}px`, "--pptx-gallery-left": `${left}px` });
    const outside = (e: PointerEvent) => {
      if (g.contains(e.target as Node) || anchor.contains(e.target as Node)) return;
      this.closeGallery();
    };
    window.addEventListener("pointerdown", outside, { capture: true });
    this.closeGalleryListener = () => window.removeEventListener("pointerdown", outside, { capture: true });
  }

  private closeGalleryListener: (() => void) | null = null;

  private closeGallery(): void {
    this.closeGalleryListener?.();
    this.closeGalleryListener = null;
    this.galleryEl?.remove();
    this.galleryEl = null;
  }

  private cancelDrawing(): void {
    this.drawing = null;
    this.stageEl.removeClass("is-drawing");
  }

  /** The connection site of a slide shape near `p` (EMU), for gluing connector ends. */
  private snapToSite(p: { x: number; y: number }): { p: { x: number; y: number }; glue: Glue | null } {
    let best: { p: { x: number; y: number }; glue: Glue; d: number } | null = null;
    for (const rs of this.rendered) {
      const s = rs.shape;
      if (s.inherited || s.el.localName === "cxnSp") continue;
      s.sites.forEach((_, idx) => {
        const q = sitePoint(s, idx);
        const d = q ? Math.hypot(q.x - p.x, q.y - p.y) : Infinity;
        if (q && d * this.scale <= SNAP_PX && (!best || d < best.d)) best = { p: q, glue: { id: shapeId(s.el), idx }, d };
      });
    }
    const found = best as { p: { x: number; y: number }; glue: Glue } | null;
    return found ? { p: found.p, glue: found.glue } : { p, glue: null };
  }

  private startDraw(e: PointerEvent): void {
    const choice = this.drawing;
    const slideEl = this.slideHostEl.querySelector<HTMLElement>(".pptx-slide");
    if (!choice || !slideEl || !this.doc) return;
    e.preventDefault();
    e.stopPropagation();
    const box = slideEl.getBoundingClientRect();
    const toEmu = (ev: PointerEvent) => ({ x: (ev.clientX - box.left) / this.scale, y: (ev.clientY - box.top) / this.scale });
    const isLine = CONNECTORS.has(choice.prst);
    const snap = (p: { x: number; y: number }) => (isLine ? this.snapToSite(p) : { p, glue: null });
    const a = snap(toEmu(e));
    let b = a;
    const preview = slideEl.createDiv({ cls: "pptx-draw-preview" });

    const onMove = (ev: PointerEvent) => {
      let p = toEmu(ev);
      if (ev.shiftKey) {
        const [dx, dy] = [p.x - a.p.x, p.y - a.p.y];
        if (isLine) {
          // Snap to multiples of 45 degrees.
          const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
          const len = Math.hypot(dx, dy);
          p = { x: a.p.x + len * Math.cos(ang), y: a.p.y + len * Math.sin(ang) };
        } else {
          const side = Math.max(Math.abs(dx), Math.abs(dy));
          p = { x: a.p.x + Math.sign(dx || 1) * side, y: a.p.y + Math.sign(dy || 1) * side };
        }
      }
      b = snap(p);
      const r = lineBox(a.p, b.p).rect;
      Object.assign(preview.style, {
        left: `${r.x * this.scale}px`,
        top: `${r.y * this.scale}px`,
        width: `${r.w * this.scale}px`,
        height: `${r.h * this.scale}px`,
      });
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      preview.remove();
      this.cancelDrawing();
      const dragged = Math.hypot(b.p.x - a.p.x, b.p.y - a.p.y) * this.scale > 4;
      let box2: { rect: Rect; flipH: boolean; flipV: boolean };
      if (isLine) box2 = lineBox(a.p, dragged ? b.p : { x: a.p.x + EMU_PER_INCH, y: a.p.y });
      else
        box2 = dragged
          ? { rect: lineBox(a.p, b.p).rect, flipH: false, flipV: false }
          : { rect: { x: a.p.x - EMU_PER_INCH / 2, y: a.p.y - EMU_PER_INCH / 2, w: EMU_PER_INCH, h: EMU_PER_INCH }, flipH: false, flipV: false };
      let added: Element | null = null;
      this.mutate(() => {
        added = addShape(this.currentSlideXml()!, choice.prst, box2.rect, { flipH: box2.flipH, flipV: box2.flipV, tailArrow: choice.tailArrow });
        if (isLine) {
          connect(added, "st", a.glue);
          connect(added, "end", dragged ? b.glue : null);
          this.reroute();
        }
      }, "slide");
      this.reselect(added);
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
      this.updateFormatControls(rs, range);
    }
  };

  /** Font and Size show the run at the cursor (from the model: the DOM only carries what differs from defaults). */
  private updateFormatControls(rs: RenderedShape, range: Range): void {
    const text = rs.shape.text;
    if (!text || !rs.textEl || !this.fontSelect || !this.sizeSelect) return;
    const ps = Array.from(rs.textEl.querySelectorAll(":scope > p"));
    const p = ps.find((el) => el.contains(range.startContainer));
    const para = p ? text.paras[ps.indexOf(p)] : undefined;
    if (!p || !para) return;
    const [offset] = textOffsets(p as HTMLElement, range);
    let seen = 0;
    const run = para.runs.find((r) => (seen += r.text.length) >= offset) ?? para.runs[para.runs.length - 1];
    const props = run ?? para.defaults;
    showValue(this.fontSelect, props.font || null);
    showValue(this.sizeSelect, props.sz ? String(Math.round(props.sz * 10) / 10) : null);
  }

  private scheduleCommit(rs: RenderedShape): void {
    if (this.pendingCommit && this.pendingCommit !== rs) this.flushCommit();
    this.pendingCommit = rs;
    this.autosave.status?.saving();
    if (this.commitTimer) window.clearTimeout(this.commitTimer);
    this.commitTimer = window.setTimeout(() => this.flushCommit(), 400);
  }

  /** Write pending speaker-notes edits for the current slide. */
  private flushNotes(): void {
    if (this.notesTimer) window.clearTimeout(this.notesTimer);
    this.notesTimer = null;
    if (!this.doc || this.notesEl.value === this.doc.notesText(this.index)) return;
    this.pushHistory();
    this.doc.setNotes(this.index, this.notesEl.value);
    this.autosave.markDirty();
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
    this.needsFit.add(rs.shape.el);
    if (!this.doc || !rs.textEl || !rs.shape.text) return;
    this.pushHistory();
    writeParagraphs(rs.shape.text, paras ?? readParagraphs(rs.textEl, rs.pxPerPt, this.defaultsFor(rs)));
    this.refreshModel();
    this.autosave.markDirty();
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
    const after = sel.rangeCount ? sel.getRangeAt(0) : null;
    const offsets = after && textEl.contains(after.commonAncestorContainer) ? textOffsets(textEl, after) : null;
    if (this.pendingCommit === rs) this.pendingCommit = null;
    this.commitText(rs);
    // An inner span can't cancel an outer span's underline in CSS: redraw flat runs from the saved model.
    if (rs.shape.text) {
      renderParagraphs(textEl, rs.shape.text.paras, rs.pxPerPt, this.scale, rs.shape.text.lnSpcReduction, { slideNumber: this.index + 1 });
      if (offsets) {
        const range = rangeAt(textEl, offsets);
        sel.removeAllRanges();
        sel.addRange(range);
        this.savedRange = range.cloneRange();
      }
    }
    return rs;
  }

  private toggle(prop: "b" | "i" | "u" | "strike"): void {
    this.withText((rs, el) => {
      const current = selectionRunProps(el, rs.pxPerPt, this.defaultsFor(rs));
      applyRunStyle(el, { [prop]: !current?.[prop] }, rs.pxPerPt);
    });
  }

  /** Superscript (30% up) / subscript (25% down), as PowerPoint and Google Slides set them; again turns it off. */
  private toggleBaseline(shift: number): void {
    this.withText((rs, el) => {
      const current = selectionRunProps(el, rs.pxPerPt, this.defaultsFor(rs))?.baseline ?? 0;
      applyRunStyle(el, { baseline: current === shift ? 0 : shift }, rs.pxPerPt);
    });
  }

  private setRunStyle(patch: Partial<RunProps>): void {
    this.withText((rs, el) => applyRunStyle(el, patch, rs.pxPerPt));
  }

  /** Bulleted list / Numbered list: on, or off when the paragraphs already have that kind. */
  private toggleList(numbered: boolean): void {
    this.updateParas((p) => ({ ...p, bullet: numbered ? (p.bullet?.startsWith("#") ? null : "#arabicPeriod") : p.bullet && !p.bullet.startsWith("#") ? null : "•" }));
  }

  private updateParas(fn: Parameters<typeof updateParagraphs>[1]): void {
    this.withText((_, el) => updateParagraphs(el, fn), false);
  }

  /** Slide titles for the link dialog: the title placeholder's text, else "Slide N". */
  private slideTitles(): string[] {
    const doc = this.doc!;
    return doc.slides().map((s, i) => {
      const title = parseSlide(doc, s.path).shapes.find((sh) => sh.placeholder && !sh.inherited && sh.text && /title/i.test(sh.el.getElementsByTagNameNS(NS.p, "ph")[0]?.getAttribute("type") ?? ""));
      const text = title?.text ? plainText(title.text.el).trim() : "";
      return text ? `Slide ${i + 1}: ${text}` : `Slide ${i + 1}`;
    });
  }

  /** Follow a link from the editor (Ctrl+click): slides in this deck open here, web addresses in the browser. */
  private followLink(link: string): void {
    if (!link.startsWith("#")) return void window.open(link);
    const to = linkTarget(link, this.index, this.doc?.slides().map((s) => s.path) ?? []);
    if (to !== null) this.goTo(to);
  }

  /**
   * Google's Insert link: while editing text it links the selected text, otherwise the selected shapes. A web
   * address or a slide of this deck.
   */
  private async editLink(): Promise<void> {
    if (!this.doc) return;
    const rs = this.activeText;
    const slides = this.doc.slides().map((s) => s.path);
    if (!rs?.textEl) {
      const items = this.items;
      if (!items.length) {
        new Notice("Select a shape or some text first.");
        return;
      }
      const part = this.editPart ?? slides[this.index];
      const current = this.rendered.find((r) => r.shape.el === items[0])?.shape.link ?? "";
      const link = await linkDialog(this.app, current, slides, this.slideTitles());
      if (link === undefined) return;
      this.mutate((doc) => {
        for (const item of items) setShapeLink(doc, part, item, link);
      }, "slide");
      return;
    }
    const current = selectionRunProps(rs.textEl, rs.pxPerPt, this.defaultsFor(rs))?.link ?? "";
    // The prompt takes focus and the text may be redrawn meanwhile: remember the selection as character offsets.
    const offsets = this.savedRange && rs.textEl.contains(this.savedRange.commonAncestorContainer) ? textOffsets(rs.textEl, this.savedRange) : null;
    const link = await linkDialog(this.app, current, slides, this.slideTitles());
    if (link === undefined) return;
    const live = this.rendered.find((r) => r.shape.el === rs.shape.el) ?? rs;
    this.activeText = live;
    this.savedRange = offsets && live.textEl ? rangeAt(live.textEl, offsets) : null;
    this.withText((r, el) => applyRunStyle(el, { link }, r.pxPerPt));
  }

  /** Without a caret in the text box, insert at its end rather than over everything. */
  private caretOrEnd(): void {
    const el = (this.activeText ?? this.selected)?.textEl;
    if (el && !(this.savedRange && el.contains(this.savedRange.commonAncestorContainer))) {
      const end = document.createRange();
      end.selectNodeContents(el.lastElementChild ?? el);
      end.collapse(false);
      this.savedRange = end;
    }
  }

  private fillFieldMenu(menu: Menu): void {
    const add = (title: string, type: string, text: () => string) =>
      menu.addItem((i) =>
        i.setTitle(title).onClick(() => {
          this.caretOrEnd();
          this.withText((_, host) => insertField(host, type, text()), false);
        })
      );
    add("Slide number", "slidenum", () => String(this.index + 1));
    add("Date", "datetime1", () => new Date().toLocaleDateString());
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
    renderParagraphs(rs.textEl, rs.shape.text.paras, rs.pxPerPt, this.scale, rs.shape.text.lnSpcReduction, { slideNumber: this.index + 1 });
  }

  private onTextKey(e: KeyboardEvent, rs: RenderedShape): void {
    const key = e.key.toLowerCase();
    if (isMod(e) && (key === "z" || key === "y")) {
      e.preventDefault();
      // The slide is redrawn: put the caret back in the same box, at the same place in its text.
      const ids = rs.el.dataset.spids;
      const at = rs.textEl ? caretOffset(rs.textEl) : 0;
      if (key === "y" || e.shiftKey) this.redo();
      else this.undo();
      const again = ids ? this.rendered.find((r) => r.el.dataset.spids === ids)?.textEl : null;
      if (again) {
        again.focus();
        placeCaret(again, at);
      }
    } else if (isMod(e) && (key === "b" || key === "i" || key === "u")) {
      e.preventDefault();
      this.toggle(key);
    } else if (e.key === "Tab" && rs.shape.kind === "cell") {
      e.preventDefault();
      this.nextCell(rs, e.shiftKey ? -1 : 1);
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
    } else if (e.key === "Escape" && this.painter) {
      this.stopPainting();
    } else if (e.key === "Escape" && (this.drawing || this.galleryEl)) {
      this.cancelDrawing();
      this.closeGallery();
    } else if (e.key.startsWith("Arrow") && this.items.length) {
      // Nudge: 0.1", or one pixel with Alt.
      e.preventDefault();
      const step = e.altKey ? 1 / this.scale : EMU_PER_INCH / 10;
      const [dx, dy] = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key] ?? [0, 0];
      const items = this.items;
      const rects = items.map((i) => this.itemBounds(i));
      this.editItems(items, () => items.forEach((it, k) => setItemRect(it, { ...rects[k], x: rects[k].x + dx, y: rects[k].y + dy })));
    } else if ((e.key === "Delete" || e.key === "Backspace") && this.items.length) {
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
    this.editPart = null; this.showTheme = false;
    this.renderCurrent();
    (this.gridView ? this.thumbsEl : this.stageEl).focus({ preventScroll: true });
  }

  /** File › Download, as Google Slides: the deck as PDF or the current slide as PNG (desktop: both print through Electron). */
  async exportAs(initial?: "pdf" | "png"): Promise<void> {
    const doc = this.doc, file = this.file;
    if (!doc || !file) return;
    this.flushCommit();
    const i = this.index;
    const choice = await exportDialog(this.app, file, [
      { id: "pdf", label: "PDF document (.pdf)", ext: "pdf", outside: true },
      { id: "png", label: "PNG image of the current slide (.png)", ext: "png", outside: true, name: `${file.basename} - slide ${i + 1}` },
    ], initial);
    if (!choice) return;
    const data = choice.format === "png" ? await this.slidePng(doc, i) : await this.deckPdf(doc);
    if (data) await writeExport(this.app, choice.target, data, choice.open);
  }

  /** The slide, as large as the window allows, captured at full HD. */
  private slidePng(doc: PptxDocument, i: number): Promise<Uint8Array | null> {
    const { cx, cy } = doc.slideSize;
    const width = Math.floor(Math.min(window.innerWidth, (window.innerHeight * cx) / cy, 1920));
    const height = Math.floor((width * cy) / cx);
    return renderPng({ width, height }, (host) =>
      renderSlide(host, parseSlide(doc, doc.slides()[i].path), doc.slideSize, { width, editable: false, imageUrl: this.imageUrl, slideNumber: i + 1 })
    );
  }

  /** The slides (not hidden ones) as a PDF, one slide per page. */
  private deckPdf(doc: PptxDocument): Promise<Uint8Array | null> {
    const width = doc.slideSize.cx / 9525;
    const height = doc.slideSize.cy / 9525;
    const pages = doc.slides().flatMap((s, i) => {
      if (doc.isHidden(i)) return [];
      const page = createDiv({ cls: "dw-print-page" });
      renderSlide(page, parseSlide(doc, s.path), doc.slideSize, { width, editable: false, imageUrl: this.imageUrl, slideNumber: i + 1 });
      return [page];
    });
    return renderPdf(pages, { width, height });
  }

  /** Run the slideshow from the first visible slide, or from the current one. */
  startPresentation(fromStart: boolean, previewTransition = false): void {
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
      slidePaths: doc.slides().map((s) => s.path),
      render: (host: HTMLElement, i: number, width: number) =>
        renderSlide(host, parseSlide(doc, doc.slides()[i].path), doc.slideSize, { width, editable: false, imageUrl: this.imageUrl, slideNumber: i + 1, playMedia: true }),
      effects: (i: number) => {
        const xml = doc.xml(doc.slides()[i].path);
        return { transition: xml && readTransition(xml), steps: xml ? readAnimations(xml) : [] };
      },
    };
    new Slideshow(deck, start, (i) => this.goTo(i)).open(previewTransition);
  }

  private pushHistory(): void {
    if (this.doc) this.history.push({ snap: this.doc.snapshot(), index: this.index, editPart: this.editPart });
  }

  /** Snapshot, apply a change, then redraw everything (or only the current slide and its thumbnail). */
  private mutate(fn: (doc: PptxDocument) => void, scope: "all" | "slide" = "all"): void {
    if (!this.doc) return;
    this.flushCommit();
    const before = this.doc.snapshot();
    const index = this.index;
    const editPart = this.editPart;
    try {
      fn(this.doc);
    } catch (err) {
      this.doc.restore(before);
      this.index = index;
      new Notice((err as Error).message);
      return;
    }
    this.history.push({ snap: before, index, editPart });
    this.autosave.markDirty();
    if (scope === "all") this.renderAll();
    else {
      this.renderCurrent();
      this.refreshThumb(this.index);
    }
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
    const target = step({ snap: this.doc.snapshot(), index: this.index, editPart: this.editPart });
    if (!target) return;
    this.doc.restore(target.snap);
    this.index = target.index;
    this.editPart = target.editPart ?? null;
    this.autosave.markDirty();
    this.renderAll();
  }

  private pickMedia(): void {
    const input = createEl("input", { type: "file", attr: { accept: Object.keys(MEDIA_MIME).map(ext => "." + ext).join(",") } });
    input.addEventListener("change", () => { const file = input.files?.[0]; if (file)
      void this.insertMediaFile(file); });
    input.click();
  }
  private async insertMediaFile(file: File): Promise<void> {
    if (!this.doc)
      return;
    const doc = this.doc;
    const part = this.editPart ?? doc.slides()[this.index].path;
    const ext = extOf(file.name);
    const mime = MEDIA_MIME[ext];
    if (!mime) {
      new Notice("Choose an audio or video file in a supported format.");
      return;
    }
    const bytes = new Uint8Array(await file.arrayBuffer());
    const canvas=createEl("canvas");canvas.width=320;canvas.height=180;
    const ctx=canvas.getContext("2d");
    if(ctx){ctx.fillStyle="#202c40";ctx.fillRect(0,0,320,180);ctx.fillStyle="white";ctx.font="64px sans-serif";ctx.textAlign="center";ctx.fillText(mime.startsWith("video/")?"▶":"♫",160,112);}
    const poster=await new Promise<Uint8Array|undefined>(resolve=>canvas.toBlob(blob=>{if(blob)void blob.arrayBuffer().then(b=>resolve(new Uint8Array(b)));else resolve(undefined);},"image/png"));

    if (this.doc !== doc || (this.editPart ?? doc.slides()[this.index].path) !== part)
      return;
    this.mutate(d => { const shape = addMedia(d, part, bytes, mime.startsWith("video/") ? "video" : "audio", ext, file.name, poster); const media = parseSlide(d, part).shapes.find(s => s.el === shape)!.media!; this.imageUrls.set(media.path, URL.createObjectURL(new Blob([bytes], { type: mime }))); });
    await this.loadMedia();
    this.renderCurrent();
  }
  private author(): string {
    return this.settings.author.trim() || "Doc Works";
  }

  /** The current slide's comment threads (Google Slides' comments): post, reply, edit, resolve, delete. */
  private commentsDialog(hit?: { messageId?: string; start: number; end: number }): void {
    const doc = this.doc;
    if (!doc) return;
    const slide = doc.slides()[this.index]?.path;
    if (!slide) return;
    const modal = new Modal(this.app);
    modal.titleEl.setText(`Comments on slide ${this.index + 1}`);
    modal.contentEl.addClass("xlsx-comments-modal");
    const change = (fn: (d: PptxDocument) => void) => {
      this.mutate(fn);
      draw();
    };
    const button = (parent: HTMLElement, text: string, fn: () => void) => {
      const b = parent.createEl("button", { text });
      b.addEventListener("click", fn);
      return b;
    };
    const compose = (parent: HTMLElement, label: string, post: (text: string) => void, value = "") => {
      const input = parent.createEl("textarea", { value, attr: { "aria-label": label, rows: "3", placeholder: label } });
      button(parent, value ? "Save" : "Post", () => {
        if (input.value.trim()) post(input.value);
      });
    };
    const draw = () => {
      modal.contentEl.empty();
      const threads = slideComments(doc, slide);
      if (!threads.length) modal.contentEl.createEl("p", { text: "No comments on this slide." });
      for (const thread of threads) {
        const box = modal.contentEl.createDiv({ cls: "xlsx-comment-thread" });
        if (thread.resolved) box.createDiv({ text: "Resolved", cls: "dw-muted" });
        for (const message of thread.messages) {
          const item = box.createDiv({ cls: "xlsx-comment-message" });
          item.createEl("strong", { text: message.author || "Unknown author" });
          if (message.date) item.createEl("time", { text: new Date(message.date).toLocaleString(), attr: { datetime: message.date } });
          const textEl = item.createDiv({ cls: "xlsx-comment-text" });
          if (hit?.messageId === message.id) {
            textEl.appendText(message.text.slice(0, hit.start));
            textEl.createEl("mark", { text: message.text.slice(hit.start, hit.end) });
            textEl.appendText(message.text.slice(hit.end));
          } else textEl.setText(message.text);
          if (thread.legacy) continue;
          button(item, "Edit", () => {
            item.empty();
            compose(item, "Edit comment", (text) => change((d) => editSlideComment(d, slide, message.id, text)), message.text);
            button(item, "Cancel", draw);
          });
        }
        const actions = box.createDiv({ cls: "xlsx-comment-actions" });
        if (!thread.legacy) button(actions, thread.resolved ? "Reopen" : "Resolve", () => change((d) => resolveSlideThread(d, slide, thread.id, !thread.resolved)));
        button(actions, "Delete thread", () => change((d) => removeSlideThread(d, slide, thread.id)));
        if (!thread.resolved && !thread.legacy) compose(box, "Reply", (text) => change((d) => void addSlideComment(d, slide, text, this.author(), thread.id)));
      }
      compose(modal.contentEl, "New comment", (text) => change((d) => void addSlideComment(d, slide, text, this.author())));
      button(modal.contentEl, "Close", () => modal.close());
    };
    draw();
    modal.open();
    modal.contentEl.querySelector("mark")?.scrollIntoView({ block: "nearest" });
  }

  /** Edit the text of a SmartArt's items (in its data and the drawing shown). */
  private editSmartArt(frame: Element): void {
    const doc = this.doc;
    if (!doc) return;
    const part = this.editPart ?? doc.slides()[this.index].path;
    const parts = smartArtParts(doc, frame, part);
    const data = parts && doc.xml(parts.data);
    if (!parts || !data) return void new Notice("This diagram's data couldn't be found.");
    const items = smartArtItems(data);
    const modal = new Modal(this.app);
    modal.titleEl.setText("Diagram text");
    modal.contentEl.createDiv({ cls: "pptx-props-hint", text: "Edit each item's text, or add and remove top-level items." });
    const fields: HTMLTextAreaElement[] = [];
    // Typed text plus one structural change, as one undo step; the dialog then shows the new items.
    const apply = (change?: (data: Document, drawing: Document | null) => void) => {
      const changed = items.filter((item, i) => fields[i].value !== item.text);
      if (!changed.length && !change) return modal.close();
      // mutate rolls back and shows the message when a change is refused.
      this.mutate((d) => {
        const drawing = parts.drawing ? d.xml(parts.drawing) : null;
        change?.(d.xml(parts.data)!, drawing);
        for (const item of changed) setSmartArtText(d.xml(parts.data)!, drawing, item.id, fields[items.indexOf(item)].value);
      });
      modal.close();
      if (change) this.editSmartArt(frame);
    };
    for (const item of items) {
      const row = modal.contentEl.createDiv({ cls: "pptx-smartart-row" });
      const field = row.createEl("textarea", { cls: "pptx-smartart-item", attr: { rows: "1", "aria-label": `Item ${item.text}` } });
      field.value = item.text;
      field.setCssProps({ "--pptx-depth": String(item.depth) });
      fields.push(field);
      if (item.depth) continue;
      const btn = (icon: string, label: string, fn: () => void) => {
        const b = row.createEl("button", { cls: "clickable-icon", attr: { "aria-label": label } });
        setIcon(b, icon);
        b.addEventListener("click", fn);
      };
      btn("plus", "Add item after", () => apply((d, dr) => void addSmartArtItem(d, dr, item.id, "Text")));
      btn("trash-2", "Remove item", () => apply((d, dr) => removeSmartArtItem(d, dr, item.id)));
    }
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Apply").setCta().onClick(() => apply()))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }

  private editChart(path?: string, kind: ChartKind = "column"): void {
    if (!this.doc)
      return;
    const doc = this.doc;
    const slide = this.editPart ?? doc.slides()[this.index].path;
    const model = path ? parseSlide(doc, slide).shapes.find(s => s.chart?.path === path)?.chart : undefined;
    if (path && !model?.data)
      return;
    const colors = slideColors(doc, slide);
    chartDialog(this.app, {
      path, data: model?.data, format: model?.format ?? newChartFormat(kind), palette: chartPalette(colors),
      apply: async (data, format) => {
        const update = path ? await prepareChartEdit(doc, path, data) : null;
        if (update) setChartFormat(update.xml, format, colors);
        const workbook = path ? null : await chartWorkbook(data);
        if (this.doc !== doc || (this.editPart ?? doc.slides()[this.index].path) !== slide)
          throw new Error("Return to the slide where this editor was opened.");
        this.mutate(d => { if (path && update) {
          d.putXml(path, update.xml);
          if (update.workbookPath && update.workbook)
            d.putBinary(update.workbookPath, update.workbook);
        }
        else
          addChart(d, slide, data, "barChart", workbook!, f => Object.assign(f, structuredClone(format))); });
      },
    });
  }
  private renderThemePanel(): void {
    if (!this.doc)
      return;
    const doc = this.doc;
    const layout = doc.layoutOf(doc.slides()[this.index].path)!;
    const master = doc.masterOf(layout)!;
    const theme = doc.themeOf(master)!;
    const host = this.propsEl;
    host.replaceChildren();
    host.createEl("h3", { text: "Theme and layouts" });
    const select = host.createEl("select", { cls: "dropdown", attr: { "aria-label": "Edit master or layout" } });
    select.createEl("option", { text: "Slides", value: "" });
    select.createEl("option", { text: "Master", value: master });
    for (const item of doc.layouts(master))
      select.createEl("option", { text: item.name || item.type, value: item.path });
    select.value = this.editPart ?? "";
    select.addEventListener("change", () => { this.flushCommit(); this.editPart = select.value || null; this.setSelection([]); this.renderCurrent(); });
    host.createDiv({ cls: "pptx-props-hint", text: this.editPart ? "Edit objects on the canvas. Changes apply to slides using this master or layout." : "Choose a master or layout to edit its objects." });
    const create = host.createEl("button", { text: "Create custom layout" });
    create.addEventListener("click", () => {
      void (async () => {
        const name = await promptText(this.app, "Layout name", "Custom layout");
        if (name && this.doc === doc)
          this.mutate(d => { this.editPart = d.createLayout(this.editPart && d.masterOf(this.editPart) ? this.editPart : layout, name); });
      })();
    });
    if (this.editPart)
      for (const type of ["title", "body"] as const) {
        const button = host.createEl("button", { text: `Add ${type} placeholder` });
        button.addEventListener("click", () => this.mutate(d => d.addLayoutPlaceholder(this.editPart!, type)));
      }
    const colors = doc.xml(theme)?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "clrScheme")[0];
    for (const slot of Array.from(colors?.children ?? [])) {
      const label = host.createEl("label", { cls: "pptx-props-row" });
      label.createSpan({ text: slot.localName });
      const field = label.createEl("input", { type: "color", attr: { "aria-label": "Theme " + slot.localName } });
      field.value = "#" + (slot.firstElementChild?.getAttribute("lastClr") ?? slot.firstElementChild?.getAttribute("val") ?? "000000");
      field.addEventListener("change", () => this.mutate(d => d.setTheme(theme, { [slot.localName]: field.value.slice(1) })));
    }
    for (const [label, tag] of [["Heading font", "majorFont"], ["Body font", "minorFont"]]) {
      const row = host.createEl("label", { cls: "pptx-props-row" });
      row.createSpan({ text: label });
      const field = row.createEl("input", { type: "text", attr: { "aria-label": label } });
      field.value = doc.xml(theme)?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", tag)[0]?.firstElementChild?.getAttribute("typeface") ?? "";
      field.addEventListener("change", () => this.mutate(d => d.setTheme(theme, {}, tag === "majorFont" ? field.value : undefined, tag === "minorFont" ? field.value : undefined)));
    }
    const back = host.createEl("button", { text: "Done editing theme" });
    back.addEventListener("click", () => { this.flushCommit(); this.editPart = null; this.showTheme = false; this.renderCurrent(); });
    host.show();
  }
  private currentSlideXml(): Document | null {
    const slide = this.doc?.slides()[this.index];
    return this.editPart ? this.doc!.xml(this.editPart) : slide ? this.doc!.xml(slide.path) : null;
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
    const items = this.items;
    if (!items.length) return;
    this.mutate(() => items.forEach((i) => i.remove()), "slide");
    this.setSelection([]);
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
      const path = this.editPart ?? doc.slides()[this.index].path;
      const rId = doc.addImage(path, bytes, ext);
      this.imageUrls.set(doc.relTarget(path, rId)!, url);
      addPicture(doc.xml(path)!, { x: (cx - w) / 2, y: (cy - h) / 2, w, h }, rId, file.name);
    });
  }

  // ---------------------------------------------------------------------------
  // Find and replace, zoom
  // ---------------------------------------------------------------------------

  private findEl: HTMLElement | null = null;
  private zoomSelect: HTMLSelectElement | null = null;
  private findInput!: HTMLInputElement;
  private findMatches: SlideSearchMatch[] = [];
  private findIndex = -1;
  private zoom = "fit";

  private openFind(): void {
    if (!this.doc) return;
    if (this.findEl) {
      this.findInput.focus();
      this.findInput.select();
      return;
    }
    const bar = findBar(this.contentEl, this.contentEl.querySelector(".pptx-body"), {
      options: [["Aa", "Match case", false], ["Notes", "Search speaker notes", true], ["Comments", "Search comments", true]],
      replaceTitles: ["Replace this slide-text match", "Replace matching slide text"],
      changed: () => {
        refresh();
        if (this.findMatches[this.findIndex]) show(this.findMatches[this.findIndex]);
      },
      step: (dir) => step(dir),
      replace: () => {
        refresh(true);
        const m = this.findMatches[this.findIndex];
        if (!m || m.kind !== "text") return;
        this.mutate(() => replaceMatch(m, bar.replaceInput.value), "slide");
        refresh(true);
        this.findIndex--;
        step(1);
      },
      replaceAll: () => {
        let n = 0;
        this.mutate((d) => (n = replaceAll(d, this.findInput.value, bar.replaceInput.value, opts())));
        new Notice(`Replaced ${n} ${n === 1 ? "match" : "matches"}.`);
        refresh();
      },
      close: () => this.closeFind(),
    });
    this.findEl = bar.el;
    this.findInput = bar.input;
    const { preview } = bar;
    const [matchCase, notes, comments] = bar.options;
    const opts = () => ({ matchCase: matchCase.checked, notes: notes.checked, comments: comments.checked });
    const refresh = (keep = false) => {
      this.findMatches = findSlideMatches(this.doc!, this.findInput.value, opts());
      if (!keep) this.findIndex = this.findMatches.findIndex((m) => m.slide >= this.index);
      if (this.findMatches.length && this.findIndex < 0) this.findIndex = 0;
      this.findIndex = Math.min(this.findIndex, this.findMatches.length - 1);
      bar.setCount(this.findInput.value, this.findIndex, this.findMatches.length);
      preview.empty();
      this.clearNotesHighlight();
      if ("highlights" in CSS) highlights().delete("dw-find");
      bar.replaceButton.disabled = this.findMatches[this.findIndex]?.kind !== "text";
      bar.replaceAllButton.disabled = !this.findMatches.some(m => m.kind === "text");
    };
    const step = (dir: number) => {
      refresh(true);
      if (!this.findMatches.length) return;
      this.findIndex = (this.findIndex + dir + this.findMatches.length) % this.findMatches.length;
      bar.setCount(this.findInput.value, this.findIndex, this.findMatches.length);
      show(this.findMatches[this.findIndex]);
    };
    const show = (m: SlideSearchMatch) => {
      this.showMatch(m);
      bar.replaceButton.disabled = m.kind !== "text";
      if (m.kind !== "comment") { preview.empty(); return; }
      showFindPreview(preview, `Slide ${m.slide + 1} · ${m.resolved ? "Resolved comment" : "Comment"}${m.author ? ` · ${m.author}` : ""}`, m.text, m.start, m.end,
        "Open comment", () => {
          if (this.index !== m.slide) this.goTo(m.slide);
          this.commentsDialog(m);
        });
    };
    this.flushNotes();
    refresh();
    this.findInput.focus();
  }

  private closeFind(): void {
    this.findEl?.remove();
    this.findEl = null;
    this.findMatches = [];
    this.clearNotesHighlight();
    if ("highlights" in CSS) highlights().delete("dw-find");
    this.stageEl.focus({ preventScroll: true });
  }

  /** Go to a match's slide and highlight it (without taking focus from the find box). */
  private showMatch(m: SlideSearchMatch): void {
    if (m.slide !== this.index) this.goTo(m.slide);
    this.findInput.focus({ preventScroll: true });
    if ("highlights" in CSS) highlights().delete("dw-find");
    this.clearNotesHighlight();
    if (m.kind !== "text") {
      if (m.kind === "notes") {
        this.highlightNotes(m.start, m.end);
        this.notesEl.scrollIntoView({ block: "nearest" });
      }
      return;
    }
    const rs = this.rendered.find((r) => r.shape.el === m.owner && r.textEl);
    const body = Array.from(m.owner.getElementsByTagNameNS("*", "txBody"))[0];
    const paraIndex = Array.from(body?.children ?? []).filter((c) => c.localName === "p").indexOf(m.para);
    const p = rs?.textEl?.querySelectorAll<HTMLElement>(":scope > p")[paraIndex];
    if (!p) return;
    const range = rangeInParagraph(p, m.start, m.end);
    if ("highlights" in CSS) highlights().set("dw-find", new Highlight(range));
    p.scrollIntoView({ block: "nearest" });
  }

  private clearNotesHighlight(): void {
    this.notesHighlightEl?.empty();
    this.notesHighlightEl?.parentElement?.removeClass("is-find");
  }

  private highlightNotes(start: number, end: number): void {
    this.clearNotesHighlight();
    const mirror = this.notesHighlightEl;
    mirror.appendText(this.notesEl.value.slice(0, start));
    const mark = mirror.createEl("mark", { text: this.notesEl.value.slice(start, end) });
    mirror.appendText(this.notesEl.value.slice(end) + "\n");
    mirror.parentElement?.addClass("is-find");
    this.notesEl.dispatchEvent(new Event("scroll"));
    const top = mark.offsetTop;
    if (top < this.notesEl.scrollTop || top + mark.offsetHeight > this.notesEl.scrollTop + this.notesEl.clientHeight) {
      this.notesEl.scrollTop = Math.max(0, top - this.notesEl.clientHeight / 2);
      this.notesEl.dispatchEvent(new Event("scroll"));
    }
  }

  private setZoom(zoom: string): void {
    this.zoom = zoom === "fit" ? zoom : String(Math.min(4, Math.max(0.25, Number(zoom))));
    if (this.zoomSelect) this.zoomSelect.value = ZOOMS.includes(this.zoom) ? this.zoom : "";
    this.renderCurrent();
  }

  // ---------------------------------------------------------------------------
  // Toolbar
  // ---------------------------------------------------------------------------

  private buildToolbar(bar: HTMLElement): void {
    // Keep the text selection when clicking toolbar buttons.
    bar.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).closest("button")) e.preventDefault();
    });
    const tb = (this.toolbar = new Toolbar(bar, this.settings.readOnly));
    const btn = (label: string, title: string, run: () => void, cls = "") => tb.button(label, title, run, cls);
    const sep = () => tb.gap();
    const bullets = () => this.toggleList(false);
    const numbering = () => this.toggleList(true);
    const align = (align: Align) => this.updateParas((p) => ({ ...p, align }));
    const lineSpacing = (v: string) => this.updateParas((p) => ({ ...p, lnSpc: { pct: Number(v) } }));
    const shape = (fn: (spPr: Element) => void) => {
      const rs = this.selectedShape();
      if (rs) this.editShape(rs.shape.el, (el) => fn(spPrOf(el)), true);
    };
    const picture = (fn: (el: Element) => void) => () => {
      const rs = this.selectedShape();
      if (rs?.shape.kind === "pic") fn(rs.shape.el);
    };

    // --- The row: everyday controls, as Google Slides' toolbar ---
    tb.icon("undo-2", "Undo (Ctrl+Z)", () => this.undo());
    tb.icon("redo-2", "Redo (Ctrl+Y)", () => this.redo());
    tb.pin();
    sep();
    this.zoomSelect = tb.select("Zoom", "Zoom", ZOOMS.map((z) => [zoomLabel(z), z]), (z) => this.setZoom(z));
    this.zoomSelect.value = this.zoom;
    tb.keep();
    sep();
    tb.icon("square-dashed-text", "Insert text box", () => this.addTextBox());
    const shapesBtn = tb.icon("shapes", "Insert a shape or line", () => this.toggleGallery(shapesBtn));
    tb.icon("table", "Insert table", () => void this.insertTable());
    sep();
    this.fontSelect = tb.select("Font", "Font", FONTS.map((f) => [f, f]), (font) => this.setRunStyle({ font }));
    this.sizeSelect = tb.select("Font size", "Size", FONT_SIZES.map((sz) => [`${sz} pt`, String(sz)]), (sz) => this.setRunStyle({ sz: Number(sz) }));
    tb.pin();
    sep();
    btn("B", "Bold (Ctrl+B)", () => this.toggle("b"), "bold");
    btn("I", "Italic (Ctrl+I)", () => this.toggle("i"), "italic");
    btn("U", "Underline (Ctrl+U)", () => this.toggle("u"), "underline");
    tb.color("A", "Text color", "#000000", (color) => this.setRunStyle({ color }));
    tb.color("HL", "Highlight color", "#FFFF00", (highlight) => this.setRunStyle({ highlight }));
    sep();
    tb.icon("link", "Insert link (Ctrl+K)", () => void this.editLink());
    tb.icon("message-square-plus", "Add comment (Ctrl+Alt+M)", () => this.commentsDialog());
    tb.icon("image", "Insert image", () => pickImage((file) => void this.insertImageFile(file)));
    sep();
    tb.iconSelect("align-left", "Align", ALIGNS.map(([label, a]) => [label, a]), (v) => align(v as Align));
    tb.iconSelect("unfold-vertical", "Line spacing", LINE_SPACINGS.map((v) => [v, v]), lineSpacing);
    sep();
    tb.icon("list", "Bulleted list (Ctrl+Shift+8)", bullets);
    tb.icon("list-ordered", "Numbered list (Ctrl+Shift+7)", numbering);
    tb.icon("outdent", "Decrease indent (Shift+Tab)", () => this.updateParas((p) => ({ ...p, lvl: p.lvl - 1 })));
    tb.icon("indent", "Increase indent (Tab)", () => this.updateParas((p) => ({ ...p, lvl: p.lvl + 1 })));
    sep();
    this.brushBtn = tb.icon("paintbrush", "Paint format (double-click to keep painting)", () => this.startPainting(false));
    this.brushBtn.addEventListener("dblclick", () => this.startPainting(true));
    btn("Tx", "Clear formatting", () => this.clearFormatting());
    // Shape tools, while a shape is selected (Google Slides' fill and border buttons).
    sep();
    tb.color("▇", "Shape fill", "#4472C4", (color) => shape((spPr) => writeFill(spPr, { type: "solid", color })));
    tb.color("▢", "Border color", "#000000", (color) => shape((spPr) => writeLine(spPr, { color })));
    tb.iconSelect("minus", "Border weight", ["0.5", "1", "2", "3", "4.5", "6"].map((v) => [`${v} pt`, v]), (v) => shape((spPr) => writeLine(spPr, { width: Number(v) * EMU_PER_PT })));
    tb.iconSelect("more-horizontal", "Border dash", DASHES, (dash) => shape((spPr) => writeLine(spPr, { dash })));
    tb.context(() => this.selectedShape()?.shape.el.localName === "sp");
    // Picture tools, while a picture is selected.
    sep();
    btn("Crop", "Crop image", () => this.showCrop());
    btn("Replace", "Replace image", picture((el) => this.replaceImage(el)));
    tb.iconSelect("rotate-ccw", "Reset image", [["Reset crop", "crop"], ["Reset colour adjustments", "adjust"]], (v) =>
      picture((el) => this.editShape(el, (p) => (v === "crop" ? setCrop(p, { l: 0, t: 0, r: 0, b: 0 }) : setPictureAdjust(p, NO_ADJUST))))()
    );
    tb.context(() => this.selectedShape()?.shape.kind === "pic" && !this.selectedShape()?.shape.media);
    sep();
    const present = btn("▶ Present", "Present from this slide (Shift+F5). F5 presents from the beginning.", () => this.startPresentation(false), "mod-cta");

    // --- The menus: every command ---
    tb.fileMenu(this, (mn) => menuItem(mn, "Export…", () => void this.exportAs()));
    tb.menu("Edit", (mn) => {
      const has = this.items.length > 0;
      menuItem(mn, "Undo", () => this.undo());
      menuItem(mn, "Redo", () => this.redo());
      mn.addSeparator();
      menuItem(mn, "Cut", () => void this.cutSelection(), { disabled: !has });
      menuItem(mn, "Copy", () => void this.copySelection(), { disabled: !has });
      menuItem(mn, "Paste", () => this.paste(), { disabled: !clipboard });
      // In a text box: the clipboard's plain text; otherwise as Paste.
      menuItem(mn, "Paste without formatting", () => (activeDocument.activeElement?.closest(".pptx-text") ? void runEditCommand("pasteAndMatchStyle") : this.paste()));
      menuItem(mn, "Select all", () => this.setSelection(this.slideItems()));
      menuItem(mn, "Delete", () => this.deleteSelected(), { disabled: !has });
      mn.addSeparator();
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("View", (mn) => {
      menuItem(mn, "Present", () => this.startPresentation(false));
      menuItem(mn, "Present from start", () => this.startPresentation(true));
      mn.addSeparator();
      menuItem(mn, "Grid view", () => this.setGridView(!this.gridView), { checked: this.gridView });
      menuItem(mn, "Format panel", () => this.toggleProps(), { checked: this.showProps });
      subMenu(mn, "Zoom", (s) => {
        for (const z of ZOOMS) menuItem(s, zoomLabel(z), () => this.setZoom(z), { checked: this.zoom === z });
      });
      mn.addSeparator();
      menuItem(mn, "Light", () => this.setPaperTheme("light"), { checked: this.paperTheme === "light" });
      menuItem(mn, "Dark", () => this.setPaperTheme("dark"), { checked: this.paperTheme === "dark" });
    });
    tb.menu("Insert", (mn) => {
      menuItem(mn, "Text box", () => this.addTextBox());
      menuItem(mn, "Image", () => pickImage((file) => void this.insertImageFile(file)));
      menuItem(mn, "Audio or video", () => this.pickMedia());
      menuItem(mn, "Shape", () => this.toggleGallery(shapesBtn));
      menuItem(mn, "Table", () => void this.insertTable());
      subMenu(mn, "Chart", (s) => NEW_CHART_TYPES.forEach(([label, kind]) => menuItem(s, label, () => this.editChart(undefined, kind))));
      menuItem(mn, "Link", () => void this.editLink());
      subMenu(mn, "Field", (s) => this.fillFieldMenu(s));
      menuItem(mn, "Special characters", () => specialCharacters(this.app, (ch) => {
        this.caretOrEnd();
        this.withText(() => insertPlainText(ch), false);
      }));
      menuItem(mn, "Comment", () => this.commentsDialog());
      mn.addSeparator();
      menuItem(mn, "New slide", () => this.mutate((d) => (this.index = d.addSlide(this.index))));
    });
    tb.menu("Format", (mn) => {
      subMenu(mn, "Text", (s) => {
        menuItem(s, "Bold", () => this.toggle("b"));
        menuItem(s, "Italic", () => this.toggle("i"));
        menuItem(s, "Underline", () => this.toggle("u"));
        menuItem(s, "Strikethrough", () => this.toggle("strike"));
        menuItem(s, "Superscript", () => this.toggleBaseline(30));
        menuItem(s, "Subscript", () => this.toggleBaseline(-25));
        menuItem(s, "Remove highlight", () => this.setRunStyle({ highlight: null }));
      });
      subMenu(mn, "Align", (s) => {
        for (const [label, a] of ALIGNS) menuItem(s, label, () => align(a));
      });
      subMenu(mn, "Line spacing", (s) => {
        for (const v of LINE_SPACINGS) menuItem(s, v, () => lineSpacing(v));
      });
      subMenu(mn, "Character spacing", (s) => {
        for (const [label, v] of CHAR_SPACINGS) menuItem(s, label, () => this.setRunStyle({ spc: Number(v) }));
      });
      menuItem(mn, "Bulleted list", bullets);
      menuItem(mn, "Numbered list", numbering);
      mn.addSeparator();
      menuItem(mn, "Clear formatting", () => this.clearFormatting());
    });
    tb.menu("Slide", (mn) => {
      this.slideMenu(this.index, mn);
      mn.addSeparator();
      menuItem(mn, "Theme…", () => this.openTheme());
      menuItem(mn, "Motion…", () => this.openMotion());
    });
    tb.menu("Arrange", (mn) => {
      const a = this.arrangeControl(), has = this.items.length > 0;
      subMenu(mn, "Order", (s) => {
        for (const [label, z] of [["Bring to front", "front"], ["Bring forward", "forward"], ["Send backward", "backward"], ["Send to back", "back"]] as const) menuItem(s, label, () => a.order(z), { disabled: !has });
      });
      subMenu(mn, "Align", (s) => {
        for (const [label, z] of [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Top", "top"], ["Middle", "middle"], ["Bottom", "bottom"]] as const) menuItem(s, label, () => a.align(z), { disabled: !has });
      });
      subMenu(mn, "Distribute", (s) => {
        menuItem(s, "Horizontally", () => a.distribute("horizontal"), { disabled: a.count < 3 });
        menuItem(s, "Vertically", () => a.distribute("vertical"), { disabled: a.count < 3 });
      });
      menuItem(mn, "Group", () => this.groupSelection(), { disabled: this.items.length < 2 });
      menuItem(mn, "Ungroup", () => this.ungroupSelection(), { disabled: !this.items.some((i) => i.localName === "grpSp") });
    });
    tb.menu("Tools", (mn) => {
      menuItem(mn, "Paint format", () => this.startPainting(false));
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("Help", (mn) => menuItem(mn, "Keyboard shortcuts", () => shortcutsDialog(this.app, [...COMMON_SHORTCUTS, ...SLIDES_SHORTCUTS])));

    const end = tb.end();
    this.autosave.status = end.status;
    end.el.prepend(present);
    tb.sync();
  }

  /** The selected shape the user can edit (not one drawn from the layout), if exactly one is selected. */
  private selectedShape(): RenderedShape | null {
    return this.selected && !this.selected.shape.inherited ? this.selected : null;
  }

  private openTheme(): void {
    this.showMotion = false;
    this.showTheme = true;
    this.showProps = true;
    this.openPropsDrawer();
    this.setSelection([]);
  }

  /** Google's Crop: here the format panel's crop fields. */
  private showCrop(): void {
    if (!this.showProps || this.contentEl.hasClass("is-narrow")) this.toggleProps();
    this.propsEl.querySelector<HTMLInputElement>('[aria-label="Crop left %"]')?.focus();
  }

  private replaceImage(el: Element): void {
    pickImage((file) => void file.arrayBuffer().then((buf) => {
      const ext = file.name.includes(".") ? file.name.split(".").pop()!.toLowerCase() : "png";
      this.mutate((doc) => replacePicture(doc, this.editPart ?? doc.slides()[this.index].path, el, new Uint8Array(buf), ext), "slide");
      void this.loadMedia().then(() => this.renderCurrent());
    }));
  }

  /** The format panel shows when docked, or while it's open as a drawer (narrow panes). */
  private showPropsEl(): void {
    this.propsEl.toggle(this.showProps || this.propsEl.hasClass("is-drawer"));
  }

  /** In a narrow pane, panel content (theme, motion, crop) only shows in the drawer: open it. */
  private openPropsDrawer(): void {
    if (this.contentEl.hasClass("is-narrow") && !this.closeDrawer) this.toggleProps();
  }

  private toggleProps(): void {
    if (this.contentEl.hasClass("is-narrow")) {
      if (this.closeDrawer) return this.closeDrawer();
      this.closeDrawer = openDrawer(this.propsEl, {
        root: this.contentEl,
        onClose: () => {
          this.closeDrawer = null;
          this.propsEl.toggle(this.showProps);
        },
      });
      return;
    }
    this.showProps = !this.showProps;
    this.app.saveLocalStorage(PROPS_KEY, this.showProps ? "shown" : "hidden");
    this.propsEl.toggle(this.showProps);
  }

  setPaperTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    this.contentEl.toggleClass("pptx-theme-dark", theme === "dark");
    this.contentEl.toggleClass("pptx-theme-light", theme === "light");
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------


}

/** Characters of `el`'s text before the caret (0 when the caret is elsewhere). */
function caretOffset(el: HTMLElement): number {
  const sel = el.ownerDocument.getSelection();
  if (!sel?.rangeCount || !el.contains(sel.anchorNode)) return 0;
  const range = el.ownerDocument.createRange();
  range.selectNodeContents(el);
  range.setEnd(sel.anchorNode!, sel.anchorOffset);
  return range.toString().length;
}

/** Put the caret `offset` characters into `el`'s text (at the end when it's shorter). */
function placeCaret(el: HTMLElement, offset: number): void {
  const sel = el.ownerDocument.getSelection();
  if (!sel) return;
  const walker = el.ownerDocument.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let left = offset;
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    if (left <= node.length) {
      sel.collapse(node, left);
      return;
    }
    left -= node.length;
  }
  const range = el.ownerDocument.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Google's link box: a web address, or a slide of this deck (next, previous, first, last or a specific one).
 * Resolves the new link, null to remove it, or undefined when cancelled.
 */
function linkDialog(app: App, current: string, slides: string[], titles: string[]): Promise<string | null | undefined> {
  return new Promise((resolve) => {
    let result: string | null | undefined;
    const modal = new Modal(app);
    modal.titleEl.setText("Link");
    const done = (link: string | null) => {
      result = link;
      modal.close();
    };
    if (current) modal.contentEl.createEl("p", { cls: "pptx-props-hint", text: `Current link: ${linkLabel(current, slides)}` });
    const url = modal.contentEl.createEl("input", { type: "text", cls: "pptx-link-url", attr: { "aria-label": "Link address", placeholder: "Paste a link" } });
    if (current && !current.startsWith("#")) url.value = current;
    url.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      if (url.value.trim()) done(url.value.trim());
    });
    new Setting(modal.contentEl).addButton((b) => b.setButtonText("Apply").setCta().onClick(() => url.value.trim() && done(url.value.trim())));
    modal.contentEl.createEl("h4", { text: "Slides in this presentation" });
    const list = modal.contentEl.createDiv({ cls: "pptx-link-slides" });
    const option = (text: string, link: string) => {
      const b = list.createEl("button", { text });
      b.toggleClass("is-active", link === current);
      b.addEventListener("click", () => done(link));
    };
    for (const jump of JUMPS) option(linkLabel(`#jump:${jump}`, slides), `#jump:${jump}`);
    slides.forEach((path, i) => option(titles[i] ?? `Slide ${i + 1}`, `#slide:${path}`));
    if (current) new Setting(modal.contentEl).addButton((b) => b.setButtonText("Remove link").onClick(() => done(null)));
    modal.onClose = () => resolve(result);
    modal.open();
    window.setTimeout(() => url.focus());
  });
}
