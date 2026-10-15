import { readDrawing, newDrawing, writeDrawing } from "../../domain/docx/drawings";
import { DrawingView } from "./docx-drawing-view";
import { ChartView } from "./docx-chart-view";
import { chartAt, chartRun, docxChartColors, editedChartRun } from "../../domain/docx/chart";
import { chartPalette } from "../../domain/pptx/chart-format";
import { chartDialog, NEW_CHART_TYPES, newChartFormat } from "../common/chart-form";
import type { ChartKind } from "../../domain/pptx/chart-format";
import { drawingEditor } from "./drawing-editor";
import { App, FileView, Menu, Modal, Notice, Scope, Setting, setIcon, TFile, WorkspaceLeaf } from "obsidian";
import { Command, EditorState, NodeSelection, Plugin, TextSelection } from "prosemirror-state";
import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { deleteSelection, selectAll, toggleMark } from "prosemirror-commands";
import { closeHistory, redo, undo } from "prosemirror-history";
import { DocxDocument, HFType } from "../../infrastructure/docx/docx-document";
import { Origins, readDocx, startsWordPage } from "../../domain/docx/docx-reader";
import { writeBody } from "../../domain/docx/docx-writer";
import { DocxStyles } from "../../domain/docx/docx-styles";
import { docxSchema as schema, Align } from "../../domain/docx/schema";
import { Autosave, reloadOnOutsideChange } from "../common/autosave";
import { keepVersion, versions } from "../common/version-history";
import { versionHistoryDialog } from "../common/version-dialog";
import { LinkedFrom } from "../common/backlinks";
import { findHeading, headingLinkText, headings } from "../../domain/docx/headings";
import { copyLink } from "../common/copy-link";
import { escapeToEditor, focusRegions, returnFocusAfterOverlays } from "../common/keyboard";

import { FONTS, menuItem, pickColor, promptText, showValue, subMenu, Toolbar, shieldEditingKeys, addFileItems } from "../common/toolbar";
import { openDrawer, watchNarrow } from "../common/narrow";
import { pickImage } from "../common/images";
import { electronSpeller, runEditCommand, Speller } from "../common/spelling";
import { specialCharacters } from "../common/special-chars";
import { COMMON_SHORTCUTS, DOCS_SHORTCUTS, registerShortcutsKey, shortcutsDialog } from "../common/shortcuts";
import { recrop, Wrap } from "../../domain/docx/drawing";
import { readWatermark, setWatermark, TextWatermark, Watermark } from "../../domain/docx/watermark";
import { directFormat, updateStyleToMatch } from "../../domain/docx/style-edit";
import { fontStack } from "../../core/utils/fonts";
import { IMAGE_MIME } from "../../core/utils/images";
import { cmd, commentRanges, docxPlugins, EditorContext, PaintFormat, footnoteIds, linkHref, pageNumberNode, refreshText, styleRenderedParagraphs } from "./docx-editor";
import { loadEmbeddedFonts } from "../pptx/embedded-fonts";
import { pageAt, pagesPlugin, printPages, refreshSectionWidths, repaginate, sectionWidthsPlugin } from "./docx-pages";
import { rewrap, wrapBothSidesPlugin } from "./docx-wrap";
import { latexToOmml, mathSpec } from "../../domain/docx/math";
import { canExportPdf, renderPdf } from "../common/pdf";
import { ExportFormat, exportDialog, writeExport } from "../common/export-dialog";
import { exportDocxToMarkdown, MARKDOWN_FORMAT } from "../common/export-markdown";
import { acceptRevisions, groupRevisions, rejectRevisions, revisions } from "../../domain/docx/revisions";
import { countWords, findRanges, findCommentMatches, CommentMatch, Range, replaceAllCmd, replaceRangeCmd, setFind } from "./docx-find";
import { showFindPreview } from "../common/find-preview";
import { findBar } from "../common/find-bar";
import { pageSetup, PageSetup, sectionSetups } from "../../domain/docx/pages";
import { attrsOf, ImageAttrs, ParagraphAttrs } from "../../domain/docx/schema";
import { DocxImageView, WRAPS } from "./docx-image-view";
import { TextBoxView } from "./docx-textbox-view";
import { retainOrigins, suggestTransaction } from "../../domain/docx/suggestions";
import { DocxRuler, outlineEntries } from "./docx-ruler";

import { DEFAULT_SETTINGS, DocWorksSettings, paperFor } from "../common/settings";
export { DEFAULT_SETTINGS };
export type { DocWorksSettings };

export const VIEW_TYPE_DOCX = "doc-works-docx-view";

const SAVE_DELAY_MS = 800;
const WORD_COUNT_KEY = "doc-works-word-count";
const SHADES: [string, string][] = [["Blue", "#EBF3FB"], ["Gray", "#F2F2F2"], ["Yellow", "#FFF8E1"], ["Green", "#E8F5E9"]];
const ALIGNS: [string, Align][] = [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Justify", "justify"]];
const LINE_SPACINGS: [string, string][] = [["1.0", "1"], ["1.15", "1.15"], ["1.5", "1.5"], ["2.0", "2"]];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const STYLES: [string, string][] = [
  ["Normal", "Normal"],
  ["Title", "Title"],
  ["Heading 1", "Heading1"],
  ["Heading 2", "Heading2"],
  ["Heading 3", "Heading3"],
  ["Heading 4", "Heading4"],
  ["Quote", "Quote"],
];

/** Equation palette rows: label, LaTeX to insert. */
const EQUATION_PALETTE: [string, string][][] = [
  [["a/b", "\\frac{}{}"], ["√", "\\sqrt{}"], ["ⁿ√", "\\sqrt[]{}"], ["xⁿ", "^{}"], ["xₙ", "_{}"], ["∑", "\\sum_{}^{}{}"], ["∫", "\\int_{}^{}{}"], ["lim", "\\lim_{}{}"], ["( )", "\\left(\\right)"], ["[ ]", "\\left[\\right]"], ["x̂", "\\hat{}"], ["x̄", "\\overline{}"], ["vec", "\\vec{}"], ["▦", "\\begin{matrix}{}&\\\\&\\end{matrix}"]],
  [["α", "\\alpha "], ["β", "\\beta "], ["γ", "\\gamma "], ["δ", "\\delta "], ["ε", "\\varepsilon "], ["θ", "\\theta "], ["λ", "\\lambda "], ["μ", "\\mu "], ["π", "\\pi "], ["σ", "\\sigma "], ["φ", "\\varphi "], ["ω", "\\omega "], ["Δ", "\\Delta "], ["Σ", "\\Sigma "], ["Ω", "\\Omega "]],
  [["±", "\\pm "], ["×", "\\times "], ["÷", "\\div "], ["·", "\\cdot "], ["∞", "\\infty "], ["∂", "\\partial "], ["∇", "\\nabla "], ["≤", "\\leq "], ["≥", "\\geq "], ["≠", "\\neq "], ["≈", "\\approx "], ["≡", "\\equiv "], ["∈", "\\in "], ["⊂", "\\subset "], ["∪", "\\cup "], ["∩", "\\cap "], ["∀", "\\forall "], ["∃", "\\exists "], ["→", "\\to "], ["⇒", "\\Rightarrow "], ["⇔", "\\Leftrightarrow "]],
];

const PAGE_VIEW_KEY = "doc-works-page-view";
const ZOOM_KEY = "doc-works-docx-zoom";
const ZOOMS = ["fit", "0.5", "0.75", "0.9", "1", "1.25", "1.5", "2"];
const zoomLabel = (z: string) => (z === "fit" ? "Fit" : `${Math.round(Number(z) * 100)}%`);
const REVIEW_KEY = "doc-works-review-panel";
type HF = "header" | "footer";
/** A header or footer part: "header:default", "footer:first"... */
type HFPart = `${HF}:${HFType}`;
const HF_PARTS: HFPart[] = ["header:default", "header:first", "header:even", "footer:default", "footer:first", "footer:even"];
const splitPart = (part: HFPart) => part.split(":") as [HF, HFType];

const dataUrl = (file: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error("Could not read file"));
    reader.readAsDataURL(file);
  });

const naturalSize = (src: string) =>
  new Promise<{ w: number; h: number }>((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 300, h: img.naturalHeight || 200 });
    img.onerror = () => resolve({ w: 300, h: 200 });
    img.src = src;
  });

export class DocxView extends FileView {
  /** Stops giving focus back to the document after dialogs close. */
  private stopFocusReturn: (() => void) | null = null;
  allowNoFile = false;
  paperTheme: "light" | "dark" = "light";

  private docx: DocxDocument | null = null;
  private origins: Origins = new Map();
  private editor: EditorView | null = null;
  private pageWidth = 624;
  private autosave = new Autosave({
    delay: SAVE_DELAY_MS,
    readOnly: () => this.settings.readOnly,
    canSave: () => !!this.editor && !!this.docx && !!this.file,
    name: () => this.file?.name ?? "",
    write: async () => {
      const file = this.file!, docx = this.docx!;
      writeBody(docx.body, this.editor!.state.doc, this.origins, docx);
      const bytes = await docx.save();
      await keepVersion(this.app, file);
      await this.app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    },
  });
  private toolbar: Toolbar | null = null;
  private wrapSelect: HTMLSelectElement | null = null;
  /** The last cell colour picked, where the custom picker starts. */
  private lastShade = SHADES[0][1];
  private narrowWatch: ReturnType<typeof watchNarrow> | null = null;
  private closeDrawer: (() => void) | null = null;
  private ctx: EditorContext | null = null;
  private setup: PageSetup | null = null;
  /** The body's own page setup (`setup` is as wide as the widest section's page). */
  private finalSetup: PageSetup | null = null;
  private pageView = this.app.loadLocalStorage(PAGE_VIEW_KEY) !== "off";
  private zoom: string = ZOOMS.includes(this.app.loadLocalStorage(ZOOM_KEY) as string) ? (this.app.loadLocalStorage(ZOOM_KEY) as string) : "fit";
  private zoomSelect: HTMLSelectElement | null = null;
  private hostEl: HTMLElement | null = null;
  private linked = new LinkedFrom(this, "docx", () => this.renderOutline());
  private wordsEl: HTMLElement | null = null;
  private wordsTimer: number | null = null;
  private rulerFrame = 0;
  private findEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  /** Header and footer content, for drawing them in the page view. */
  private hf = new Map<HFPart, PMNode>();
  private styleSelect: HTMLSelectElement | null = null;
  private fontSelect: HTMLSelectElement | null = null;
  private sizeSelect: HTMLSelectElement | null = null;
  /** rId -> data URL for pictures in the header and footer parts. */
  /** Pictures of each header / footer part, by part path. */
  private hfUrls = new Map<string, Map<string, string>>();
  /** Every header / footer part any section shows, read for drawing on the pages, by part path. */
  private hfByPath = new Map<string, PMNode>();
  private reviewEl: HTMLElement | null = null;
  private reviewOpen = this.app.loadLocalStorage(REVIEW_KEY) === "on";
  private reviewTimer: number | null = null;
  private suggesting = false;
  /** The key behind the transaction being dispatched (Delete moves the cursor past what it strikes through). */
  private lastKey: string | null = null;
  private ruler: DocxRuler | null = null;
  private rulerViewport: HTMLElement | null = null;
  private outlineEl: HTMLElement | null = null;
  private outlineOpen = this.app.loadLocalStorage("doc-works-outline") === "on";

  constructor(
    leaf: WorkspaceLeaf,
    private settings: DocWorksSettings = { ...DEFAULT_SETTINGS },
    private saveSettings: () => Promise<void> = async () => {}
  ) {
    super(leaf);
    this.scope = new Scope(this.app.scope);
    focusRegions(this.scope, this.contentEl, () => [this.contentEl.querySelector<HTMLElement>(".dw-menubar"), this.contentEl.querySelector<HTMLElement>(".dw-toolbar"), this.editor?.dom, this.outlineEl, this.reviewEl, this.findEl]);
    this.scope.register(["Mod"], "f", () => {
      this.openFind();
      return false;
    });
    registerShortcutsKey(this.scope, () => shortcutsDialog(this.app, [...COMMON_SHORTCUTS, ...DOCS_SHORTCUTS]));
    this.scope.register(["Mod"], "h", () => {
      this.openFind();
      return false;
    });
    // Obsidian binds these to its markdown editor (bold, italic, link...) or the user to other commands: inside the
    // document they belong to the document's editor, which handles them itself.
    shieldEditingKeys(this.scope);
    reloadOnOutsideChange(this, this.autosave, async () => {
      const top = this.hostEl?.scrollTop ?? 0;
      if (this.file) await this.onLoadFile(this.file);
      if (this.hostEl) this.hostEl.scrollTop = top;
    });
  }

  onPaneMenu(menu: Menu, source: string): void {
    super.onPaneMenu(menu, source);
    addFileItems(this, menu, source);
    if (versions) menu.addItem((i) => i.setTitle("Version history…").setIcon("history").setSection("action").onClick(() => void versionHistoryDialog(this.app, this, this.settings.readOnly)));
    menu.addItem((i) => i.setTitle("Export…").setIcon("file-output").setSection("action").onClick(() => void this.exportAs()));
  }

  /** File › Download, as Google Docs: PDF or Markdown, through the shared export dialog. */
  async exportAs(initial?: "pdf" | "md"): Promise<void> {
    const file = this.file;
    if (!file) return;
    const formats: ExportFormat[] = [];
    if (canExportPdf()) formats.push({ id: "pdf", label: "PDF document (.pdf)", ext: "pdf", outside: true });
    formats.push(MARKDOWN_FORMAT);
    const choice = await exportDialog(this.app, file, formats, initial);
    if (!choice) return;
    if (choice.format === "pdf") {
      const data = await this.pdfBytes();
      if (data) await writeExport(this.app, choice.target, data, choice.open);
    } else if ("vault" in choice.target) {
      await exportDocxToMarkdown(this.app, file, this.docx && this.editor ? { document: this.docx, content: this.editor.state.doc } : undefined, choice.target.vault, choice.open);
    }
  }

  getViewType(): string {
    return VIEW_TYPE_DOCX;
  }

  getIcon(): string {
    return "file-text";
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Docx";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "docx";
  }

  async onLoadFile(file: TFile): Promise<void> {
    const load = ++this.loads;
    this.stopFocusReturn?.();
    this.stopFocusReturn = returnFocusAfterOverlays(() => this.app.workspace.getActiveViewOfType(DocxView) === this, () => this.editor?.focus());
    this.contentEl.empty();
    this.contentEl.addClass("dw-root");
    this.contentEl.toggleClass("dw-read-only", this.settings.readOnly);
    const status = this.contentEl.createDiv({ cls: "dw-loading", text: "Loading document..." });

    let docx: DocxDocument;
    let urls: Map<string, string>;
    try {
      docx = await DocxDocument.load(await this.app.vault.readBinary(file));
      urls = await docx.imageUrls();
      for (const path of hfPaths(docx)) this.hfUrls.set(path, await docx.imageUrls(path));
    } catch (err) {
      status.setText(`Could not open this .docx file: ${(err as Error).message}`);
      console.error(err);
      return;
    }
    // Opened again meanwhile (a setting change while loading): only the newest load draws.
    if (load !== this.loads) return;
    status.remove();
    this.docx = docx;

    const { doc, origins } = readDocx(docx.body, {
      imageSrc: (id) => urls.get(id) ?? null,
      hyperlinkHref: (id) => docx.hyperlinkHref(id),
    });
    this.origins = origins;
    this.readAllHF();
    await this.loadWatermark();
    if (load !== this.loads) return;

    const ctx: EditorContext = {
      styles: new DocxStyles(docx.stylesXml(), docx.numberingXml(), docx.themeXml()),
      // A new list may create numbering.xml, so refresh the display styles afterwards.
      ensureList: (kind) => {
        const numId = docx.ensureList(kind);
        ctx.styles = new DocxStyles(docx.stylesXml(), docx.numberingXml(), docx.themeXml());
        return numId;
      },
      listKind: (numId) => {
        const fmt = docx.listFormat(numId, 0);
        return fmt ? (fmt === "bullet" || fmt === "none" ? "bullet" : "ordered") : null;
      },
    };

    const toolbar = this.contentEl.createDiv({ cls: "dw-toolbar" });
    const body = this.contentEl.createDiv({ cls: "docx-body" });
    // Regions are named through hidden labels: Obsidian shows aria-label as a hover tooltip.
    const outlineName = body.createDiv({ cls: "dw-sr-only", text: "Document outline", attr: { id: `docx-outline-name-${Math.random().toString(36).slice(2)}` } });
    this.outlineEl = body.createDiv({ cls: "docx-outline", attr: { "aria-labelledby": outlineName.id } });
    this.outlineEl.toggle(this.outlineOpen);
    const pane = body.createDiv({ cls: "docx-editor-pane" });
    const rulerViewport = this.rulerViewport = pane.createDiv({ cls: "docx-ruler-viewport" });
    const host = (this.hostEl = pane.createDiv({ cls: "docx-editor-host" }));
    this.ruler = new DocxRuler(rulerViewport, () => this.editor, () => { this.applyPageSize(); if (this.editor) repaginate(this.editor); });
    host.addEventListener("scroll", () => this.ruler?.align());
    this.reviewEl = body.createDiv({ cls: "docx-review" });
    for (const panel of [this.outlineEl, this.reviewEl]) escapeToEditor(panel, () => this.editor?.focus());
    this.reviewEl.toggle(this.reviewOpen);
    this.applyPageSize();
    host.toggleClass("is-paged", this.pageView);
    this.ctx = ctx;

    // Double-click a page's header or footer to edit it, as in Word.
    host.addEventListener("dblclick", (e) => {
      if (!this.settings.readOnly && (e.target as HTMLElement).closest(".docx-page-first, .docx-page-head, .docx-page-foot")) this.headerFooterDialog();
    });
    // Click an equation to edit it, as in Google Docs.
    host.addEventListener("click", (e) => {
      const el = (e.target as HTMLElement).closest(".docx-math")?.parentElement;
      const editor = this.editor;
      if (!el || !editor || !editor.dom.contains(el) || this.settings.readOnly) return;
      const at = editor.posAtDOM(el, 0);
      const pos = [at, at - 1].find((p) => p >= 0 && cmd.equationAt(editor.state, p) !== null);
      if (pos !== undefined) this.equationDialog(pos);
    });
    this.editor = new EditorView(host, {
      attributes: { role: "textbox", "aria-labelledby": host.createDiv({ cls: "dw-sr-only", text: "Document content", attr: { id: `docx-content-name-${Math.random().toString(36).slice(2)}` } }).id, "aria-multiline": "true" },
      state: EditorState.create({ doc, plugins: [...docxPlugins(ctx, { "Mod-k": () => (void this.editLink(), true), "Mod-Alt-m": () => (void this.addComment(), true), "Mod-Shift-8": cmd.toggleList(ctx, "bullet"), "Mod-Shift-7": cmd.toggleList(ctx, "ordered") }), this.pagesPlugin(), wrapBothSidesPlugin(), sectionWidthsPlugin(() => (this.pageView ? this.setup : null), () => this.finalSetup, () => this.settings.pageSize), readOnlyPlugin(() => this.settings.readOnly)] }),
      editable: () => !this.settings.readOnly,
      nodeViews: {
        image: (node, view, getPos) =>
          new DocxImageView(node, view, getPos, {
            replace: (pos) => pickImage((f) => void this.replaceImage(pos, f)),
            crop: (current) => cropDialog(this.app, current),
            editAlt: (pos, current) =>
              void promptText(this.app, "Image description (alt text)", current).then((alt) => {
                const node = alt === null ? null : view.state.doc.nodeAt(pos);
                if (node) view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, alt }));
              }),
          }),
        // Text boxes are edited in place; other run content is drawn as usual.
        opaque: (node, view, getPos) => {
          const xml=node.attrs.xml as string;
          if (chartAt(docx, xml)) return new ChartView(node, view, getPos, (x) => chartAt(docx, x), (pos) => this.editChart(pos));
          return readDrawing(xml) && !/txBox="1"/.test(xml) ? new DrawingView(node,view,getPos,pos=>this.drawingDialog(pos)) : new TextBoxView(node, view, getPos, ctx, { imageSrc: (id) => urls.get(id) ?? null, hyperlinkHref: (id) => docx.hyperlinkHref(id) }, docx);
        },
      },
      dispatchTransaction: (tr) => {
        if (!this.editor) return;
        if (this.suggesting) tr = suggestTransaction(tr, this.origins, docx,
          { imageSrc: id => this.imageSrc(id) ?? urls.get(id) ?? null, hyperlinkHref: id => docx.hyperlinkHref(id) }, this.settings.author, { forward: this.lastKey === "Delete" });
        this.lastKey = null;
        const applied = this.editor.state.applyTransaction(tr);
        this.editor.updateState(applied.state);
        const changed = applied.transactions.some(t => t.docChanged);
        if (changed) this.autosave.markDirty();
        if (changed || tr.selectionSet) {
          this.scheduleWordCount();
          this.updateFormatControls();
          this.scheduleRuler();
          this.renderOutline();
        }
        if (changed) this.scheduleReview();
      },
      handleKeyDown: (_view, event) => {
        this.lastKey = event.key;
        if (event.key === "Escape" && this.painter) {
          this.stopPainting();
          return true;
        }
        return false;
      },
      handleDOMEvents: {
        // Paint format: the selection made with the mouse takes the copied formatting.
        mouseup: (view) => {
          if (!this.painter) return false;
          // After the mouse selection settles.
          window.setTimeout(() => this.paintSelection(view), 0);
          return false;
        },
        click: (_view, event) => {
          // Ctrl / Cmd + click follows a link.
          const a = (event.target as HTMLElement).closest<HTMLAnchorElement>("a.docx-link");
          if (!a || !(event.ctrlKey || event.metaKey)) return false;
          event.preventDefault();
          if (/^https?:|^mailto:/.test(a.getAttribute("href") ?? "")) window.open(a.getAttribute("href")!);
          return true;
        },
        blur: () => {
          void this.autosave.flush();
          return false;
        },
      },
      handlePaste: (_view, event) => this.takeImage(event.clipboardData?.files),
      handleDrop: (view, event) => {
        const pos = view.posAtCoords({ left: event.clientX, top: event.clientY });
        const files = event.dataTransfer?.files;
        if (pos && files?.length) view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(pos.pos))));
        return this.takeImage(files);
      },
    });
    // On the editor's element (not a ProseMirror handler): picture node views would stop the event from reaching one.
    this.editor.dom.addEventListener("contextmenu", (e) => !this.settings.readOnly && this.contextMenu(e));

    this.buildToolbar(toolbar, ctx);
    this.setPaperTheme(paperFor(this.settings, "light"));
    window.addEventListener("blur", this.onWindowBlur);
    this.resizeObserver = new ResizeObserver(() => {
      this.fitDocument();
    });
    this.resizeObserver.observe(host);
    this.narrowWatch = watchNarrow(this.contentEl, (narrow) => {
      if (!narrow) this.closeDrawer?.();
      this.toolbar?.fit();
      this.fitDocument();
    });
    this.updateWordCount();
    this.renderReview();
    this.ruler.update();
    this.fitDocument();
    this.renderOutline();
    void this.loadFonts();
    this.loaded = load;
    void this.linked.refresh();
    this.goToPlace();
  }

  /** A link's place in the file (#3, #Sheet!B4, #Heading), kept until the file has loaded. */
  private place: string | null = null;
  /** The onLoadFile call that last finished (equal to `loads` once the newest one has). */
  private loaded = 0;

  setEphemeralState(state: unknown): void {
    super.setEphemeralState(state);
    const subpath = (state as { subpath?: string } | null)?.subpath;
    if (!subpath) return;
    this.place = subpath;
    this.goToPlace();
  }

  /** Go to the link's place, once the file is in. */
  private goToPlace(): void {
    if (!this.place || this.loaded !== this.loads) return;
    const place = this.place;
    this.place = null;
    const heading = this.editor && findHeading(this.editor.state.doc, place);
    if (heading) this.showHeading(heading.pos + 1);
  }

  /** A link to the heading the cursor is under ([[report.docx#Results]]) for a note. */
  private copyHeadingLink(): void {
    const from = this.editor?.state.selection.from ?? 0;
    const heading = this.editor && [...headings(this.editor.state.doc)].reverse().find((h) => h.pos < from);
    if (!heading) return void new Notice("No heading above this point to link to.");
    copyLink(this.app, this.file, `#${headingLinkText(heading.text)}`, "heading");
  }

  /** Put the cursor on the heading at `pos` and scroll it to the top. */
  private showHeading(pos: number): void {
    const editor = this.editor;
    if (!editor) return;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)));
    editor.focus();
    const heading = editor.nodeDOM(pos - 1);
    if (heading instanceof HTMLElement && this.hostEl) {
      this.hostEl.scrollTop = Math.max(0, this.hostEl.scrollTop + heading.getBoundingClientRect().top - this.hostEl.getBoundingClientRect().top - 16);
    }
  }

  /** Removes the document's embedded fonts from the page. */
  private unloadFonts: (() => void) | null = null;

  /** Load the fonts the document carries, then lay it out again with them (after opening, so that isn't held up). */
  private async loadFonts(): Promise<void> {
    const docx = this.docx;
    if (!docx) return;
    const { loaded, unload } = await loadEmbeddedFonts(await docx.embeddedFonts(), (data) => data);
    if (this.docx !== docx) return unload(); // another file was opened meanwhile
    this.unloadFonts = unload;
    if (loaded && this.editor) {
      refreshText(this.editor);
      repaginate(this.editor);
    }
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    this.unloadFonts?.();
    this.unloadFonts = null;
    window.removeEventListener("blur", this.onWindowBlur);
    this.resizeObserver?.disconnect();
    this.narrowWatch?.();
    this.closeDrawer?.();
    this.closeFind();
    for (const t of [this.wordsTimer, this.reviewTimer]) if (t) window.clearTimeout(t);
    await this.autosave.close();
    this.stopFocusReturn?.();
    this.stopFocusReturn = null;
    this.editor?.destroy();
    this.ruler?.destroy();
    this.ruler = null;
    this.editor = null;
    this.docx = null;
  }

  /** Page geometry from the document, or Letter when the page size setting says so; re-lays out open pages. */
  applyPageSize(): void {
    if (!this.docx || !this.hostEl) return;
    const finalSectPr = this.editor?.state.doc.attrs.sectPr as string | null | undefined;
    const body = finalSectPr ? new DOMParser().parseFromString(`<body>${finalSectPr}</body>`, "application/xml").documentElement : this.docx.body;
    const st = (this.finalSetup = pageSetup(body, this.settings.pageSize));
    // A section with its own page size (landscape pages): the editor is as wide as the widest page.
    const widest = this.editor ? Math.max(...sectionSetups(this.editor.state.doc, st, this.settings.pageSize).map((x) => x.width), st.width) : st.width;
    this.setup = { ...st, width: widest };
    this.pageWidth = widest - st.left - st.right;
    this.hostEl.setCssProps({
      "--docx-page-width": `${this.pageWidth}px`,
      "--docx-m-top": `${st.top}px`,
      "--docx-m-right": `${st.right}px`,
      "--docx-m-bottom": `${st.bottom}px`,
      "--docx-m-left": `${st.left}px`,
      "--docx-hf-header": `${st.header}px`,
      "--docx-hf-footer": `${st.footer}px`,
    });
    if (this.editor) {
      refreshSectionWidths(this.editor);
      repaginate(this.editor);
    }
    this.fitDocument();
  }

  /** View › Zoom: "fit" shrinks a page wider than the pane to fit it, as before; a number is that zoom, remembered. */
  private setZoom(zoom: string): void {
    this.zoom = ZOOMS.includes(zoom) ? zoom : "fit";
    this.app.saveLocalStorage(ZOOM_KEY, this.zoom);
    if (this.zoomSelect) this.zoomSelect.value = this.zoom;
    this.fitDocument();
  }

  private fitDocument(): void {
    const editor = this.editor, host = this.hostEl, setup = this.setup;
    if (!editor || !host || !setup) return;
    const padding = getComputedStyle(host);
    const available = host.clientWidth - (Number.parseFloat(padding.paddingLeft) || 0) - (Number.parseFloat(padding.paddingRight) || 0);
    const fit = this.pageView && available > 0 ? Math.min(1, available / setup.width) : 1;
    const zoom = this.zoom === "fit" ? fit : Number(this.zoom);
    const value = String(zoom);
    const changed = editor.dom.style.zoom !== value;
    editor.dom.style.zoom = value;
    if (this.ruler) this.ruler.dom.style.zoom = value;
    this.ruler?.align();
    if (changed) repaginate(editor);
    rewrap(editor);
  }

  /** Force an immediate save, e.g. before Obsidian quits. */
  async requestSave(): Promise<void> {
    await this.autosave.flush();
  }

  private onWindowBlur = () => void this.autosave.flush();

  setPaperTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    this.contentEl.toggleClass("docx-theme-light", theme === "light");
    this.contentEl.toggleClass("docx-theme-dark", theme === "dark");
  }

  // ---------------------------------------------------------------------------
  // Images
  // ---------------------------------------------------------------------------

  private takeImage(files: FileList | undefined | null): boolean {
    const image = Array.from(files ?? []).find((f) => f.type.startsWith("image/"));
    if (!image) return false;
    void this.insertImage(image);
    return true;
  }

  /** Add the file to the package; returns attrs for an image node sized to fit the page. */
  private async imageAttrs(file: File) {
    const ext = file.name.includes(".") ? file.name.split(".").pop()! : file.type.split("/")[1] ?? "png";
    const rId = this.docx!.addImage(new Uint8Array(await file.arrayBuffer()), ext);
    const src = await dataUrl(file);
    const { w, h } = await naturalSize(src);
    const width = Math.min(w, this.pageWidth);
    return { rId, src, width, height: Math.round((width / w) * h) };
  }

  private async insertImage(file: File): Promise<void> {
    if (!this.editor || !this.docx) return;
    const attrs = await this.imageAttrs(file);
    const node = schema.nodes.image.create({ ...attrs, alt: file.name });
    this.editor.dispatch(this.editor.state.tr.replaceSelectionWith(node).scrollIntoView());
    this.editor.focus();
  }

  private async replaceImage(pos: number, file: File): Promise<void> {
    if (!this.editor || !this.docx) return;
    const attrs = await this.imageAttrs(file);
    const node = this.editor.state.doc.nodeAt(pos);
    if (node?.type === schema.nodes.image) this.editor.dispatch(this.editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs }));
  }

  // ---------------------------------------------------------------------------
  // Toolbar
  // ---------------------------------------------------------------------------

  private pageViewBtn: HTMLButtonElement | null = null;
  /** Counts onLoadFile calls, so an older load still reading the file gives way. */
  private loads = 0;

  private buildToolbar(bar: HTMLElement, ctx: EditorContext): void {
    // Buttons must not steal focus/selection from the editor.
    bar.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).closest("button")) e.preventDefault();
    });
    const run = (command: Command) => {
      if (!this.editor) return;
      command(this.editor.state, this.editor.dispatch);
      this.editor.focus();
    };
    const tb = (this.toolbar = new Toolbar(bar, this.settings.readOnly));
    const action = (label: string, title: string, onClick: () => void, cls = "") => tb.button(label, title, onClick, cls);
    const btn = (label: string, title: string, command: Command, cls = "") => action(label, title, () => run(command), cls);
    const sep = () => tb.gap();
    const m = schema.marks;
    const image = (fn: (pos: number) => void) => () => {
      const pos = this.selectedImage();
      if (pos !== null) fn(pos);
    };

    // --- The row: everyday controls, as Google Docs' toolbar ---
    tb.icon("undo-2", "Undo (Ctrl+Z)", () => run(undo));
    tb.icon("redo-2", "Redo (Ctrl+Y)", () => run(redo));
    tb.pin();
    sep();
    this.zoomSelect = tb.select("Zoom", "Zoom", ZOOMS.map((z) => [zoomLabel(z), z]), (z) => this.setZoom(z));
    this.zoomSelect.value = this.zoom;
    this.pageViewBtn = tb.icon("file-text", "Page view", () => this.togglePageView());
    this.pageViewBtn.toggleClass("is-active", this.pageView);
    tb.keep();
    sep();
    this.styleSelect = tb.select("Paragraph style", "Style", [...STYLES, ["Update style to match", "update"]], (v) => {
      if (v === "update") return this.updateStyle();
      run(cmd.setStyle(v === "Normal" ? null : v));
    });
    this.fontSelect = tb.select("Font", "Font", FONTS.map((f) => [f, f]), (font) => run(cmd.setMark(m.font, { font })));
    this.sizeSelect = tb.select("Font size", "Size", FONT_SIZES.map((s) => [`${s} pt`, String(s)]), (v) => run(cmd.setMark(m.size, { pt: Number(v) })));
    tb.pin();
    sep();
    btn("B", "Bold (Ctrl+B)", toggleMark(m.bold), "bold");
    btn("I", "Italic (Ctrl+I)", toggleMark(m.italic), "italic");
    btn("U", "Underline (Ctrl+U)", toggleMark(m.underline), "underline");
    tb.color("A", "Text color", "#000000", (color) => run(cmd.setMark(m.color, { color })));
    tb.color("HL", "Highlight color", "#FFFF00", (color) => run(cmd.setMark(m.highlight, { color })));
    sep();
    tb.icon("link", "Insert link (Ctrl+K)", () => void this.editLink());
    tb.icon("message-square-plus", "Add comment (Ctrl+Alt+M)", () => void this.addComment());
    tb.icon("image", "Insert image", () => pickImage((f) => void this.insertImage(f)));
    sep();
    tb.iconSelect("align-left", "Align", ALIGNS.map(([label, a]) => [label, a]), (v) => run(cmd.setAlign(v as Align)));
    tb.iconSelect("unfold-vertical", "Line spacing", LINE_SPACINGS, (v) => run(cmd.setLineSpacing(v)));
    sep();
    tb.icon("list", "Bulleted list (Ctrl+Shift+8)", () => run(cmd.toggleList(ctx, "bullet")));
    tb.icon("list-ordered", "Numbered list (Ctrl+Shift+7)", () => run(cmd.toggleList(ctx, "ordered")));
    tb.icon("outdent", "Decrease indent (Shift+Tab)", () => run(cmd.shiftLevel(-1)));
    tb.icon("indent", "Increase indent (Tab)", () => run(cmd.shiftLevel(1)));
    sep();
    this.brushBtn = tb.icon("paintbrush", "Paint format (double-click to keep painting)", () => this.startPainting(false));
    this.brushBtn.addEventListener("dblclick", () => this.startPainting(true));
    btn("Tx", "Clear formatting", cmd.clearFormatting);
    // Picture tools, while a picture is selected (Google's image options).
    sep();
    this.wrapSelect = tb.select("Text wrapping", "Wrap", WRAPS.map(([w, label]) => [label, w]), (v) => image((pos) => this.setWrap(pos, v as Wrap))());
    action("Crop", "Crop image", image((pos) => this.cropImage(pos)));
    action("Alt text", "Image description (alt text)", image((pos) => this.altText(pos)));
    tb.context(() => this.selectedImage() !== null);
    // Table tools, in a table.
    sep();
    tb.color("▇", "Cell fill", SHADES[0][1], (color) => run(cmd.setCellBackground((this.lastShade = color))));
    tb.iconSelect("square-dashed", "Table borders", [["All borders", "all"], ["Outside only", "outer"], ["No borders", "none"]], (v) => run(cmd.setTableBorders(v as "all" | "outer" | "none")));
    tb.context(() => this.selectedImage() === null && this.inTable());
    sep();
    const setMode = (value: string) => void (async () => {
      if (value === "suggesting" && !await this.author()) { mode.value = "editing"; return; }
      this.suggesting = value === "suggesting";
      mode.value = value;
      this.contentEl.toggleClass("is-suggesting", this.suggesting);
      if (this.suggesting) this.setReview(true);
      this.editor?.focus();
    })();
    const mode = tb.select("Editing mode", "Mode", [["Editing", "editing"], ["Suggesting", "suggesting"]], setMode);
    mode.value = this.suggesting ? "suggesting" : "editing";
    const showRuler = this.app.loadLocalStorage("doc-works-ruler") !== "off";
    this.rulerViewport?.toggle(showRuler && !this.settings.readOnly);

    // --- The menus: every command ---
    tb.fileMenu(this, (mn) => menuItem(mn, "Export…", () => void this.exportAs()));
    tb.menu("Edit", (mn) => {
      menuItem(mn, "Undo", () => run(undo));
      menuItem(mn, "Redo", () => run(redo));
      mn.addSeparator();
      for (const [title, c] of [["Cut", "cut"], ["Copy", "copy"], ["Paste", "paste"], ["Paste without formatting", "pasteAndMatchStyle"]] as const) {
        menuItem(mn, title, () => {
          this.editor?.focus();
          if (!runEditCommand(c)) new Notice("Paste with your keyboard.");
        });
      }
      menuItem(mn, "Select all", () => run(selectAll));
      menuItem(mn, "Copy link to heading", () => this.copyHeadingLink());
      menuItem(mn, "Delete", () => run(deleteSelection));
      mn.addSeparator();
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("View", (mn) => {
      menuItem(mn, "Editing", () => setMode("editing"), { checked: !this.suggesting });
      menuItem(mn, "Suggesting", () => setMode("suggesting"), { checked: this.suggesting });
      mn.addSeparator();
      menuItem(mn, "Page view", () => this.togglePageView(), { checked: this.pageView });
      menuItem(mn, "Ruler", () => this.toggleRuler(), { checked: this.rulerViewport?.style.display !== "none" });
      menuItem(mn, "Outline", () => this.toggleOutline(), { checked: this.outlineOpen });
      menuItem(mn, "Review panel", () => this.toggleReview(), { checked: this.reviewOpen });
      menuItem(mn, "Linked from", () => this.linked.toggle());
      subMenu(mn, "Zoom", (z) => {
        for (const value of ZOOMS) menuItem(z, zoomLabel(value), () => this.setZoom(value), { checked: this.zoom === value });
      });
      mn.addSeparator();
      menuItem(mn, "Light", () => this.setPaperTheme("light"), { checked: this.paperTheme === "light" });
      menuItem(mn, "Dark", () => this.setPaperTheme("dark"), { checked: this.paperTheme === "dark" });
    });
    tb.menu("Insert", (mn) => {
      menuItem(mn, "Image", () => pickImage((f) => void this.insertImage(f)));
      subMenu(mn, "Table", (s) => this.fillTableMenu(s));
      menuItem(mn, "Drawing", () => this.drawingDialog(null));
      subMenu(mn, "Chart", (s) => NEW_CHART_TYPES.forEach(([label, kind]) => menuItem(s, label, () => this.editChart(null, kind))));
      menuItem(mn, "Equation", () => this.equationDialog(null));
      menuItem(mn, "Special characters", () => specialCharacters(this.app, (ch) => this.editor?.dispatch(this.editor.state.tr.insertText(ch).scrollIntoView())));
      menuItem(mn, "Horizontal line", () => run(cmd.insertRule));
      // As Google Docs' Insert › Break.
      subMenu(mn, "Break", (s) => {
        menuItem(s, "Page break", () => run(cmd.insertPageBreak));
        menuItem(s, "Column break", () => run(cmd.insertColumnBreak));
        menuItem(s, "Section break (next page)", () => run(cmd.insertSectionBreak("nextPage")));
        menuItem(s, "Section break (continuous)", () => run(cmd.insertSectionBreak("continuous")));
      });
      menuItem(mn, "Footnote", () => void this.addFootnote());
      menuItem(mn, "Comment", () => void this.addComment());
      menuItem(mn, "Link", () => void this.editLink());
      subMenu(mn, "Table of contents", (s) => {
        menuItem(s, "Insert table of contents", () => {
          this.docx?.prepareToc();
          run(cmd.insertToc);
        });
        menuItem(s, "Update table of contents", () => {
          if (this.editor && !cmd.updateToc(this.editor.state, this.editor.dispatch)) new Notice("This document has no table of contents.");
        });
      });
      mn.addSeparator();
      menuItem(mn, "Header & footer", () => this.headerFooterDialog());
      menuItem(mn, "Watermark", () => this.watermarkDialog());
    });
    tb.menu("Format", (mn) => {
      subMenu(mn, "Text", (s) => {
        for (const [title, mark] of [["Bold", m.bold], ["Italic", m.italic], ["Underline", m.underline], ["Strikethrough", m.strike], ["Subscript", m.subscript], ["Superscript", m.superscript]] as const) {
          menuItem(s, title, () => run(toggleMark(mark)));
        }
        menuItem(s, "Code", () => run(cmd.toggleCode));
      });
      subMenu(mn, "Paragraph style", (s) => {
        for (const [label, id] of STYLES) menuItem(s, label, () => run(cmd.setStyle(id === "Normal" ? null : id)));
        s.addSeparator();
        menuItem(s, "Update style to match", () => this.updateStyle());
      });
      subMenu(mn, "Align", (s) => {
        for (const [label, a] of ALIGNS) menuItem(s, label, () => run(cmd.setAlign(a)));
      });
      subMenu(mn, "Line spacing", (s) => {
        for (const [label, v] of LINE_SPACINGS) menuItem(s, label, () => run(cmd.setLineSpacing(v)));
      });
      menuItem(mn, "Paragraph…", () => this.paragraphDialog());
      menuItem(mn, "Columns…", () => this.columnsDialog());
      mn.addSeparator();
      menuItem(mn, "Bulleted list", () => run(cmd.toggleList(ctx, "bullet")));
      menuItem(mn, "Numbered list", () => run(cmd.toggleList(ctx, "ordered")));
      mn.addSeparator();
      menuItem(mn, "Clear formatting", () => run(cmd.clearFormatting));
    });
    tb.menu("Tools", (mn) => {
      menuItem(mn, "Paint format", () => this.startPainting(false));
      menuItem(mn, "Word count", () => this.wordCountDialog());
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("Help", (mn) => menuItem(mn, "Keyboard shortcuts", () => shortcutsDialog(this.app, [...COMMON_SHORTCUTS, ...DOCS_SHORTCUTS])));
    if (versions) tb.side("history", "Version history", () => void versionHistoryDialog(this.app, this, this.settings.readOnly));
    const linkedBtn = tb.side("links-coming-in", "Linked from", () => this.linked.toggle(linkedBtn));
    if (!this.settings.readOnly) tb.side("messages-square", "Comments", () => this.toggleReview());

    mode.parentElement?.addClass("dw-push");
    const end = tb.end();
    this.autosave.status = end.status;
    this.wordsEl = createSpan({ cls: "dw-status docx-words" });
    this.wordsEl.hidden = this.app.loadLocalStorage(WORD_COUNT_KEY) === "off";
    end.el.prepend(this.wordsEl);
    tb.sync();
  }

  /** The selected picture's position (a node selection on an image), else null. */
  private selectedImage(): number | null {
    const sel = this.editor?.state.selection;
    return sel instanceof NodeSelection && sel.node.type === schema.nodes.image ? sel.from : null;
  }

  private inTable(): boolean {
    const $from = this.editor?.state.selection.$from;
    if (!$from) return false;
    for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === schema.nodes.table_cell) return true;
    return false;
  }

  private setWrap(pos: number, wrap: Wrap): void {
    const editor = this.editor, node = editor?.state.doc.nodeAt(pos);
    if (!editor || node?.type !== schema.nodes.image) return;
    editor.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, wrap, ...(wrap === "inline" ? { x: 0, y: 0, align: null } : {}) }));
  }

  private cropImage(pos: number): void {
    const editor = this.editor, node = editor?.state.doc.nodeAt(pos);
    if (!editor || node?.type !== schema.nodes.image) return;
    void cropDialog(this.app, attrsOf<ImageAttrs>(node).crop).then((crop) => {
      const now = editor.state.doc.nodeAt(pos);
      if (crop === undefined || now?.type !== schema.nodes.image) return;
      editor.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...now.attrs, ...recrop(attrsOf<ImageAttrs>(now), crop) }));
    });
  }

  private altText(pos: number): void {
    const editor = this.editor, node = editor?.state.doc.nodeAt(pos);
    if (!editor || node?.type !== schema.nodes.image) return;
    void promptText(this.app, "Image description (alt text)", attrsOf<ImageAttrs>(node).alt).then((alt) => {
      const now = alt === null ? null : editor.state.doc.nodeAt(pos);
      if (now) editor.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...now.attrs, alt }));
    });
  }

  private togglePageView(): void {
    this.setPageView(!this.pageView);
    this.app.saveLocalStorage(PAGE_VIEW_KEY, this.pageView ? "on" : "off");
  }

  private toggleRuler(): void {
    if (!this.ruler) return;
    const shown = this.rulerViewport?.style.display !== "none";
    this.rulerViewport?.toggle(!shown);
    this.app.saveLocalStorage("doc-works-ruler", shown ? "off" : "on");
  }

  private toggleOutline(): void {
    if (this.contentEl.hasClass("is-narrow") && this.outlineEl) return this.drawer(this.outlineEl, () => this.renderOutline());
    this.outlineOpen = !this.outlineOpen;
    this.app.saveLocalStorage("doc-works-outline", this.outlineOpen ? "on" : "off");
    this.outlineEl?.toggle(this.outlineOpen);
    this.renderOutline();
    this.fitDocument();
  }

  private toggleReview(): void {
    if (this.contentEl.hasClass("is-narrow") && this.reviewEl) return this.drawer(this.reviewEl, () => this.renderReview());
    this.setReview(!this.reviewOpen);
  }

  /** In a narrow pane a side panel opens over the document; its docked (wide) state is left as it was. */
  private drawer(panel: HTMLElement, render: () => void): void {
    if (this.closeDrawer) {
      const same = panel.hasClass("is-drawer");
      this.closeDrawer();
      if (same) return;
    }
    this.closeDrawer = openDrawer(panel, {
      root: this.contentEl,
      onClose: () => {
        this.closeDrawer = null;
        panel.toggle(panel === this.reviewEl ? this.reviewOpen : this.outlineOpen);
      },
    });
    render();
  }

  // ---------------------------------------------------------------------------
  // Word count, links, paragraph dialog
  // ---------------------------------------------------------------------------

  /** Font and Size show what's at the cursor: the text's own formatting, else its paragraph (and table) style. */
  private renderOutline(): void {
    const editor = this.editor, panel = this.outlineEl;
    if (!editor || !panel || !(this.outlineOpen || panel.hasClass("is-drawer"))) return;
    panel.empty();
    const title = panel.createDiv({ cls: "docx-review-title docx-outline-title", text: "Document outline" });
    const close = title.createEl("button", { cls: "clickable-icon docx-outline-close", attr: { "aria-label": "Close outline" } });
    setIcon(close, "x");
    // A drawer (narrow pane) just closes; docked, the outline turns off as View › Outline does.
    close.addEventListener("click", () => (panel.hasClass("is-drawer") ? this.closeDrawer?.() : this.toggleOutline()));
    const entries = outlineEntries(editor.state.doc);
    if (!entries.length) panel.createDiv({ cls: "docx-review-empty", text: "Headings appear here." });
    const active = [...entries].reverse().find(e => e.pos <= editor.state.selection.from)?.pos;
    // Nesting from the top heading level used (a document may start at Heading 2), as Google Docs' outline shows it.
    const top = Math.min(...entries.map((e) => e.level));
    // Headings notes link to ([[file.docx#Heading]]), by position inside the heading.
    const linked = new Map<number, number>();
    for (const g of this.linked.groups) {
      const h = g.subpath ? findHeading(editor.state.doc, g.subpath) : null;
      if (h) linked.set(h.pos + 1, (linked.get(h.pos + 1) ?? 0) + g.links.length);
    }
    for (const entry of entries) {
      const button = panel.createEl("button", { text: entry.text, cls: "docx-outline-entry", attr: { title: entry.text, "data-depth": String(Math.min(entry.level - top, 5)) } });
      const n = linked.get(entry.pos);
      if (n) button.createSpan({ cls: "dw-linked-badge", text: String(n), attr: { "aria-label": `Linked from ${n} note${n === 1 ? "" : "s"}` } });
      button.toggleClass("is-active", entry.pos === active);
      button.addEventListener("click", () => this.showHeading(entry.pos));
    }
  }

  private updateFormatControls(): void {
    const editor = this.editor;
    if (!editor || !this.ctx || !this.fontSelect || !this.sizeSelect) return;
    const { $from } = editor.state.selection;
    const marks = editor.state.storedMarks ?? $from.marks();
    const m = schema.marks;
    let tableStyle: string | null = null;
    for (let d = $from.depth; d > 0 && !tableStyle; d--) if ($from.node(d).type === schema.nodes.table) tableStyle = ($from.node(d).attrs.style as string | null) ?? null;
    const paraStyle = ($from.parent.attrs.style as string | null) ?? null;
    const font = (marks.find((x) => x.type === m.font)?.attrs.font as string | undefined) ?? this.ctx.styles.fontFor(paraStyle, tableStyle);
    const size = (marks.find((x) => x.type === m.size)?.attrs.pt as number | undefined) ?? this.ctx.styles.sizeFor(paraStyle, tableStyle);
    if (this.styleSelect) {
      showValue(this.styleSelect, paraStyle ?? "Normal");
      const update = Array.from(this.styleSelect.options).find((o) => o.value === "update");
      const name = STYLES.find(([, id]) => id === (paraStyle ?? "Normal"))?.[0] ?? paraStyle ?? "Normal";
      if (update) update.text = `Update '${name}' to match`;
    }
    showValue(this.fontSelect, font);
    showValue(this.sizeSelect, size === null ? null : String(size));
    const image = this.selectedImage();
    if (image !== null && this.wrapSelect) showValue(this.wrapSelect, attrsOf<ImageAttrs>(editor.state.doc.nodeAt(image)!).wrap);
    this.toolbar?.sync();
  }

  /** The ruler reads layout: once per frame, after typing, not on every keystroke (which forced a layout each time). */
  private scheduleRuler(): void {
    if (this.rulerFrame) return;
    this.rulerFrame = window.requestAnimationFrame(() => {
      this.rulerFrame = 0;
      this.ruler?.update();
    });
  }

  private scheduleWordCount(): void {
    if (this.wordsTimer) window.clearTimeout(this.wordsTimer);
    this.wordsTimer = window.setTimeout(() => this.updateWordCount(), 250);
  }

  /** Google Docs' Tools › Word count: pages, words, characters with and without spaces ("n of total" for a selection). */
  private wordCountDialog(): void {
    const editor = this.editor;
    if (!editor) return;
    const { doc, selection } = editor.state;
    const all = countWords(doc), part = selection.empty ? null : countWords(doc, selection.from, selection.to);
    const pages = pageAt(editor.state, doc.content.size);
    const show = (key: "words" | "chars" | "letters") => part ? `${part[key].toLocaleString()} of ${all[key].toLocaleString()}` : all[key].toLocaleString();
    const modal = new Modal(this.app);
    modal.titleEl.setText("Word count");
    const table = modal.contentEl.createDiv({ cls: "dw-count" });
    for (const [name, value] of [["Pages", pages === null ? "—" : pages.toLocaleString()], ["Words", show("words")], ["Characters", show("chars")], ["Characters excluding spaces", show("letters")]]) {
      const row = table.createDiv({ cls: "dw-count-row" });
      row.createSpan({ text: name });
      row.createSpan({ text: value });
    }
    const label = modal.contentEl.createEl("label", { cls: "dw-count-toggle" });
    const box = label.createEl("input", { type: "checkbox" });
    box.checked = !this.wordsEl?.hidden;
    label.appendText("Display word count while typing");
    const buttons = modal.contentEl.createDiv({ cls: "modal-button-container" });
    buttons.createEl("button", { text: "Cancel" }).addEventListener("click", () => modal.close());
    buttons.createEl("button", { text: "OK", cls: "mod-cta" }).addEventListener("click", () => {
      if (this.wordsEl) this.wordsEl.hidden = !box.checked;
      this.app.saveLocalStorage(WORD_COUNT_KEY, box.checked ? "on" : "off");
      this.toolbar?.fit();
      modal.close();
    });
    modal.open();
  }

  private updateWordCount(): void {
    if (!this.editor || !this.wordsEl) return;
    const { doc, selection } = this.editor.state;
    const all = countWords(doc).words;
    const label = (n: number) => `${n.toLocaleString()} ${n === 1 ? "word" : "words"}`;
    this.wordsEl.setText(selection.empty ? label(all) : `${countWords(doc, selection.from, selection.to).words.toLocaleString()} of ${label(all)}`);
  }

  private async editLink(): Promise<void> {
    if (!this.editor) return;
    const current = linkHref(this.editor.state);
    const url = await promptText(this.app, "Link address (empty removes the link)", current ?? "https://");
    if (url === null || !this.editor) return;
    const href = url.trim() && url.trim() !== "https://" ? url.trim() : null;
    if (!cmd.setLink(href)(this.editor.state, this.editor.dispatch) && href === null) new Notice("No link here.");
    this.editor.focus();
  }

  private paragraphDialog(): void {
    const editor = this.editor;
    if (!editor) return;
    const { $from } = editor.state.selection;
    const a = attrsOf<ParagraphAttrs>($from.parent.type === schema.nodes.paragraph ? $from.parent : schema.nodes.paragraph.create());
    const patch: Partial<ParagraphAttrs> = {};
    const modal = new Modal(this.app);
    modal.titleEl.setText("Paragraph");
    const num = (name: string, desc: string, value: number | null, set: (v: number | null) => void) =>
      new Setting(modal.contentEl)
        .setName(name)
        .setDesc(desc)
        .addText((t) => {
          t.inputEl.type = "number";
          t.setPlaceholder("Style default").setValue(value === null ? "" : String(value));
          t.onChange((v) => set(v.trim() === "" ? null : Number(v)));
        });
    num("Space before (pt)", "Empty uses the paragraph style.", a.spaceBefore, (v) => (patch.spaceBefore = v));
    num("Space after (pt)", "Empty uses the paragraph style.", a.spaceAfter, (v) => (patch.spaceAfter = v));
    num("Left indent (in)", "In steps of 0.5 in.", a.indent / 2, (v) => (patch.indent = Math.max(0, Math.round((v ?? 0) * 2))));
    num("First line (in)", "Negative for a hanging indent.", a.firstLine === null ? null : Math.round((a.firstLine / 72) * 100) / 100, (v) =>
      (patch.firstLine = v === null || v === 0 ? null : Math.round(v * 72 * 100) / 100)
    );
    new Setting(modal.contentEl).setName("Line spacing").addDropdown((d) =>
      d
        .addOptions({ "": "Style default", "1": "Single", "1.15": "1.15", "1.5": "1.5", "2": "Double" })
        .setValue(a.lineSpacing ?? "")
        .onChange((v) => (patch.lineSpacing = v || null))
    );
    // As in Google Docs' "Line & paragraph spacing" menu; shows the style's setting until changed here.
    const style = this.ctx?.styles.keep(a.style) ?? { next: false, lines: false };
    const toggle = (name: string, desc: string, value: boolean, set: (v: boolean) => void) =>
      new Setting(modal.contentEl)
        .setName(name)
        .setDesc(desc)
        .addToggle((t) => t.setValue(value).onChange(set));
    toggle("Keep with next", "Stay on the same page as the next paragraph.", a.keepNext ?? style.next, (v) => (patch.keepNext = v));
    toggle("Keep lines together", "Don't break this paragraph across pages.", a.keepLines ?? style.lines, (v) => (patch.keepLines = v));
    new Setting(modal.contentEl)
      .addButton((b) =>
        b
          .setButtonText("Apply")
          .setCta()
          .onClick(() => {
            cmd.setParagraph(patch)(editor.state, editor.dispatch);
            modal.close();
            editor.focus();
          })
      )
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }

  // ---------------------------------------------------------------------------
  // Review: comments, footnotes, tracked changes
  // ---------------------------------------------------------------------------

  private scheduleReview(): void {
    if (!this.reviewOpen) return;
    if (this.reviewTimer) window.clearTimeout(this.reviewTimer);
    this.reviewTimer = window.setTimeout(() => this.renderReview(), 400);
  }

  /** The comment author from settings, asked for (and saved) the first time. */
  private async author(): Promise<string | null> {
    if (this.settings.author.trim()) return this.settings.author.trim();
    const name = await promptText(this.app, "Your name (shown on comments; change it in settings)", "");
    if (!name?.trim()) return null;
    this.settings.author = name.trim();
    await this.saveSettings();
    return this.settings.author;
  }

  private async addComment(): Promise<void> {
    const editor = this.editor;
    if (!editor || !this.docx) return;
    if (editor.state.selection.empty) {
      new Notice("Select the text to comment on first.");
      return;
    }
    const author = await this.author();
    if (!author) return;
    const text = await promptText(this.app, "Comment", "");
    if (!text?.trim()) return;
    const id = this.docx.addComment(text.trim(), author);
    cmd.addCommentAnchors(id)(editor.state, editor.dispatch);
    this.showReview();
  }

  private async addFootnote(): Promise<void> {
    const editor = this.editor;
    if (!editor || !this.docx) return;
    const text = await promptText(this.app, "Footnote text", "");
    if (!text?.trim()) return;
    cmd.insertFootnote(this.docx.addFootnote(text.trim()))(editor.state, editor.dispatch);
    this.showReview();
  }

  private showReview(): void {
    this.setReview(true);
  }

  /** Resolve / Reply / Delete for comment `id`, shared by the review panel's cards and the right-click menu. */
  private commentActions(id: string): [string, () => void][] {
    const docx = this.docx, editor = this.editor;
    const c = docx?.comments().find((x) => x.id === id);
    if (!docx || !editor || !c) return [];
    return [
      [c.resolved ? "Reopen" : "Resolve", () => {
        docx.setCommentResolved(c.id, !c.resolved);
        this.autosave.markDirty();
        this.renderReview();
      }],
      ["Reply", () => void (async () => {
        const text = await promptText(this.app, "Reply to comment", "");
        if (!text?.trim()) return;
        const author = await this.author();
        if (!author || this.docx !== docx) return;
        docx.replyToComment(c.id, text.trim(), author);
        docx.setCommentResolved(c.id, false);
        this.autosave.markDirty();
        this.renderReview();
      })()],
      ["Delete", () => {
        docx.deleteComment(c.id);
        cmd.removeCommentAnchors(c.id)(editor.state, editor.dispatch);
        this.autosave.markDirty();
        this.renderReview();
      }],
    ];
  }

  /** Spell checker for our own right-click menus (null where there is none). */
  speller: Speller | null = electronSpeller();

  /**
   * Right-click in the document. Plain text keeps the app's native menu (with its spelling fixes); in a table, on a
   * picture, a link or commented text our menu shows instead: spelling fixes first (as the native menu would), then
   * Cut / Copy / Paste and the actions for what was clicked.
   */
  private contextMenu(e: MouseEvent): void {
    const editor = this.editor;
    if (!editor) return;
    const target = e.target as HTMLElement;
    const imageEl = target.closest<HTMLElement>(".docx-image");
    let pos: number | null = null;
    try {
      pos = editor.posAtCoords({ left: e.clientX, top: e.clientY })?.pos ?? null;
    } catch {
      // no layout (tests): fall back to the clicked element
    }
    if (pos === null) {
      try {
        pos = editor.posAtDOM(target, 0);
      } catch {
        return;
      }
    }
    const state = editor.state;
    const imagePos = imageEl ? [pos, pos - 1, pos + 1].find((p) => p >= 0 && state.doc.nodeAt(p)?.type === schema.nodes.image) ?? null : null;
    const $pos = state.doc.resolve(pos);
    let inTable = false;
    for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type === schema.nodes.table_cell) inTable = true;
    const link = [...$pos.marks(), ...($pos.nodeAfter?.marks ?? [])].find((m) => m.type === schema.marks.link)?.attrs.href as string | undefined;
    const commentId = target.closest<HTMLElement>("[data-comment]")?.dataset.comment ?? null;
    if (imagePos === null && !inTable && !link && !commentId) return;
    e.preventDefault();

    // The caret moves to the click unless it already selects there, as with the native menu.
    const sel = state.selection;
    if (imagePos === null && !sel.ranges.some((r) => r.$from.pos <= pos && pos <= r.$to.pos)) {
      editor.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(pos))));
    }
    const menu = new Menu();
    const item = (title: string, run: () => void) => menu.addItem((i) => i.setTitle(title).onClick(run));
    const run = (command: Command) => () => {
      command(editor.state, editor.dispatch);
      editor.focus();
    };

    // Spelling, as the native menu offers it.
    const word = imagePos === null ? wordAt(editor.state, editor.state.selection.empty ? editor.state.selection.head : pos) : null;
    if (word && this.speller?.misspelled(word.text)) {
      editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, word.from, word.to)));
      const fixes = this.speller.suggestions(word.text).slice(0, 5);
      for (const fix of fixes) item(fix, () => {
        editor.dispatch(editor.state.tr.insertText(fix, word.from, word.to));
        editor.focus();
      });
      if (!fixes.length) menu.addItem((i) => i.setTitle("No suggestions").setDisabled(true));
      item("Add to dictionary", () => this.speller?.addToDictionary(word.text));
      menu.addSeparator();
    }
    const edit = (command: "cut" | "copy" | "paste") => () => {
      editor.focus();
      if (!runEditCommand(command)) new Notice("Paste with your keyboard.");
    };
    item("Cut", edit("cut"));
    item("Copy", edit("copy"));
    item("Paste", edit("paste"));

    if (link) {
      menu.addSeparator();
      item("Open link", () => window.open(link));
      item("Edit link…", () => void this.editLink());
      item("Remove link", run(cmd.setLink(null)));
    }
    if (commentId) {
      menu.addSeparator();
      for (const [label, fn] of this.commentActions(commentId)) item(label === "Delete" ? "Delete comment" : label === "Reply" ? "Reply…" : label, fn);
    }
    if (inTable) {
      menu.addSeparator();
      item("Insert row above", run(cmd.addRow(false)));
      item("Insert row below", run(cmd.addRow(true)));
      item("Insert column left", run(cmd.addColumn(false)));
      item("Insert column right", run(cmd.addColumn(true)));
      item("Delete row", run(cmd.deleteRow));
      item("Delete column", run(cmd.deleteColumn));
      item("Merge with cell to the right", run(cmd.mergeRight));
      item("Merge with cell below", run(cmd.mergeDown));
      item("Split cell", run(cmd.splitCell));
      menu.addSeparator();
      subMenu(menu, "Cell shading", (m) => this.fillShadingMenu(m));
      item("Delete table", run(cmd.deleteTable));
    }
    if (imagePos !== null) {
      menu.addSeparator();
      item("Replace image…", () => pickImage((f) => void this.replaceImage(imagePos, f)));
      item("Crop…", () => this.cropImage(imagePos));
      item("Alt text…", () => this.altText(imagePos));
      item("Delete picture", () => editor.dispatch(editor.state.tr.delete(imagePos, imagePos + 1)));
    }
    menu.showAtMouseEvent(e);
  }

  private setReview(open: boolean): void {
    if (open && this.settings.readOnly) return;
    // Narrow pane: show it as a drawer (a new comment, suggesting mode) without changing the remembered docked state.
    if (this.contentEl.hasClass("is-narrow") && this.reviewEl) {
      if (open && !this.reviewEl.hasClass("is-drawer")) this.drawer(this.reviewEl, () => this.renderReview());
      return;
    }
    this.reviewOpen = open;
    this.reviewEl?.toggle(open);
    this.app.saveLocalStorage(REVIEW_KEY, open ? "on" : "off");
    this.renderReview();
    this.fitDocument();
  }

  /** Write the editor into the XML, change the XML, and load the result back (undo history restarts). */
  private rewriteBody(fn: (body: Element) => void): void {
    const editor = this.editor;
    const docx = this.docx;
    if (!editor || !docx) return;
    writeBody(docx.body, editor.state.doc, this.origins, docx);
    fn(docx.body);
    const read = readDocx(docx.body, { imageSrc: (id) => this.imageSrc(id), hyperlinkHref: (id) => docx.hyperlinkHref(id) });
    // One undoable step (Accept / Reject undo like Google Docs'); the old version's XML stays in the origins for undo.
    const doc = retainOrigins(read.doc, read.origins, this.origins);
    editor.dispatch(closeHistory(editor.state.tr.replaceWith(0, editor.state.doc.content.size, doc.content)
      .setDocAttribute("sectPr", doc.attrs.sectPr).setMeta("suggestionsRecorded", true)));
    this.autosave.markDirty();
    this.renderReview();
    repaginate(editor);
  }

  private imageSrc(rId: string): string | null {
    let src: string | null = null;
    this.editor?.state.doc.descendants((n) => {
      if (n.type === schema.nodes.image && n.attrs.rId === rId) src = n.attrs.src as string;
      return !src;
    });
    return src;
  }

  private renderReview(): void {
    const panel = this.reviewEl;
    const editor = this.editor;
    const docx = this.docx;
    // Closed panels aren't kept up to date, except while one is open as a drawer (narrow panes).
    if (!panel || !editor || !docx || !(this.reviewOpen || panel.hasClass("is-drawer"))) return;
    panel.empty();
    const doc = editor.state.doc;
    const goTo = (from: number, to = from) => {
      editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)).scrollIntoView());
      editor.focus();
    };
    const card = (parent: HTMLElement, meta: string, text: string, onText: (t: string) => void, actions: [string, () => void][]) => {
      const c = parent.createDiv({ cls: "docx-review-card" });
      if (meta) c.createDiv({ cls: "docx-review-meta", text: meta });
      const area = c.createEl("textarea", { text });
      area.rows = Math.min(6, Math.max(2, text.split("\n").length));
      area.addEventListener("change", () => {
        onText(area.value);
        this.autosave.markDirty();
      });
      const row = c.createDiv({ cls: "docx-review-actions" });
      for (const [label, fn] of actions) row.createEl("button", { text: label }).addEventListener("click", fn);
      // The same actions on right-click.
      c.addEventListener("contextmenu", (e) => {
        if ((e.target as HTMLElement).closest("textarea")) return;
        e.preventDefault();
        const menu = new Menu();
        for (const [label, fn] of actions) menu.addItem((i) => i.setTitle(label).onClick(fn));
        menu.showAtMouseEvent(e);
      });
      return c;
    };

    // Comments
    const title = panel.createDiv({ cls: "docx-review-title docx-outline-title", text: "Comments" });
    const close = title.createEl("button", { cls: "clickable-icon docx-outline-close", attr: { "aria-label": "Close comments" } });
    setIcon(close, "x");
    close.addEventListener("click", () => (panel.hasClass("is-drawer") ? this.closeDrawer?.() : this.setReview(false)));
    const ranges = commentRanges(doc);
    const comments = docx.comments();
    if (!comments.length) panel.createDiv({ cls: "docx-review-empty", text: "Select text and use Comment to add one." });
    for (const c of comments.filter(c => !c.parentId)) {
      const thread = panel.createDiv({ cls: `docx-comment-thread${c.resolved ? " is-resolved" : ""}` });
      const range = ranges.get(c.id);
      const quote = range ? doc.textBetween(range.from, range.to, " ").slice(0, 60) : "";
      card(thread, `${c.resolved ? "Resolved · " : ""}${c.author}${c.date ? ` · ${new Date(c.date).toLocaleDateString()}` : ""}${quote ? ` · “${quote}”` : ""}`, c.text, (t) => docx.setCommentText(c.id, t), [
        ...(range ? ([["Go to", () => goTo(range.from, range.to)]] as [string, () => void][]) : []),
        ...this.commentActions(c.id),
      ]).dataset.commentId = c.id;
      for (const reply of comments.filter(r => r.parentId === c.id)) {
        const replies = thread.createDiv({ cls: "docx-comment-reply" });
        card(replies, `${reply.author}${reply.date ? ` · ${new Date(reply.date).toLocaleDateString()}` : ""}`, reply.text,
          text => docx.setCommentText(reply.id, text), [["Delete reply", () => {
            docx.deleteComment(reply.id);
            this.autosave.markDirty();
            this.renderReview();
          }]]).dataset.commentId = reply.id;
      }
    }

    // Footnotes, in reference order
    panel.createDiv({ cls: "docx-review-title", text: "Footnotes" });
    const notes = new Map(docx.footnotes().map((f) => [f.id, f.text]));
    const order = footnoteIds(doc).filter((id) => notes.has(id));
    if (!order.length) panel.createDiv({ cls: "docx-review-empty", text: "Use Footnote to add one at the cursor." });
    order.forEach((id, i) => {
      card(panel, `${i + 1}`, notes.get(id) ?? "", (t) => docx.setFootnoteText(id, t), [
        [
          "Delete",
          () => {
            cmd.removeFootnoteRef(id)(editor.state, editor.dispatch);
            docx.deleteFootnote(id);
            this.autosave.markDirty();
            this.renderReview();
          },
        ],
      ]);
    });

    // Tracked changes (from the saved XML plus the current edits), in the body and in headers / footers
    panel.createDiv({ cls: "docx-review-title", text: "Tracked changes" });
    writeBody(docx.body, doc, this.origins, docx);
    const parts = HF_PARTS.flatMap((part) => {
      const root = docx.headerFooter(...splitPart(part));
      const [kind, type] = splitPart(part);
      const name = `${kind === "header" ? "Header" : "Footer"}${type === "default" ? "" : type === "first" ? " (first page)" : " (even pages)"}`;
      return root ? [{ root, name }] : [];
    });
    // One card per edit, as Google Docs lists suggestions.
    const changes = [
      ...groupRevisions(revisions(docx.body)).map((r) => ({ r, where: null as { root: Element; name: string } | null })),
      ...parts.flatMap((where) => groupRevisions(revisions(where.root)).map((r) => ({ r, where }))),
    ];
    if (!changes.length) {
      panel.createDiv({ cls: "docx-review-empty", text: "No tracked changes." });
      return;
    }
    /** Accept or reject in the body (`where` null) or one header / footer part. */
    const decide = (accept: boolean, where: { root: Element } | null, ids?: Set<string>) => {
      const fn = (root: Element) => (accept ? acceptRevisions(root, ids) : rejectRevisions(root, ids));
      if (where) {
        fn(where.root);
        this.readAllHF();
        this.autosave.markDirty();
        this.renderReview();
        if (this.editor) repaginate(this.editor);
      } else this.rewriteBody(fn);
    };
    const all = panel.createDiv({ cls: "docx-review-actions" });
    const everything = (accept: boolean) => {
      for (const where of parts) (accept ? acceptRevisions : rejectRevisions)(where.root);
      if (parts.length) this.readAllHF();
      this.rewriteBody((b) => (accept ? acceptRevisions(b) : rejectRevisions(b)));
    };
    all.createEl("button", { text: "Accept all" }).addEventListener("click", () => everything(true));
    all.createEl("button", { text: "Reject all" }).addEventListener("click", () => everything(false));
    const kinds = { insert: "Inserted", delete: "Deleted", format: "Formatted", replace: "Replaced" };
    for (const { r, where } of changes) {
      const c = panel.createDiv({ cls: `docx-review-card docx-change-${r.kind === "replace" ? "delete" : r.kind}` });
      c.createDiv({ cls: "docx-review-meta", text: `${kinds[r.kind]}${r.author ? ` by ${r.author}` : ""}${where ? ` · ${where.name}` : ""}` });
      c.createDiv({ cls: "docx-review-text", text: r.text.slice(0, 200) || "(formatting)" });
      if (r.kind === "replace") c.createDiv({ cls: "docx-review-text docx-review-inserted", text: `with ${r.inserted.slice(0, 200)}` });
      const row = c.createDiv({ cls: "docx-review-actions" });
      row.createEl("button", { text: "Accept" }).addEventListener("click", () => decide(true, where, r.ids));
      row.createEl("button", { text: "Reject" }).addEventListener("click", () => decide(false, where, r.ids));
    }
  }

  // ---------------------------------------------------------------------------
  // Find and replace
  // ---------------------------------------------------------------------------

  /** Ctrl+F / Ctrl+H: the find bar (also opened by a vault search result). */
  openFind(): void {
    if (!this.editor) return;
    if (this.findEl) {
      this.findEl.querySelector<HTMLInputElement>("input")?.focus();
      return;
    }
    const editor = this.editor;
    // Above the document area: the content element's child that holds the editor (outline, pages, review).
    const body = Array.from(this.contentEl.children).find((c) => this.hostEl && c.contains(this.hostEl)) ?? null;
    type Result = ({ kind: "text" } & Range) | ({ kind: "comment" } & CommentMatch);
    let matches: Result[] = [], index = -1;
    const bar = findBar(this.contentEl, body, {
      key: this.file?.path,
      options: [["Comments", "Search comments", true]],
      replaceTitles: ["Replace this document-text match", "Replace matching document text"],
      changed: () => show(0),
      step: (dir) => {
        refresh();
        const n = matches.length;
        if (n) show((index + dir + n) % n);
      },
      replace: () => {
        refresh();
        const hit = matches[index];
        if (!hit || hit.kind !== "text") return;
        replaceRangeCmd(hit, bar.replaceInput.value)(editor.state, editor.dispatch);
        show(index);
      },
      replaceAll: () => {
        const n = findRanges(editor.state.doc, input.value, bar.search()).length;
        replaceAllCmd(input.value, bar.replaceInput.value, bar.search())(editor.state, editor.dispatch);
        new Notice(`Replaced ${n} ${n === 1 ? "match" : "matches"}.`);
        show(0);
      },
      close: () => this.closeFind(),
    });
    this.findEl = bar.el;
    const { input, preview } = bar;
    const [comments] = bar.options;
    const refresh = () => {
      matches = [...findRanges(editor.state.doc, input.value, bar.search()).map(r => ({ ...r, kind: "text" as const })),
        ...(comments.checked && this.docx ? findCommentMatches(this.docx, input.value, bar.search()).map(r => ({ ...r, kind: "comment" as const })) : [])];
    };
    const show = (current: number) => {
      refresh();
      index = matches.length ? Math.min(Math.max(current, 0), matches.length - 1) : -1;
      const hit = matches[index];
      editor.dispatch(setFind(editor.state.tr, input.value, bar.search(), hit?.kind === "text" ? index : -1));
      bar.setCount(input.value, index, matches.length);
      bar.replaceButton.disabled = hit?.kind !== "text";
      bar.replaceAllButton.disabled = !matches.some(m => m.kind === "text");
      preview.empty();
      if (hit?.kind === "text") editor.dom.querySelector(".docx-find-hit.is-current")?.scrollIntoView({ block: "center" });
      else if (hit?.kind === "comment") showFindPreview(preview, `${hit.comment.resolved ? "Resolved comment" : "Comment"} · ${hit.comment.author}`, hit.comment.text, hit.start, hit.end, "Open comment", () => {
        this.setReview(true);
        const card = Array.from(this.reviewEl?.querySelectorAll<HTMLElement>("[data-comment-id]") ?? []).find(c => c.dataset.commentId === hit.comment.id);
        const area = card?.querySelector("textarea");
        card?.scrollIntoView({ block: "nearest" });
        area?.focus();
        area?.setSelectionRange(hit.start, hit.end);
      });
    };
    show(0);
    input.focus();
  }

  private closeFind(): void {
    if (!this.findEl) return;
    this.findEl.remove();
    this.findEl = null;
    if (this.editor) {
      this.editor.dispatch(setFind(this.editor.state.tr, "", false, -1));
      this.editor.focus();
    }
  }

  // ---------------------------------------------------------------------------
  // Pages, headers and footers
  // ---------------------------------------------------------------------------

  private readAllHF(): void {
    this.hf.clear();
    this.hfByPath.clear();
    const docx = this.docx;
    if (!docx) return;
    for (const part of HF_PARTS) {
      const root = docx.headerFooter(...splitPart(part));
      if (root) this.hf.set(part, readDocx(root, this.hfReadContext(part)).doc);
    }
    for (const path of hfPaths(docx)) {
      const root = docx.xml(path)?.documentElement;
      if (root) this.hfByPath.set(path, readDocx(root, this.hfReadContextFor(path)).doc);
    }
  }

  /** Images and links in a header / footer resolve through that part's own relationships. */
  private hfReadContext(part: HFPart, section?: number) {
    return this.hfReadContextFor(this.docx?.headerFooterPath(...splitPart(part), section) ?? null);
  }

  private hfReadContextFor(path: string | null) {
    const docx = this.docx;
    return { imageSrc: (id: string) => (path && this.hfUrls.get(path)?.get(id)) ?? null, hyperlinkHref: (id: string) => docx?.hyperlinkHref(id, path ?? undefined) ?? null };
  }

  /**
   * The header / footer a page shows: its section's (first page of the section, even pages or the default), with
   * page number fields filled in. `pos` is where the page starts.
   */
  private renderHF(kind: HF, page: number, total: number, pos = 0): Node | null {
    const docx = this.docx;
    if (!docx || !this.editor) return null;
    const { index, start } = sectionOfPos(this.editor.state.doc, pos);
    const first = pos <= start && docx.differentFirstPage(index);
    const type: HFType = first ? "first" : docx.differentOddEven() && page % 2 === 0 ? "even" : "default";
    const path = docx.headerFooterPath(kind, type, index);
    const doc = path ? this.hfByPath.get(path) : undefined;
    // The watermark sits outside the header's own stacking context, so it can go behind the page's text.
    const mark = kind === "header" && path ? this.watermarkEl() : null;
    if (!doc || !doc.textContent.trim() && !doc.content.size) return mark;
    const frag = DOMSerializer.fromSchema(schema).serializeFragment(doc.content);
    const wrap = createDiv({ cls: `docx-hf docx-hf-${kind}` });
    wrap.appendChild(frag);
    if (this.ctx) styleRenderedParagraphs(wrap, doc, this.ctx.styles);
    wrap.querySelectorAll<HTMLElement>("[data-field]").forEach((f) => f.setText(String(f.dataset.field === "NUMPAGES" ? total : page)));
    if (!mark) return wrap;
    const both = createDiv({ cls: "docx-hf-holder" });
    both.append(mark, wrap);
    return both;
  }

  private pagesPlugin(): Plugin {
    return pagesPlugin({
      setup: () => (this.pageView ? this.setup : null),
      pageSize: () => this.settings.pageSize,
      header: (page, total, pos) => this.renderHF("header", page, total, pos),
      footer: (page, total, pos) => this.renderHF("footer", page, total, pos),
      keep: (node) => this.ctx?.styles.keep(attrsOf<ParagraphAttrs>(node).style) ?? { next: false, lines: false },
      // Word's own layout only describes the document at its own page size.
      wordPages: () => (this.settings.pageSize === "document" ? this.docx?.wordPages() ?? null : null),
      wordBreak: (node) => {
        const key = (node.attrs as { key?: string | null }).key;
        const el = key ? this.origins.get(key)?.el : null;
        return !!el && startsWordPage(el);
      },
    });
  }

  private headerFooterDialog(): void {
    const docx = this.docx;
    if (!docx || !this.ctx) return;
    // The dialog edits the body's final w:sectPr (header references, first page option), which the editor also keeps
    // as its doc attribute: bring the body up to date first, and the attribute after.
    const editor = this.editor;
    if (editor) writeBody(docx.body, editor.state.doc, this.origins, docx);
    // The section the cursor is in (headers and footers belong to sections).
    const section = editor ? sectionOfPos(editor.state.doc, editor.state.selection.from).index : 0;
    const modal = new HeaderFooterModal(this.app, this.ctx, docx, section, (part) => this.hfReadContext(part, section), () => (this.suggesting ? this.settings.author : null), () => this.afterHeaderChange());
    modal.open();
  }

  /** Paint format: what was copied, and whether it stays on after one paint (double-click). */
  private painter: { format: PaintFormat; sticky: boolean } | null = null;
  private brushBtn: HTMLElement | null = null;

  private startPainting(sticky: boolean): void {
    const editor = this.editor;
    if (!editor) return;
    // A second click turns it off (a double-click's first click already turned it on).
    if (this.painter && !sticky) return this.stopPainting();
    this.painter = { format: cmd.copyFormat(editor.state), sticky };
    this.brushBtn?.addClass("is-active");
    this.hostEl?.addClass("is-painting");
    editor.focus();
  }

  private stopPainting(): void {
    this.painter = null;
    this.brushBtn?.removeClass("is-active");
    this.hostEl?.removeClass("is-painting");
  }

  private paintSelection(view: EditorView): void {
    const painter = this.painter;
    if (!painter || view.state.selection.empty) return;
    cmd.paintFormat(painter.format)(view.state, view.dispatch);
    if (!painter.sticky) this.stopPainting();
  }

  /**
   * Google Docs' "Update '<style>' to match": the cursor paragraph's own formatting becomes its style's. The same
   * formatting is then cleared from the selected paragraphs of that style, which now get it from the style.
   */
  private updateStyle(): void {
    const editor = this.editor, docx = this.docx, ctx = this.ctx;
    if (!editor || !docx || !ctx) return;
    const { $from, from, to } = editor.state.selection;
    const para = $from.parent;
    if (para.type !== schema.nodes.paragraph) return;
    const styleId = attrsOf<ParagraphAttrs>(para).style ?? "Normal";
    const { attrs, marks } = directFormat(para);
    updateStyleToMatch(docx, styleId, para);
    ctx.styles = new DocxStyles(docx.stylesXml(), docx.numberingXml(), docx.themeXml());
    const tr = editor.state.tr;
    editor.state.doc.nodesBetween(from, Math.max(to, from + 1), (node, pos) => {
      if (node.type !== schema.nodes.paragraph) return true;
      const a = attrsOf<ParagraphAttrs>(node);
      if ((a.style ?? "Normal") !== styleId) return false;
      const next: Record<string, unknown> = { ...node.attrs };
      for (const [k, v] of Object.entries(attrs)) if (v !== null && next[k] === v) next[k] = k === "indent" ? 0 : null;
      tr.setNodeMarkup(pos, undefined, next);
      for (const m of marks) tr.removeMark(pos + 1, pos + 1 + node.content.size, m);
      return false;
    });
    editor.dispatch(tr);
    refreshText(editor);
    this.autosave.markDirty();
    this.updateFormatControls();
  }

  /** After header / footer parts or section header references changed: reread them, sync the editor's sections, save. */
  private afterHeaderChange(): void {
    const docx = this.docx;
    if (!docx) return;
    this.readAllHF();
    const last = docx.body.lastElementChild;
    if (this.editor && last?.localName === "sectPr") {
      const sectPr = new XMLSerializer().serializeToString(last);
      if (sectPr !== this.editor.state.doc.attrs.sectPr) this.editor.dispatch(this.editor.state.tr.setDocAttribute("sectPr", sectPr).setMeta("addToHistory", false));
    }
    // Sections ended by a paragraph keep their w:sectPr on it: bring those up to date too.
    if (this.editor) {
      const sections = docx.sections().slice(0, -1).map((e) => new XMLSerializer().serializeToString(e));
      const tr = this.editor.state.tr;
      let k = 0;
      this.editor.state.doc.descendants((node, pos) => {
        if (node.type !== schema.nodes.paragraph) return true;
        if (node.attrs.sectPr) {
          const now = sections[k++];
          if (now && now !== node.attrs.sectPr) tr.setNodeMarkup(pos, undefined, { ...node.attrs, sectPr: now });
        }
        return false;
      });
      if (tr.docChanged) this.editor.dispatch(tr.setMeta("addToHistory", false));
    }
    this.autosave.markDirty();
    void this.autosave.flush();
    if (this.editor) repaginate(this.editor);
    this.renderReview();
  }
  /** The document's watermark (drawn on every page), and a URL for a picture watermark. */
  private watermark: Watermark | null = null;
  private watermarkUrl: string | null = null;

  private async loadWatermark(): Promise<void> {
    const docx = this.docx;
    this.watermark = docx ? readWatermark(docx) : null;
    const path = this.watermark?.kind === "image" ? this.watermark.path : null;
    this.watermarkUrl = null;
    if (docx && path) {
      const mime = IMAGE_MIME[path.split(".").pop()?.toLowerCase() ?? ""] ?? "image/png";
      const bytes = await docx.media(path);
      this.watermarkUrl = await dataUrl(new Blob([bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer], { type: mime }));
    }
  }

  /** The watermark drawn centred on the page's text area, behind the text (the page head is the page's top). */
  private watermarkEl(): HTMLElement | null {
    const wm = this.watermark, st = this.setup;
    if (!wm || !st) return null;
    const el = createDiv({ cls: "docx-watermark", attr: { "aria-hidden": "true" } });
    const w = (wm.width * 4) / 3, h = (wm.height * 4) / 3;
    el.setCssProps({
      "--x": `${st.left + this.pageWidth / 2}px`,
      "--y": `${st.top + (st.height - st.top - st.bottom) / 2}px`,
      "--w": `${w}px`,
      "--h": `${h}px`,
      "--docx-wm-rot": wm.kind === "text" && wm.diagonal ? "-45deg" : "0deg",
    });
    if (wm.kind === "text") {
      // Word stretches the text to its box (a VML text path that fits the shape).
      const svg = el.appendChild(createSvg("svg", { attr: { viewBox: `0 0 ${w} ${h}`, width: String(w), height: String(h) } }));
      const t = svg.appendChild(createSvg("text"));
      t.textContent = wm.text;
      for (const [k, v] of Object.entries({ x: "0", y: String(h * 0.8), "font-size": String(h), textLength: String(w), lengthAdjust: "spacingAndGlyphs", fill: wm.color, "fill-opacity": String(wm.opacity), "font-family": fontStack(wm.font) })) t.setAttribute(k, v);
    } else if (this.watermarkUrl) {
      const img = el.createEl("img", { attr: { src: this.watermarkUrl, alt: "" } });
      img.toggleClass("is-washout", wm.washout);
    }
    return el;
  }

  /** Google Docs' Insert › Watermark: text or a picture behind the text on every page. */
  private watermarkDialog(): void {
    const docx = this.docx;
    if (!docx) return;
    const current = this.watermark;
    const modal = new Modal(this.app);
    modal.titleEl.setText("Watermark");
    const t: TextWatermark = current?.kind === "text" ? { ...current } : { kind: "text", text: "CONFIDENTIAL", font: "Calibri", width: 0, height: 0, color: "#C0C0C0", opacity: 0.5, diagonal: true };
    let kind: "text" | "image" = current?.kind ?? "text";
    let size = "auto";
    let picture: { data?: Uint8Array; ext?: string; path?: string; w: number; h: number } | null =
      current?.kind === "image" ? { path: current.path, w: current.width, h: current.height } : null;
    let scale = "auto", washout = current?.kind === "image" ? current.washout : true;
    const body = modal.contentEl.createDiv();
    const draw = () => {
      body.empty();
      new Setting(body).setName("Type").addDropdown((d) => d.addOptions({ text: "Text", image: "Image" }).setValue(kind).onChange((v: string) => ((kind = v as typeof kind), draw())));
      if (kind === "text") {
        new Setting(body).setName("Text").addText((c) => {
          c.inputEl.setAttribute("aria-label", "Watermark text");
          c.setValue(t.text);
          c.inputEl.addEventListener("input", () => (t.text = c.inputEl.value));
        });
        new Setting(body).setName("Font").addDropdown((d) => d.addOptions(Object.fromEntries(FONTS.map((f) => [f, f]))).setValue(t.font).onChange((v: string) => (t.font = v)));
        new Setting(body).setName("Size").addDropdown((d) => d.addOptions({ auto: "Auto", ...Object.fromEntries(["36", "54", "72", "96", "120", "144"].map((x) => [x, `${x} pt`])) }).setValue(size).onChange((v: string) => (size = v)));
        new Setting(body).setName("Color").addColorPicker((c) => c.setValue(t.color).onChange((v: string) => (t.color = v.toUpperCase())));
        new Setting(body).setName("Transparency").addSlider((c) => c.setLimits(0, 90, 5).setValue(Math.round((1 - t.opacity) * 100)).onChange((v: number) => (t.opacity = 1 - v / 100)));
        new Setting(body).setName("Layout").addDropdown((d) => d.addOptions({ diagonal: "Diagonal", horizontal: "Horizontal" }).setValue(t.diagonal ? "diagonal" : "horizontal").onChange((v: string) => (t.diagonal = v === "diagonal")));
      } else {
        new Setting(body).setName("Picture").setDesc(picture ? "Chosen" : "None yet").addButton((b) =>
          b.setButtonText("Choose image…").onClick(() =>
            pickImage((file) => void (async () => {
              const src = await dataUrl(file);
              const { w, h } = await naturalSize(src);
              const ext = file.name.includes(".") ? file.name.split(".").pop()! : "png";
              picture = { data: new Uint8Array(await file.arrayBuffer()), ext, w: w * 0.75, h: h * 0.75 };
              draw();
            })())
          )
        );
        new Setting(body).setName("Scale").addDropdown((d) => d.addOptions({ auto: "Auto", "0.5": "50%", "1": "100%", "1.5": "150%", "2": "200%" }).setValue(scale).onChange((v: string) => (scale = v)));
        new Setting(body).setName("Faded").addToggle((c) => c.setValue(washout).onChange((v: boolean) => (washout = v)));
      }
    };
    draw();
    const apply = (wm: Watermark | null) => {
      const editor = this.editor;
      // Headers may be created for sections: bring the body (and its sections) up to date first.
      if (editor) writeBody(docx.body, editor.state.doc, this.origins, docx);
      setWatermark(docx, wm);
      void this.loadWatermark().then(() => this.afterHeaderChange());
      modal.close();
    };
    const buttons = new Setting(modal.contentEl).addButton((b) =>
      b.setButtonText("Apply").setCta().onClick(() => {
        const room = this.pageWidth * 0.75;
        if (kind === "text") {
          if (!t.text.trim()) return void new Notice("Enter the watermark text.");
          const box = textBox(t.text, t.font, size === "auto" ? null : Number(size), room);
          apply({ ...t, width: box.w, height: box.h });
        } else {
          if (!picture) return void new Notice("Choose a picture.");
          const s = scale === "auto" ? Math.min(1, room / picture.w) : Number(scale);
          apply({ kind: "image", data: picture.data, ext: picture.ext, path: picture.data ? undefined : picture.path, width: picture.w * s, height: picture.h * s, washout });
        }
      })
    );
    if (current) buttons.addButton((b) => b.setButtonText("Remove watermark").onClick(() => apply(null)));
    buttons.addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }


  /** Google Docs' Format › Columns: number, spacing and a line between. */
  private columnsDialog(): void {
    const editor = this.editor;
    if (!editor) return;
    const cols = { ...cmd.columnsAt(editor.state) };
    const modal = new Modal(this.app);
    modal.titleEl.setText("Columns");
    new Setting(modal.contentEl).setName("Number of columns").addDropdown((d) =>
      d
        .addOptions({ "1": "1", "2": "2", "3": "3" })
        .setValue(String(Math.min(3, cols.count)))
        .onChange((v) => (cols.count = Number(v)))
    );
    new Setting(modal.contentEl)
      .setName("Spacing (in)")
      .addText((t) => {
        t.inputEl.type = "number";
        t.setValue(String(Math.round((cols.space / 1440) * 100) / 100)).onChange((v) => (cols.space = Math.max(0, Number(v) || 0) * 1440));
      });
    new Setting(modal.contentEl).setName("Line between columns").addToggle((t) => t.setValue(cols.sep).onChange((v) => (cols.sep = v)));
    new Setting(modal.contentEl)
      .addButton((b) =>
        b
          .setButtonText("Apply")
          .setCta()
          .onClick(() => {
            cmd.setColumns(cols)(editor.state, editor.dispatch);
            modal.close();
            editor.focus();
          })
      )
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }

  private drawingDialog(pos: number | null): void {
    if (!this.editor || !this.docx) return;
    const editor=this.editor,docx=this.docx;
    const original=pos===null ? null : editor.state.doc.nodeAt(pos);
    const drawing=original ? readDrawing(original.attrs.xml as string) : newDrawing();
    if(!drawing?.editable)return;
    const modal=new Modal(this.app);modal.titleEl.setText(pos===null ? "Insert drawing" : "Edit drawing");
    modal.contentEl.addClass("docx-drawing-modal");drawingEditor(modal.contentEl,drawing);
    const done=modal.contentEl.createEl("button",{text:pos===null ? "Insert" : "Save"});
    done.addEventListener("click",()=>{
      if(this.editor!==editor || this.docx!==docx)return;
      if(!drawing.shapes.length)return void new Notice("Add a shape before saving.");
      if(pos!==null && editor.state.doc.nodeAt(pos)!==original)return void new Notice("The drawing changed while this dialog was open. Reopen it to continue.");
      const xml=writeDrawing(drawing,docx.nextDocPrId());
      const node=schema.nodes.opaque.create({xml,text:drawing.shapes.map(s=>s.text).join(" ")});
      if(pos===null)editor.dispatch(editor.state.tr.replaceSelectionWith(node));
      else editor.dispatch(editor.state.tr.replaceWith(pos,pos+original!.nodeSize,node));
      modal.close();editor.focus();
    });
    const cancel=modal.contentEl.createEl("button",{text:"Cancel"});cancel.addEventListener("click",()=>modal.close());modal.open();
  }

  /** Insert a chart (`pos` null) or edit the one at `pos`: a native Word chart, with the Slides / Sheets chart editor. */
  private editChart(pos: number | null, kind: ChartKind = "column"): void {
    const editor = this.editor, docx = this.docx;
    if (!editor || !docx) return;
    const original = pos === null ? null : editor.state.doc.nodeAt(pos);
    const found = original ? chartAt(docx, original.attrs.xml as string) : null;
    if (pos !== null && !found) return;
    chartDialog(this.app, {
      path: found?.path, data: found?.model.data, format: found?.model.format ?? newChartFormat(kind), palette: chartPalette(docxChartColors(docx)),
      apply: async (data, format) => {
        const xml = original && found ? await editedChartRun(docx, original.attrs.xml as string, data, format) : await chartRun(docx, data, format);
        if (this.editor !== editor || (pos !== null && editor.state.doc.nodeAt(pos) !== original)) throw new Error("The chart changed while this dialog was open. Reopen it to continue.");
        editor.dispatch(pos === null ? editor.state.tr.replaceSelectionWith(schema.nodes.opaque.create({ xml, text: "" })) : editor.state.tr.setNodeAttribute(pos, "xml", xml));
        editor.focus();
      },
    });
  }

  /** Insert an equation (`pos` null) or edit the one at `pos`, typed as LaTeX with a live preview. */
  private equationDialog(pos: number | null): void {
    const editor = this.editor;
    if (!editor) return;
    const modal = new Modal(this.app);
    modal.titleEl.setText(pos === null ? "Insert equation" : "Edit equation");
    modal.modalEl.addClass("docx-equation-modal");
    const input = modal.contentEl.createEl("textarea", {
      cls: "docx-equation-input",
      attr: { rows: "3", spellcheck: "false", placeholder: "Type LaTeX, e.g. \\frac{a}{b}, x^2, \\alpha, \\sum_{i=1}^{n}" },
    });
    input.value = pos === null ? "" : (cmd.equationAt(editor.state, pos) ?? "");
    const preview = modal.contentEl.createDiv({ cls: "docx-equation-preview" });
    const draw = () => preview.replaceChildren(DOMSerializer.renderSpec(document, mathSpec(latexToOmml(input.value))).dom);
    input.addEventListener("input", draw);
    // Symbols and structures, as in Google Docs' equation toolbar; the cursor lands in the first empty box.
    const insert = (text: string) => {
      const at = input.selectionStart;
      input.setRangeText(text, at, input.selectionEnd, "end");
      const box = text.indexOf("{}");
      if (box >= 0) input.setSelectionRange(at + box + 1, at + box + 1);
      input.focus();
      draw();
    };
    const palette = modal.contentEl.createDiv({ cls: "docx-equation-palette" });
    for (const row of EQUATION_PALETTE) {
      const line = palette.createDiv({ cls: "docx-equation-row" });
      for (const [label, text] of row) line.createEl("button", { text: label, attr: { "aria-label": text.trim() } }).addEventListener("click", () => insert(text));
    }
    const apply = (latex: string) => {
      cmd.setEquation(pos, latex)(editor.state, editor.dispatch);
      modal.close();
      editor.focus();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        apply(input.value);
      }
    });
    const buttons = new Setting(modal.contentEl).addButton((b) =>
      b
        .setButtonText(pos === null ? "Insert" : "Save")
        .setCta()
        .onClick(() => apply(input.value))
    );
    if (pos !== null) buttons.addButton((b) => b.setButtonText("Delete").onClick(() => apply("")));
    buttons.addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
    draw();
    input.focus();
  }

  private setPageView(on: boolean): void {
    this.pageView = on;
    this.pageViewBtn?.toggleClass("is-active", on);
    this.hostEl?.toggleClass("is-paged", on);
    if (this.editor) {
      refreshSectionWidths(this.editor);
      repaginate(this.editor);
    }
  }

  /** Save the document as a PDF, laid out on its pages as in the page view. */
  /** The pages as PDF (drawn in page view, which is turned on for the export if it's off). */
  private async pdfBytes(): Promise<Uint8Array | null> {
    if (!this.editor || !this.hostEl || !this.setup) return null;
    const paged = this.pageView;
    if (!paged) {
      this.setPageView(true);
      // ponytail: waits for the page layout's timer (60 ms, run twice when the header / footer size changes).
      await new Promise((r) => window.setTimeout(r, 400));
    }
    try {
      const { width, height } = this.setup;
      return await renderPdf(printPages(this.hostEl, height), { width, height }, "docx-theme-light");
    } finally {
      if (!paged) this.setPageView(false);
    }
  }


  /** Cell shading for the selected cells: none, a few light fills, or any colour from the system picker. */
  private fillShadingMenu(menu: Menu): void {
    const shade = (color: string | null) => {
      if (!this.editor) return;
      if (!cmd.setCellBackground(color)(this.editor.state, this.editor.dispatch)) new Notice("Click inside a table first.");
      this.editor.focus();
    };
    menuItem(menu, "None", () => shade(null));
    for (const [name, color] of SHADES) menuItem(menu, name, () => shade(color));
    menuItem(menu, "Custom…", () => pickColor(this.lastShade, (hex) => shade((this.lastShade = hex))));
  }

  /** Google Docs' Insert › Table and table tools. */
  private fillTableMenu(menu: Menu): void {
    const item = (title: string, command: Command) =>
      menu.addItem((i) =>
        i.setTitle(title).onClick(() => {
          if (!this.editor) return;
          if (!command(this.editor.state, this.editor.dispatch)) new Notice("Click inside a table first.");
          this.editor.focus();
        })
      );
    item("Insert 3×3 table", cmd.insertTable(3, 3));
    menu.addSeparator();
    item("Add row above", cmd.addRow(false));
    item("Add row below", cmd.addRow(true));
    item("Add column left", cmd.addColumn(false));
    item("Add column right", cmd.addColumn(true));
    menu.addSeparator();
    subMenu(menu, "Cell shading", (m) => this.fillShadingMenu(m));
    menu.addSeparator();
    item("Merge with cell to the right", cmd.mergeRight);
    item("Merge with cell below", cmd.mergeDown);
    item("Split cell", cmd.splitCell);
    menu.addSeparator();
    item("Borders: all", cmd.setTableBorders("all"));
    item("Borders: outside only", cmd.setTableBorders("outer"));
    item("Borders: none", cmd.setTableBorders("none"));
    menu.addItem((i) =>
      i.setTitle("Column width…").onClick(() => {
        const editor = this.editor;
        if (!editor) return;
        void promptText(this.app, "Column width (inches)", "1.5").then((v) => {
          const inches = Number(v);
          if (v === null || !(inches > 0)) return;
          if (!cmd.setColumnWidth(inches * 1440)(editor.state, editor.dispatch)) new Notice("Click inside a table first.");
          editor.focus();
        });
      })
    );
    item("Repeat header row on each page", cmd.toggleHeaderRow);
    const canBreak = this.editor ? cmd.rowsCanBreak(this.editor.state) : true;
    menu.addItem((i) =>
      i
        .setTitle("Allow rows to break across pages")
        .setChecked(canBreak)
        .onClick(() => {
          if (!this.editor) return;
          if (!cmd.setRowsCanBreak(!canBreak)(this.editor.state, this.editor.dispatch)) new Notice("Click inside a table first.");
          this.editor.focus();
        })
    );
    menu.addSeparator();
    item("Delete row", cmd.deleteRow);
    item("Delete column", cmd.deleteColumn);
    item("Delete table", cmd.deleteTable);
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------


}

/** Edits the default header and footer, each in its own small editor. */
class HeaderFooterModal extends Modal {
  private editors = new Map<HFPart, { view: EditorView; origins: Origins; changed: boolean }>();
  private changed = false;

  constructor(
    app: App,
    private ctx: EditorContext,
    private docx: DocxDocument,
    /** The section whose headers and footers this edits. */
    private sectionIndex: number,
    private readContext: (part: HFPart) => { imageSrc: (id: string) => string | null; hyperlinkHref: (id: string) => string | null },
    /** The suggestion author while the document is in suggesting mode, else null. */
    private suggestAuthor: () => string | null,
    private onDone: () => void
  ) {
    super(app);
  }

  onOpen(): void {
    const count = this.docx.sections().length;
    this.titleEl.setText(count > 1 ? `Header and footer: section ${this.sectionIndex + 1} of ${count}` : "Header and footer");
    this.modalEl.addClass("docx-hf-modal");
    this.render();
  }

  /** The options (as in Google Docs' header options), then an editor for each header and footer they call for. */
  private render(): void {
    this.contentEl.empty();
    const option = (name: string, value: boolean, set: (on: boolean) => void) =>
      new Setting(this.contentEl).setName(name).addToggle((t) =>
        t.setValue(value).onChange((on) => {
          set(on);
          this.changed = true;
          this.render();
        })
      );
    const at = this.sectionIndex;
    // Google Docs' / Word's "Link to previous", for sections after the first.
    if (at > 0) {
      for (const kind of ["header", "footer"] as const) {
        option(`${kind === "header" ? "Header" : "Footer"} same as previous section`, this.docx.linkedToPrevious(kind, at), (on) => {
          this.docx.setLinkedToPrevious(kind, at, on);
          // Editors open on the old parts: start them again from the section's current ones.
          for (const [part, e] of this.editors) if (part.startsWith(kind)) { e.view.destroy(); this.editors.delete(part); }
        });
      }
    }
    option("Different first page", this.docx.differentFirstPage(at), (on) => this.docx.setDifferentFirstPage(on, at));
    option("Different odd & even", this.docx.differentOddEven(), (on) => this.docx.setDifferentOddEven(on));
    const types: HFType[] = ["default", ...(this.docx.differentFirstPage(at) ? ["first" as const] : []), ...(this.docx.differentOddEven() ? ["even" as const] : [])];
    const label: Record<HFType, string> = { default: this.docx.differentOddEven() ? " (odd pages)" : "", first: " (first page)", even: " (even pages)" };
    for (const kind of ["header", "footer"] as const) {
      for (const type of types) this.section(`${kind}:${type}`, `${kind === "header" ? "Header" : "Footer"}${label[type]}`);
    }
  }

  private section(part: HFPart, title: string): void {
    new Setting(this.contentEl).setName(title).setHeading();
    const bar = this.contentEl.createDiv({ cls: "docx-hf-tools" });
    const host = this.contentEl.createDiv({ cls: "docx-hf-editor" });
    const kept = this.editors.get(part);
    let entry = kept;
    if (kept) host.appendChild(kept.view.dom);
    else {
      const root = this.docx.headerFooter(...splitPart(part), this.sectionIndex);
      const { doc, origins } = root ? readDocx(root, this.readContext(part)) : { doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create()), origins: new Map() as Origins };
      let lastKey: string | null = null;
      const view: EditorView = new EditorView(host, {
        state: EditorState.create({ doc, plugins: docxPlugins(this.ctx) }),
        handleKeyDown: (_v, event) => {
          lastKey = event.key;
          return false;
        },
        dispatchTransaction: (tr) => {
          const author = this.suggestAuthor();
          if (author) tr = suggestTransaction(tr, origins, this.docx, this.readContext(part), author, { forward: lastKey === "Delete" });
          lastKey = null;
          view.updateState(view.state.apply(tr));
          if (tr.docChanged) this.editors.get(part)!.changed = this.changed = true;
        },
      });
      entry = { view, origins, changed: false };
      this.editors.set(part, entry);
    }
    const view = entry!.view;
    const run = (c: Command) => {
      c(view.state, view.dispatch);
      view.focus();
    };
    const tb = new Toolbar(bar);
    for (const [icon, align] of [["align-left", "left"], ["align-center", "center"], ["align-right", "right"]] as [string, Align][]) {
      tb.icon(icon, `Align ${align}`, () => run(cmd.setAlign(align)));
    }
    tb.button("Page number", "Insert the page number", () => run((st, d) => (d?.(st.tr.replaceSelectionWith(pageNumberNode())), true)));
    bar.addEventListener("mousedown", (e) => (e.target as HTMLElement).closest("button") && e.preventDefault());
  }

  onClose(): void {
    if (this.changed) {
      for (const [part, { view, origins, changed }] of this.editors) {
        if (!changed) continue;
        const [kind, type] = splitPart(part);
        if (!view.state.doc.textContent.trim() && !this.docx.headerFooter(kind, type, this.sectionIndex)) continue;
        // A section linked to the previous one edits the shared header, as in Word.
        const target = this.docx.headerFooter(kind, type, this.sectionIndex) ?? this.docx.ensureHeaderFooter(kind, type, this.sectionIndex);
        writeBody(target, view.state.doc, origins, this.docx);
      }
      this.onDone();
    }
    this.editors.forEach((e) => e.view.destroy());
    this.contentEl.empty();
  }
}

/** Ask for a picture's crop, in percent of each side; resolves undefined when cancelled, null for no crop. */
function cropDialog(app: App, current: ImageAttrs["crop"]): Promise<ImageAttrs["crop"] | undefined> {
  return new Promise((resolve) => {
    let result: ImageAttrs["crop"] | undefined;
    const crop = { l: 0, t: 0, r: 0, b: 0, ...(current ?? {}) };
    const modal = new Modal(app);
    modal.titleEl.setText("Crop picture");
    for (const [key, name] of [["l", "Left"], ["t", "Top"], ["r", "Right"], ["b", "Bottom"]] as const) {
      new Setting(modal.contentEl).setName(`${name} (%)`).addText((t) => {
        t.inputEl.type = "number";
        t.inputEl.min = "0";
        t.inputEl.max = "95";
        t.setValue(String(Math.round(crop[key] * 1000) / 10)).onChange((v) => (crop[key] = Math.min(95, Math.max(0, Number(v) || 0)) / 100));
      });
    }
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Apply").setCta().onClick(() => {
        if (crop.l + crop.r >= 0.95 || crop.t + crop.b >= 0.95) return new Notice("Leave some of the picture showing.");
        result = crop.l || crop.t || crop.r || crop.b ? crop : null;
        modal.close();
      }))
      .addButton((b) => b.setButtonText("Reset").onClick(() => {
        result = null;
        modal.close();
      }))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}

/** Every header / footer part some section of the document shows. */
function hfPaths(docx: DocxDocument): string[] {
  const out = new Set<string>();
  docx.sections().forEach((_, i) => {
    for (const part of HF_PARTS) {
      const path = docx.headerFooterPath(...splitPart(part), i);
      if (path) out.add(path);
    }
  });
  return [...out];
}

/** Which section (0-based) the block at `pos` belongs to, and where that section starts: a paragraph carrying a w:sectPr ends its section. */
function sectionOfPos(doc: PMNode, pos: number): { index: number; start: number } {
  let index = 0;
  let start = 0;
  doc.descendants((node, at) => {
    if (at >= pos) return false;
    if (node.type === schema.nodes.paragraph) {
      if (node.attrs.sectPr && at + node.nodeSize <= pos) {
        index++;
        start = at + node.nodeSize;
      }
      return false;
    }
    return true;
  });
  return { index, start };
}

/** A text watermark's box in points: as wide as the text area for "Auto", else the text's width at `size`. */
function textBox(text: string, font: string, size: number | null, room: number): { w: number; h: number } {
  const ctx = createEl("canvas").getContext("2d");
  let ratio = 0.6 * Math.max(1, text.length); // width per unit of height, when there is nothing to measure with
  if (ctx) {
    ctx.font = `100px ${fontStack(font)}`;
    const width = ctx.measureText(text).width;
    if (width > 0) ratio = width / 100;
  }
  if (size === null) return { w: +room.toFixed(2), h: +(room / ratio).toFixed(2) };
  return { w: +(size * ratio).toFixed(2), h: size };
}

/** The word (letters, digits, apostrophes) around `pos` in its paragraph, or null. */
function wordAt(state: EditorState, pos: number): { from: number; to: number; text: string } | null {
  const $pos = state.doc.resolve(pos);
  const block = $pos.parent;
  if (!block.isTextblock) return null;
  const start = $pos.start();
  const text = block.textBetween(0, block.content.size, undefined, "\ufffc");
  const at = pos - start;
  const isWord = (ch: string | undefined) => !!ch && /[\p{L}\p{N}'’]/u.test(ch);
  let a = at, b = at;
  while (a > 0 && isWord(text[a - 1])) a--;
  while (b < text.length && isWord(text[b])) b++;
  const word = text.slice(a, b).replace(/^['’]+|['’]+$/g, "");
  if (!word) return null;
  const lead = text.slice(a, b).indexOf(word);
  return { from: start + a + lead, to: start + a + lead + word.length, text: word };
}

/** Read-only files: transactions that would change the document are dropped (selection, find and decorations still work). */
function readOnlyPlugin(readOnly: () => boolean): Plugin {
  return new Plugin({ filterTransaction: (tr) => !tr.docChanged || !readOnly() });
}
