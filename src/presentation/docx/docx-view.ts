import { App, FileView, Menu, Modal, Notice, Scope, Setting, TFile, WorkspaceLeaf } from "obsidian";
import { Command, EditorState, Plugin, TextSelection } from "prosemirror-state";
import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { toggleMark } from "prosemirror-commands";
import { redo, undo } from "prosemirror-history";
import { DocxDocument, HFType } from "../../infrastructure/docx/docx-document";
import { Origins, readDocx } from "../../domain/docx/docx-reader";
import { writeBody } from "../../domain/docx/docx-writer";
import { DocxStyles } from "../../domain/docx/docx-styles";
import { docxSchema as schema, Align } from "../../domain/docx/schema";
import { FONTS, PaperTheme, promptText, showValue, Toolbar } from "../common/toolbar";
import { pickImage } from "../common/images";
import { cmd, commentRanges, docxPlugins, EditorContext, footnoteIds, linkHref, pageNumberNode, styleRenderedParagraphs } from "./docx-editor";
import { pagesPlugin, printPages, refreshSectionWidths, repaginate, sectionWidthsPlugin } from "./docx-pages";
import { rewrap, wrapBothSidesPlugin } from "./docx-wrap";
import { latexToOmml, mathSpec } from "../../domain/docx/math";
import { exportPdf } from "../common/pdf";
import { acceptRevisions, rejectRevisions, revisions } from "../../domain/docx/revisions";
import { countWords, findKey, replaceAllCmd, replaceRangeCmd, setFind } from "./docx-find";
import { headerFooterType, pageSetup, PageSetup, PageSize, sectionSetups } from "../../domain/docx/pages";
import { attrsOf, ParagraphAttrs } from "../../domain/docx/schema";
import { DocxImageView } from "./docx-image-view";
import { TextBoxView } from "./docx-textbox-view";
import { suggestTransaction } from "../../domain/docx/suggestions";
import { DocxRuler, outlineEntries } from "./docx-ruler";

export interface DocWorksSettings {
  /** "document": pages use the document's own size; "letter": always 8.5 x 11 in. */
  pageSize: PageSize;
  /** Name shown on new comments; asked for on the first comment when empty. */
  author: string;
}
export const DEFAULT_SETTINGS: DocWorksSettings = { pageSize: "document", author: "" };

export const VIEW_TYPE_DOCX = "doc-works-docx-view";

const SAVE_DELAY_MS = 800;
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

/** Open a toolbar menu under its button. */
const showBelow = (menu: Menu, button: HTMLElement) => {
  const r = button.getBoundingClientRect();
  menu.showAtPosition({ x: r.left, y: r.bottom });
};

/** Equation palette rows: label, LaTeX to insert. */
const EQUATION_PALETTE: [string, string][][] = [
  [["a/b", "\\frac{}{}"], ["√", "\\sqrt{}"], ["ⁿ√", "\\sqrt[]{}"], ["xⁿ", "^{}"], ["xₙ", "_{}"], ["∑", "\\sum_{}^{}{}"], ["∫", "\\int_{}^{}{}"], ["lim", "\\lim_{}{}"], ["( )", "\\left(\\right)"], ["[ ]", "\\left[\\right]"], ["x̂", "\\hat{}"], ["x̄", "\\overline{}"], ["vec", "\\vec{}"], ["▦", "\\begin{matrix}{}&\\\\&\\end{matrix}"]],
  [["α", "\\alpha "], ["β", "\\beta "], ["γ", "\\gamma "], ["δ", "\\delta "], ["ε", "\\varepsilon "], ["θ", "\\theta "], ["λ", "\\lambda "], ["μ", "\\mu "], ["π", "\\pi "], ["σ", "\\sigma "], ["φ", "\\varphi "], ["ω", "\\omega "], ["Δ", "\\Delta "], ["Σ", "\\Sigma "], ["Ω", "\\Omega "]],
  [["±", "\\pm "], ["×", "\\times "], ["÷", "\\div "], ["·", "\\cdot "], ["∞", "\\infty "], ["∂", "\\partial "], ["∇", "\\nabla "], ["≤", "\\leq "], ["≥", "\\geq "], ["≠", "\\neq "], ["≈", "\\approx "], ["≡", "\\equiv "], ["∈", "\\in "], ["⊂", "\\subset "], ["∪", "\\cup "], ["∩", "\\cap "], ["∀", "\\forall "], ["∃", "\\exists "], ["→", "\\to "], ["⇒", "\\Rightarrow "], ["⇔", "\\Leftrightarrow "]],
];

const PAGE_VIEW_KEY = "doc-works-page-view";
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
  allowNoFile = false;
  paperTheme: "light" | "dark" = "light";

  private docx: DocxDocument | null = null;
  private origins: Origins = new Map();
  private editor: EditorView | null = null;
  private pageWidth = 624;
  private dirty = false;
  private saveTimer: number | null = null;
  private saving: Promise<void> = Promise.resolve();
  private saveEl: HTMLElement | null = null;
  private syncTheme: ((t: PaperTheme) => void) | null = null;
  private ctx: EditorContext | null = null;
  private setup: PageSetup | null = null;
  /** The body's own page setup (`setup` is as wide as the widest section's page). */
  private finalSetup: PageSetup | null = null;
  private pageView = this.app.loadLocalStorage(PAGE_VIEW_KEY) !== "off";
  private hostEl: HTMLElement | null = null;
  private wordsEl: HTMLElement | null = null;
  private wordsTimer: number | null = null;
  private findEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  /** Header and footer content, for drawing them in the page view. */
  private hf = new Map<HFPart, PMNode>();
  private styleSelect: HTMLSelectElement | null = null;
  private fontSelect: HTMLSelectElement | null = null;
  private sizeSelect: HTMLSelectElement | null = null;
  /** rId -> data URL for pictures in the header and footer parts. */
  private hfUrls = new Map<HFPart, Map<string, string>>();
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
    this.scope.register(["Mod"], "f", () => {
      this.openFind();
      return false;
    });
  }

  onPaneMenu(menu: Menu, source: string): void {
    super.onPaneMenu(menu, source);
    menu.addItem((i) => i.setTitle("Export to PDF").setIcon("file-output").setSection("action").onClick(() => void this.exportPdf()));
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
    this.contentEl.empty();
    this.contentEl.addClass("dw-root");
    const status = this.contentEl.createDiv({ cls: "dw-loading", text: "Loading document..." });

    let docx: DocxDocument;
    let urls: Map<string, string>;
    try {
      docx = await DocxDocument.load(await this.app.vault.readBinary(file));
      urls = await docx.imageUrls();
      for (const part of HF_PARTS) {
        const path = docx.headerFooterPath(...splitPart(part));
        if (path) this.hfUrls.set(part, await docx.imageUrls(path));
      }
    } catch (err) {
      status.setText(`Could not open this .docx file: ${(err as Error).message}`);
      console.error(err);
      return;
    }
    status.remove();
    this.docx = docx;

    const { doc, origins } = readDocx(docx.body, {
      imageSrc: (id) => urls.get(id) ?? null,
      hyperlinkHref: (id) => docx.hyperlinkHref(id),
    });
    this.origins = origins;
    this.readAllHF();

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
        return fmt ? (fmt === "bullet" ? "bullet" : "ordered") : null;
      },
    };

    const toolbar = this.contentEl.createDiv({ cls: "dw-toolbar" });
    const body = this.contentEl.createDiv({ cls: "docx-body" });
    this.outlineEl = body.createDiv({ cls: "docx-outline", attr: { "aria-label": "Document outline" } });
    this.outlineEl.toggle(this.outlineOpen);
    const pane = body.createDiv({ cls: "docx-editor-pane" });
    const rulerViewport = this.rulerViewport = pane.createDiv({ cls: "docx-ruler-viewport" });
    const host = (this.hostEl = pane.createDiv({ cls: "docx-editor-host" }));
    this.ruler = new DocxRuler(rulerViewport, () => this.editor, () => { this.applyPageSize(); if (this.editor) repaginate(this.editor); });
    host.addEventListener("scroll", () => this.ruler?.align());
    this.reviewEl = body.createDiv({ cls: "docx-review" });
    this.reviewEl.toggle(this.reviewOpen);
    this.applyPageSize();
    host.toggleClass("is-paged", this.pageView);
    this.ctx = ctx;

    // Double-click a page's header or footer to edit it, as in Word.
    host.addEventListener("dblclick", (e) => {
      if ((e.target as HTMLElement).closest(".docx-page-first, .docx-page-head, .docx-page-foot")) this.headerFooterDialog();
    });
    // Click an equation to edit it, as in Google Docs.
    host.addEventListener("click", (e) => {
      const el = (e.target as HTMLElement).closest(".docx-math")?.parentElement;
      const editor = this.editor;
      if (!el || !editor || !editor.dom.contains(el)) return;
      const at = editor.posAtDOM(el, 0);
      const pos = [at, at - 1].find((p) => p >= 0 && cmd.equationAt(editor.state, p) !== null);
      if (pos !== undefined) this.equationDialog(pos);
    });
    this.editor = new EditorView(host, {
      state: EditorState.create({ doc, plugins: [...docxPlugins(ctx, { "Mod-k": () => (void this.editLink(), true) }), this.pagesPlugin(), wrapBothSidesPlugin(), sectionWidthsPlugin(() => (this.pageView ? this.setup : null), () => this.finalSetup, () => this.settings.pageSize)] }),
      nodeViews: {
        image: (node, view, getPos) =>
          new DocxImageView(node, view, getPos, {
            replace: (pos) => pickImage((f) => void this.replaceImage(pos, f)),
            editAlt: (pos, current) =>
              void promptText(this.app, "Image description (alt text)", current).then((alt) => {
                const node = alt === null ? null : view.state.doc.nodeAt(pos);
                if (node) view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, alt }));
              }),
          }),
        // Text boxes are edited in place; other run content is drawn as usual.
        opaque: (node, view, getPos) => new TextBoxView(node, view, getPos, ctx, { imageSrc: (id) => urls.get(id) ?? null, hyperlinkHref: (id) => docx.hyperlinkHref(id) }, docx),
      },
      dispatchTransaction: (tr) => {
        if (!this.editor) return;
        if (this.suggesting) tr = suggestTransaction(tr, this.origins, docx,
          { imageSrc: id => this.imageSrc(id) ?? urls.get(id) ?? null, hyperlinkHref: id => docx.hyperlinkHref(id) }, this.settings.author, { forward: this.lastKey === "Delete" });
        this.lastKey = null;
        const applied = this.editor.state.applyTransaction(tr);
        this.editor.updateState(applied.state);
        const changed = applied.transactions.some(t => t.docChanged);
        if (changed) this.markDirty();
        if (changed || tr.selectionSet) {
          this.scheduleWordCount();
          this.updateFormatControls();
          this.ruler?.update();
          this.renderOutline();
        }
        if (changed) this.scheduleReview();
      },
      handleKeyDown: (_view, event) => {
        this.lastKey = event.key;
        return false;
      },
      handleDOMEvents: {
        click: (_view, event) => {
          // Ctrl / Cmd + click follows a link.
          const a = (event.target as HTMLElement).closest<HTMLAnchorElement>("a.docx-link");
          if (!a || !(event.ctrlKey || event.metaKey)) return false;
          event.preventDefault();
          if (/^https?:|^mailto:/.test(a.getAttribute("href") ?? "")) window.open(a.getAttribute("href")!);
          return true;
        },
        blur: () => {
          if (this.dirty) void this.saveNow();
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

    this.buildToolbar(toolbar, ctx);
    this.setPaperTheme(this.paperTheme);
    window.addEventListener("blur", this.onWindowBlur);
    this.resizeObserver = new ResizeObserver(() => {
      this.fitDocument();
    });
    this.resizeObserver.observe(host);
    this.updateWordCount();
    this.renderReview();
    this.ruler.update();
    this.fitDocument();
    this.renderOutline();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
    this.resizeObserver?.disconnect();
    this.closeFind();
    for (const t of [this.wordsTimer, this.reviewTimer]) if (t) window.clearTimeout(t);
    await this.requestSave();
    await this.saving;
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

  private fitDocument(): void {
    const editor = this.editor, host = this.hostEl, setup = this.setup;
    if (!editor || !host || !setup) return;
    const padding = getComputedStyle(host);
    const available = host.clientWidth - (Number.parseFloat(padding.paddingLeft) || 0) - (Number.parseFloat(padding.paddingRight) || 0);
    const zoom = this.pageView && available > 0 ? Math.min(1, available / setup.width) : 1;
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
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
  }

  private onWindowBlur = () => {
    if (this.dirty) void this.saveNow();
  };

  setPaperTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    this.contentEl.toggleClass("docx-theme-light", theme === "light");
    this.contentEl.toggleClass("docx-theme-dark", theme === "dark");
    this.syncTheme?.(theme);
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
    const tb = new Toolbar(bar);
    const action = (label: string, title: string, onClick: () => void, cls = "") => tb.button(label, title, onClick, cls);
    const btn = (label: string, title: string, command: Command, cls = "") => action(label, title, () => run(command), cls);
    const sep = () => tb.gap();
    const m = schema.marks;

    const mode = tb.select("Editing mode", "Mode", [["Editing", "editing"], ["Suggesting", "suggesting"]], value => void (async () => {
      if (value === "suggesting" && !await this.author()) { mode.value = "editing"; return; }
      this.suggesting = value === "suggesting";
      this.contentEl.toggleClass("is-suggesting", this.suggesting);
      if (this.suggesting) this.setReview(true);
      this.editor?.focus();
    })());
    mode.value = this.suggesting ? "suggesting" : "editing";
    const outline = tb.icon("list-tree", "Document outline", () => {
      this.outlineOpen = !this.outlineOpen;
      this.app.saveLocalStorage("doc-works-outline", this.outlineOpen ? "on" : "off");
      this.outlineEl?.toggle(this.outlineOpen);
      outline.toggleClass("is-active", this.outlineOpen);
      this.renderOutline();
      this.fitDocument();
    });
    outline.toggleClass("is-active", this.outlineOpen);
    const ruler = tb.icon("ruler", "Show ruler", () => {
      if (!this.ruler) return;
      const shown = this.rulerViewport?.style.display !== "none";
      this.rulerViewport?.toggle(!shown);
      ruler.toggleClass("is-active", !shown);
      this.app.saveLocalStorage("doc-works-ruler", shown ? "off" : "on");
    });
    const showRuler = this.app.loadLocalStorage("doc-works-ruler") !== "off";
    this.rulerViewport?.toggle(showRuler);
    ruler.toggleClass("is-active", showRuler);

    btn("↶", "Undo (Ctrl+Z)", undo);
    btn("↷", "Redo (Ctrl+Y)", redo);
    sep();
    this.styleSelect = tb.select("Paragraph style", "Style", STYLES, (v) => run(cmd.setStyle(v === "Normal" ? null : v)));
    this.fontSelect = tb.select("Font", "Font", FONTS.map((f) => [f, f]), (font) => run(cmd.setMark(m.font, { font })));
    this.sizeSelect = tb.select("Font size", "Size", FONT_SIZES.map((s) => [`${s} pt`, String(s)]), (v) => run(cmd.setMark(m.size, { pt: Number(v) })));
    sep();
    btn("B", "Bold (Ctrl+B)", toggleMark(m.bold), "bold");
    btn("I", "Italic (Ctrl+I)", toggleMark(m.italic), "italic");
    btn("U", "Underline (Ctrl+U)", toggleMark(m.underline), "underline");
    btn("S", "Strikethrough (Ctrl+Shift+X)", toggleMark(m.strike), "strike");
    btn("x₂", "Subscript", toggleMark(m.subscript));
    btn("x²", "Superscript", toggleMark(m.superscript));
    btn("</>", "Code font (Ctrl+`)", cmd.toggleCode);
    tb.color("A", "Text color", "#000000", (color) => run(cmd.setMark(m.color, { color })));
    tb.color("HL", "Highlight color", "#FFFF00", (color) => run(cmd.setMark(m.highlight, { color })));
    btn("Tx", "Clear formatting", cmd.clearFormatting);
    sep();
    for (const [icon, align] of [["align-left", "left"], ["align-center", "center"], ["align-right", "right"], ["align-justify", "justify"]] as [string, Align][]) {
      tb.icon(icon, `Align ${align}`, () => run(cmd.setAlign(align)));
    }
    tb.iconSelect("unfold-vertical", "Line spacing", [["1.0", "1"], ["1.15", "1.15"], ["1.5", "1.5"], ["2.0", "2"]], (v) => run(cmd.setLineSpacing(v)));
    tb.icon("list", "Bulleted list", () => run(cmd.toggleList(ctx, "bullet")));
    tb.icon("list-ordered", "Numbered list", () => run(cmd.toggleList(ctx, "ordered")));
    tb.icon("outdent", "Decrease indent / list level (Shift+Tab)", () => run(cmd.shiftLevel(-1)));
    tb.icon("indent", "Increase indent / list level (Tab)", () => run(cmd.shiftLevel(1)));
    action("Paragraph…", "Spacing and indents", () => this.paragraphDialog());
    tb.icon("link", "Add, change or remove a link (Ctrl+K; Ctrl+click follows it)", () => void this.editLink());
    sep();
    const tableBtn = action("Table ▾", "Table", () => this.tableMenu(tableBtn));
    action("Image", "Insert image", () => pickImage((f) => void this.insertImage(f)));
    const breakBtn = action("Break ▾", "Page, column and section breaks", () => {
      // As Google Docs' Insert › Break.
      const menu = new Menu();
      const entry = (title: string, command: Command) => menu.addItem((i) => i.setTitle(title).onClick(() => run(command)));
      entry("Page break", cmd.insertPageBreak);
      entry("Column break", cmd.insertColumnBreak);
      entry("Section break (next page)", cmd.insertSectionBreak("nextPage"));
      entry("Section break (continuous)", cmd.insertSectionBreak("continuous"));
      showBelow(menu, breakBtn);
    });
    action("Columns…", "Columns for this section or the selection", () => this.columnsDialog());
    btn("Line", "Insert horizontal line", cmd.insertRule);
    action("Equation", "Insert an equation", () => this.equationDialog(null));
    sep();
    action("Header & footer", "Edit the header and footer", () => this.headerFooterDialog());
    const pagesBtn = tb.icon("book-open", "Page view: show page breaks, headers and footers", () => {
      this.setPageView(!this.pageView);
      this.app.saveLocalStorage(PAGE_VIEW_KEY, this.pageView ? "on" : "off");
    });
    pagesBtn.toggleClass("is-active", this.pageView);
    tb.icon("search", "Find and replace (Ctrl+F)", () => this.openFind());
    const tocBtn = action("Contents ▾", "Table of contents", () => {
      const menu = new Menu();
      menu.addItem((i) => i.setTitle("Insert table of contents").onClick(() => {
        this.docx?.prepareToc();
        run(cmd.insertToc);
      }));
      menu.addItem((i) => i.setTitle("Update table of contents").onClick(() => {
        if (this.editor && !cmd.updateToc(this.editor.state, this.editor.dispatch)) new Notice("This document has no table of contents.");
      }));
      showBelow(menu, tocBtn);
    });
    action("Comment", "Comment on the selected text", () => void this.addComment());
    action("Footnote", "Insert a footnote at the cursor", () => void this.addFootnote());
    const reviewBtn = tb.icon("message-square", "Review panel: comments, footnotes and tracked changes", () => {
      this.setReview(!this.reviewOpen);
      reviewBtn.toggleClass("is-active", this.reviewOpen);
    });
    reviewBtn.toggleClass("is-active", this.reviewOpen);

    const end = tb.end(this.paperTheme, (t) => this.setPaperTheme(t), "paper");
    this.syncTheme = end.syncTheme;
    this.saveEl = end.status;
    this.wordsEl = createSpan({ cls: "dw-status docx-words" });
    end.status.parentElement?.insertBefore(this.wordsEl, end.status);
  }

  // ---------------------------------------------------------------------------
  // Word count, links, paragraph dialog
  // ---------------------------------------------------------------------------

  /** Font and Size show what's at the cursor: the text's own formatting, else its paragraph (and table) style. */
  private renderOutline(): void {
    const editor = this.editor, panel = this.outlineEl;
    if (!editor || !panel || !this.outlineOpen) return;
    panel.empty();
    panel.createDiv({ cls: "docx-review-title", text: "Document outline" });
    const entries = outlineEntries(editor.state.doc);
    if (!entries.length) panel.createDiv({ cls: "docx-review-empty", text: "Headings appear here." });
    const active = [...entries].reverse().find(e => e.pos <= editor.state.selection.from)?.pos;
    for (const entry of entries) {
      const button = panel.createEl("button", { text: entry.text, cls: "docx-outline-entry", attr: { title: entry.text } });
      button.style.paddingLeft = `${8 + (entry.level - 1) * 12}px`;
      button.toggleClass("is-active", entry.pos === active);
      button.addEventListener("click", () => {
        editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, entry.pos)));
        editor.focus();
        const heading = editor.nodeDOM(entry.pos - 1);
        if (heading instanceof HTMLElement && this.hostEl) {
          this.hostEl.scrollTop = Math.max(0, this.hostEl.scrollTop + heading.getBoundingClientRect().top - this.hostEl.getBoundingClientRect().top - 16);
        }
      });
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
    if (this.styleSelect) showValue(this.styleSelect, paraStyle ?? "Normal");
    showValue(this.fontSelect, font);
    showValue(this.sizeSelect, size === null ? null : String(size));
  }

  private scheduleWordCount(): void {
    if (this.wordsTimer) window.clearTimeout(this.wordsTimer);
    this.wordsTimer = window.setTimeout(() => this.updateWordCount(), 250);
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

  private setReview(open: boolean): void {
    this.reviewOpen = open;
    this.reviewEl?.toggle(open);
    this.app.saveLocalStorage(REVIEW_KEY, open ? "on" : "off");
    this.contentEl.querySelector('[aria-label^="Review panel"]')?.toggleClass("is-active", open);
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
    this.origins = read.origins;
    editor.updateState(EditorState.create({ doc: read.doc, plugins: editor.state.plugins }));
    this.markDirty();
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
    if (!panel || !editor || !docx || !this.reviewOpen) return;
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
        this.markDirty();
      });
      const row = c.createDiv({ cls: "docx-review-actions" });
      for (const [label, fn] of actions) row.createEl("button", { text: label }).addEventListener("click", fn);
    };

    // Comments
    panel.createDiv({ cls: "docx-review-title", text: "Comments" });
    const ranges = commentRanges(doc);
    const comments = docx.comments();
    if (!comments.length) panel.createDiv({ cls: "docx-review-empty", text: "Select text and use Comment to add one." });
    for (const c of comments.filter(c => !c.parentId)) {
      const thread = panel.createDiv({ cls: `docx-comment-thread${c.resolved ? " is-resolved" : ""}` });
      const range = ranges.get(c.id);
      const quote = range ? doc.textBetween(range.from, range.to, " ").slice(0, 60) : "";
      card(thread, `${c.resolved ? "Resolved · " : ""}${c.author}${c.date ? ` · ${new Date(c.date).toLocaleDateString()}` : ""}${quote ? ` · “${quote}”` : ""}`, c.text, (t) => docx.setCommentText(c.id, t), [
        ...(range ? ([["Go to", () => goTo(range.from, range.to)]] as [string, () => void][]) : []),
        [c.resolved ? "Reopen" : "Resolve", () => {
          docx.setCommentResolved(c.id, !c.resolved);
          this.markDirty();
          this.renderReview();
        }],
        ["Reply", () => void (async () => {
          const text = await promptText(this.app, "Reply to comment", "");
          if (!text?.trim()) return;
          const author = await this.author();
          if (!author || this.docx !== docx) return;
          docx.replyToComment(c.id, text.trim(), author);
          docx.setCommentResolved(c.id, false);
          this.markDirty();
          this.renderReview();
        })()],
        [
          "Delete",
          () => {
            docx.deleteComment(c.id);
            cmd.removeCommentAnchors(c.id)(editor.state, editor.dispatch);
            this.markDirty();
            this.renderReview();
          },
        ],
      ]);
      for (const reply of comments.filter(r => r.parentId === c.id)) {
        const replies = thread.createDiv({ cls: "docx-comment-reply" });
        card(replies, `${reply.author}${reply.date ? ` · ${new Date(reply.date).toLocaleDateString()}` : ""}`, reply.text,
          text => docx.setCommentText(reply.id, text), [["Delete reply", () => {
            docx.deleteComment(reply.id);
            this.markDirty();
            this.renderReview();
          }]]);
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
            this.markDirty();
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
    const changes = [
      ...revisions(docx.body).map((r) => ({ r, where: null as { root: Element; name: string } | null })),
      ...parts.flatMap((where) => revisions(where.root).map((r) => ({ r, where }))),
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
        this.markDirty();
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
    const kinds = { insert: "Inserted", delete: "Deleted", format: "Formatted" };
    for (const { r, where } of changes) {
      const c = panel.createDiv({ cls: `docx-review-card docx-change-${r.kind}` });
      c.createDiv({ cls: "docx-review-meta", text: `${kinds[r.kind]}${r.author ? ` by ${r.author}` : ""}${where ? ` · ${where.name}` : ""}` });
      c.createDiv({ cls: "docx-review-text", text: r.text.slice(0, 200) || "(formatting)" });
      const row = c.createDiv({ cls: "docx-review-actions" });
      row.createEl("button", { text: "Accept" }).addEventListener("click", () => decide(true, where, new Set([r.id])));
      row.createEl("button", { text: "Reject" }).addEventListener("click", () => decide(false, where, new Set([r.id])));
    }
  }

  // ---------------------------------------------------------------------------
  // Find and replace
  // ---------------------------------------------------------------------------

  private openFind(): void {
    if (!this.editor) return;
    if (this.findEl) {
      this.findEl.querySelector<HTMLInputElement>("input")?.focus();
      return;
    }
    const editor = this.editor;
    const bar = (this.findEl = createDiv({ cls: "pptx-find" }));
    this.contentEl.insertBefore(bar, this.hostEl?.parentElement ?? null);
    const input = bar.createEl("input", { type: "search", attr: { placeholder: "Find", "aria-label": "Find" } });
    const replace = bar.createEl("input", { type: "text", attr: { placeholder: "Replace with", "aria-label": "Replace with" } });
    const caseLabel = bar.createEl("label", { cls: "pptx-find-case", attr: { "aria-label": "Match case" } });
    const matchCase = caseLabel.createEl("input", { type: "checkbox" });
    caseLabel.appendText("Aa");
    const count = bar.createSpan({ cls: "pptx-find-count" });
    const state = () => findKey.getState(editor.state)!;
    const show = (current: number) => {
      editor.dispatch(setFind(editor.state.tr, input.value, matchCase.checked, current));
      const f = state();
      count.setText(input.value ? (f.matches.length ? `${f.current + 1} of ${f.matches.length}` : "No results") : "");
      editor.dom.querySelector(".docx-find-hit.is-current")?.scrollIntoView({ block: "center" });
    };
    const step = (dir: number) => {
      const n = state().matches.length;
      if (n) show((state().current + dir + n) % n);
    };
    const button = (text: string, title: string, fn: () => void) =>
      bar.createEl("button", { text, cls: "dw-btn", attr: { "aria-label": title } }).addEventListener("click", fn);
    button("↑", "Previous match (Shift+Enter)", () => step(-1));
    button("↓", "Next match (Enter)", () => step(1));
    button("Replace", "Replace this match", () => {
      const f = state();
      const hit = f.matches[f.current];
      if (!hit) return;
      replaceRangeCmd(hit, replace.value)(editor.state, editor.dispatch);
      show(f.current);
    });
    button("Replace all", "Replace every match", () => {
      const n = state().matches.length;
      replaceAllCmd(input.value, replace.value, matchCase.checked)(editor.state, editor.dispatch);
      new Notice(`Replaced ${n} ${n === 1 ? "match" : "matches"}.`);
      show(0);
    });
    button("×", "Close (Esc)", () => this.closeFind());
    input.addEventListener("input", () => show(0));
    matchCase.addEventListener("change", () => show(0));
    bar.addEventListener("keydown", (e) => {
      if (e.key === "Escape") this.closeFind();
      else if (e.key === "Enter" && e.target === input) {
        e.preventDefault();
        step(e.shiftKey ? -1 : 1);
      }
    });
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
    for (const part of HF_PARTS) {
      const root = this.docx?.headerFooter(...splitPart(part));
      if (root) this.hf.set(part, readDocx(root, this.hfReadContext(part)).doc);
    }
  }

  /** Images and links in a header / footer resolve through that part's own relationships. */
  private hfReadContext(part: HFPart) {
    const docx = this.docx;
    const path = docx?.headerFooterPath(...splitPart(part)) ?? undefined;
    return { imageSrc: (id: string) => this.hfUrls.get(part)?.get(id) ?? null, hyperlinkHref: (id: string) => docx?.hyperlinkHref(id, path) ?? null };
  }

  /** The header / footer a page shows (first page, even pages or the default), with page number fields filled in. */
  private renderHF(kind: HF, page: number, total: number): Node | null {
    const docx = this.docx;
    if (!docx) return null;
    const doc = this.hf.get(`${kind}:${headerFooterType(page, docx.differentFirstPage(), docx.differentOddEven())}`);
    if (!doc || !doc.textContent.trim() && !doc.content.size) return null;
    const frag = DOMSerializer.fromSchema(schema).serializeFragment(doc.content);
    const wrap = createDiv({ cls: `docx-hf docx-hf-${kind}` });
    wrap.appendChild(frag);
    if (this.ctx) styleRenderedParagraphs(wrap, doc, this.ctx.styles);
    wrap.querySelectorAll<HTMLElement>("[data-field]").forEach((f) => f.setText(String(f.dataset.field === "NUMPAGES" ? total : page)));
    return wrap;
  }

  private pagesPlugin(): Plugin {
    return pagesPlugin({
      setup: () => (this.pageView ? this.setup : null),
      pageSize: () => this.settings.pageSize,
      header: (page, total) => this.renderHF("header", page, total),
      footer: (page, total) => this.renderHF("footer", page, total),
      keep: (node) => this.ctx?.styles.keep(attrsOf<ParagraphAttrs>(node).style) ?? { next: false, lines: false },
    });
  }

  private headerFooterDialog(): void {
    const docx = this.docx;
    if (!docx || !this.ctx) return;
    // The dialog edits the body's final w:sectPr (header references, first page option), which the editor also keeps
    // as its doc attribute: bring the body up to date first, and the attribute after.
    const editor = this.editor;
    if (editor) writeBody(docx.body, editor.state.doc, this.origins, docx);
    const modal = new HeaderFooterModal(this.app, this.ctx, docx, (part) => this.hfReadContext(part), () => (this.suggesting ? this.settings.author : null), () => {
      this.readAllHF();
      const last = docx.body.lastElementChild;
      if (this.editor && last?.localName === "sectPr") {
        const sectPr = new XMLSerializer().serializeToString(last);
        if (sectPr !== this.editor.state.doc.attrs.sectPr) this.editor.dispatch(this.editor.state.tr.setDocAttribute("sectPr", sectPr).setMeta("addToHistory", false));
      }
      this.dirty = true;
      void this.saveNow();
      if (this.editor) repaginate(this.editor);
      this.renderReview();
    });
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
    this.contentEl.querySelector('[aria-label^="Page view"]')?.toggleClass("is-active", on);
    this.hostEl?.toggleClass("is-paged", on);
    if (this.editor) {
      refreshSectionWidths(this.editor);
      repaginate(this.editor);
    }
  }

  /** Save the document as a PDF, laid out on its pages as in the page view. */
  async exportPdf(): Promise<void> {
    if (!this.editor || !this.hostEl || !this.file || !this.setup) return;
    const paged = this.pageView;
    if (!paged) {
      this.setPageView(true);
      // ponytail: waits for the page layout's timer (60 ms, run twice when the header / footer size changes).
      await new Promise((r) => window.setTimeout(r, 400));
    }
    try {
      const { width, height } = this.setup;
      const pages = printPages(this.hostEl, height);
      await exportPdf(this.app, this.file, pages, { width, height }, "docx-theme-light");
    } finally {
      if (!paged) this.setPageView(false);
    }
  }

  private tableMenu(anchor: HTMLElement): void {
    const menu = new Menu();
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
    for (const [name, color] of [["Blue", "#EBF3FB"], ["Gray", "#F2F2F2"], ["Yellow", "#FFF8E1"], ["Green", "#E8F5E9"]]) {
      item(`Cell shading: ${name}`, cmd.setCellBackground(color));
    }
    item("Cell shading: none", cmd.setCellBackground(null));
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
    showBelow(menu, anchor);
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------

  private markDirty(): void {
    this.dirty = true;
    this.saveEl?.setText("Saving…");
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => void this.saveNow(), SAVE_DELAY_MS);
  }

  private saveNow(): Promise<void> {
    // Serialize saves: the package must not change while a previous save is zipping.
    this.saving = this.saving.then(async () => {
      if (!this.editor || !this.docx || !this.file || !this.dirty) return;
      const file = this.file;
      if (this.saveTimer) window.clearTimeout(this.saveTimer);
      this.dirty = false;
      try {
        writeBody(this.docx.body, this.editor.state.doc, this.origins, this.docx);
        const bytes = await this.docx.save();
        await this.app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
      } catch (err) {
        this.dirty = true;
        console.error("Failed to save .docx:", err);
        new Notice(`Failed to save ${file.name}: ${(err as Error).message}`);
      } finally {
        this.saveEl?.setText(this.dirty ? "Saving…" : "Saved");
      }
    });
    return this.saving;
  }
}

/** Edits the default header and footer, each in its own small editor. */
class HeaderFooterModal extends Modal {
  private editors = new Map<HFPart, { view: EditorView; origins: Origins; changed: boolean }>();
  private changed = false;

  constructor(
    app: App,
    private ctx: EditorContext,
    private docx: DocxDocument,
    private readContext: (part: HFPart) => { imageSrc: (id: string) => string | null; hyperlinkHref: (id: string) => string | null },
    /** The suggestion author while the document is in suggesting mode, else null. */
    private suggestAuthor: () => string | null,
    private onDone: () => void
  ) {
    super(app);
  }

  onOpen(): void {
    this.titleEl.setText("Header and footer");
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
    option("Different first page", this.docx.differentFirstPage(), (on) => this.docx.setDifferentFirstPage(on));
    option("Different odd & even", this.docx.differentOddEven(), (on) => this.docx.setDifferentOddEven(on));
    const types: HFType[] = ["default", ...(this.docx.differentFirstPage() ? ["first" as const] : []), ...(this.docx.differentOddEven() ? ["even" as const] : [])];
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
      const root = this.docx.headerFooter(...splitPart(part));
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
        if (!view.state.doc.textContent.trim() && !this.docx.headerFooter(kind, type)) continue;
        writeBody(this.docx.ensureHeaderFooter(kind, type), view.state.doc, origins, this.docx);
      }
      this.onDone();
    }
    this.editors.forEach((e) => e.view.destroy());
    this.contentEl.empty();
  }
}
