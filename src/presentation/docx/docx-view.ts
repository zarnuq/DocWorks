import { App, FileView, Menu, Modal, Notice, Scope, Setting, TFile, WorkspaceLeaf } from "obsidian";
import { Command, EditorState, Plugin, TextSelection } from "prosemirror-state";
import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { toggleMark } from "prosemirror-commands";
import { redo, undo } from "prosemirror-history";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { Origins, readDocx } from "../../domain/docx/docx-reader";
import { writeBody } from "../../domain/docx/docx-writer";
import { DocxStyles } from "../../domain/docx/docx-styles";
import { docxSchema as schema, Align } from "../../domain/docx/schema";
import { FONTS, PaperTheme, promptText, Toolbar } from "../common/toolbar";
import { pickImage } from "../common/images";
import { cmd, commentRanges, docxPlugins, EditorContext, footnoteIds, linkHref, pageNumberNode, pagesPlugin, repaginate } from "./docx-editor";
import { acceptRevisions, rejectRevisions, revisions } from "../../domain/docx/revisions";
import { countWords, findKey, replaceAllCmd, replaceRangeCmd, setFind } from "./docx-find";
import { pageSetup, PageSetup } from "../../domain/docx/pages";
import { attrsOf, ParagraphAttrs } from "../../domain/docx/schema";
import { DocxImageView } from "./docx-image-view";

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

const PAGE_VIEW_KEY = "doc-works-page-view";
const REVIEW_KEY = "doc-works-review-panel";
const AUTHOR_KEY = "doc-works-author";
type HF = "header" | "footer";

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
  private pageView = window.localStorage.getItem(PAGE_VIEW_KEY) !== "off";
  private hostEl: HTMLElement | null = null;
  private wordsEl: HTMLElement | null = null;
  private wordsTimer: number | null = null;
  private findEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  /** Header and footer content, for drawing them in the page view. */
  private hf: Record<HF, PMNode | null> = { header: null, footer: null };
  private reviewEl: HTMLElement | null = null;
  private reviewOpen = window.localStorage.getItem(REVIEW_KEY) === "on";
  private reviewTimer: number | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
    this.scope = new Scope(this.app.scope);
    this.scope.register(["Mod"], "f", () => {
      this.openFind();
      return false;
    });
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
    this.setup = pageSetup(docx.body);
    this.pageWidth = this.setup.width - this.setup.left - this.setup.right;
    this.hf = { header: this.readHF("header"), footer: this.readHF("footer") };

    const ctx: EditorContext = {
      styles: new DocxStyles(docx.stylesXml(), docx.numberingXml()),
      // A new list may create numbering.xml, so refresh the display styles afterwards.
      ensureList: (kind) => {
        const numId = docx.ensureList(kind);
        ctx.styles = new DocxStyles(docx.stylesXml(), docx.numberingXml());
        return numId;
      },
      listKind: (numId) => {
        const fmt = docx.listFormat(numId, 0);
        return fmt ? (fmt === "bullet" ? "bullet" : "ordered") : null;
      },
    };

    const toolbar = this.contentEl.createDiv({ cls: "dw-toolbar" });
    const body = this.contentEl.createDiv({ cls: "docx-body" });
    const host = (this.hostEl = body.createDiv({ cls: "docx-editor-host" }));
    this.reviewEl = body.createDiv({ cls: "docx-review" });
    this.reviewEl.toggle(this.reviewOpen);
    const st = this.setup;
    host.setCssProps({
      "--docx-page-width": `${this.pageWidth}px`,
      "--docx-m-top": `${st.top}px`,
      "--docx-m-right": `${st.right}px`,
      "--docx-m-bottom": `${st.bottom}px`,
      "--docx-m-left": `${st.left}px`,
      "--docx-hf-header": `${st.header}px`,
      "--docx-hf-footer": `${st.footer}px`,
    });
    host.toggleClass("is-paged", this.pageView);
    this.ctx = ctx;

    this.editor = new EditorView(host, {
      state: EditorState.create({ doc, plugins: [...docxPlugins(ctx, { "Mod-k": () => (void this.editLink(), true) }), this.pagesPlugin()] }),
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
      },
      dispatchTransaction: (tr) => {
        if (!this.editor) return;
        this.editor.updateState(this.editor.state.apply(tr));
        if (tr.docChanged) this.markDirty();
        if (tr.docChanged || tr.selectionSet) this.scheduleWordCount();
        if (tr.docChanged) this.scheduleReview();
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
    this.resizeObserver = new ResizeObserver(() => this.editor && repaginate(this.editor));
    this.resizeObserver.observe(host);
    this.updateWordCount();
    this.renderReview();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
    this.resizeObserver?.disconnect();
    this.closeFind();
    for (const t of [this.wordsTimer, this.reviewTimer]) if (t) window.clearTimeout(t);
    await this.requestSave();
    await this.saving;
    this.editor?.destroy();
    this.editor = null;
    this.docx = null;
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

    btn("↶", "Undo (Ctrl+Z)", undo);
    btn("↷", "Redo (Ctrl+Y)", redo);
    sep();
    tb.select("Paragraph style", "Style", STYLES, (v) => run(cmd.setStyle(v === "Normal" ? null : v)));
    tb.select("Font", "Font", FONTS.map((f) => [f, f]), (font) => run(cmd.setMark(m.font, { font })));
    tb.select("Font size", "Size", FONT_SIZES.map((s) => [`${s} pt`, String(s)]), (v) => run(cmd.setMark(m.size, { pt: Number(v) })));
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
    tb.select("Line spacing", "Spacing", [["1.0", "1"], ["1.15", "1.15"], ["1.5", "1.5"], ["2.0", "2"]], (v) => run(cmd.setLineSpacing(v)));
    tb.icon("list", "Bulleted list", () => run(cmd.toggleList(ctx, "bullet")));
    tb.icon("list-ordered", "Numbered list", () => run(cmd.toggleList(ctx, "ordered")));
    tb.icon("outdent", "Decrease indent / list level (Shift+Tab)", () => run(cmd.shiftLevel(-1)));
    tb.icon("indent", "Increase indent / list level (Tab)", () => run(cmd.shiftLevel(1)));
    action("Paragraph…", "Spacing and indents", () => this.paragraphDialog());
    tb.icon("link", "Add, change or remove a link (Ctrl+K; Ctrl+click follows it)", () => void this.editLink());
    sep();
    action("Table ▾", "Table", () => this.tableMenu());
    action("Image", "Insert image", () => pickImage((f) => void this.insertImage(f)));
    btn("Page break", "Insert page break (Ctrl+Enter)", cmd.insertPageBreak);
    btn("Line", "Insert horizontal line", cmd.insertRule);
    sep();
    action("Header & footer", "Edit the header and footer", () => this.headerFooterDialog());
    const pagesBtn = tb.icon("book-open", "Page view: show page breaks, headers and footers", () => {
      this.pageView = !this.pageView;
      window.localStorage.setItem(PAGE_VIEW_KEY, this.pageView ? "on" : "off");
      pagesBtn.toggleClass("is-active", this.pageView);
      this.hostEl?.toggleClass("is-paged", this.pageView);
      if (this.editor) repaginate(this.editor);
    });
    pagesBtn.toggleClass("is-active", this.pageView);
    tb.icon("search", "Find and replace (Ctrl+F)", () => this.openFind());
    action("Contents ▾", "Table of contents", () => {
      const menu = new Menu();
      menu.addItem((i) => i.setTitle("Insert table of contents").onClick(() => {
        this.docx?.prepareToc();
        run(cmd.insertToc);
      }));
      menu.addItem((i) => i.setTitle("Update table of contents").onClick(() => {
        if (this.editor && !cmd.updateToc(this.editor.state, this.editor.dispatch)) new Notice("This document has no table of contents.");
      }));
      const r = this.contentEl.querySelector(".dw-toolbar")!.getBoundingClientRect();
      menu.showAtPosition({ x: r.left + r.width / 2, y: r.bottom });
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

  private async author(): Promise<string | null> {
    const saved = window.localStorage.getItem(AUTHOR_KEY);
    if (saved) return saved;
    const name = await promptText(this.app, "Your name (shown on comments)", "");
    if (!name?.trim()) return null;
    window.localStorage.setItem(AUTHOR_KEY, name.trim());
    return name.trim();
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
    window.localStorage.setItem(REVIEW_KEY, open ? "on" : "off");
    this.contentEl.querySelector('[aria-label^="Review panel"]')?.toggleClass("is-active", open);
    this.renderReview();
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
    for (const c of comments) {
      const range = ranges.get(c.id);
      const quote = range ? doc.textBetween(range.from, range.to, " ").slice(0, 60) : "";
      card(panel, `${c.author}${c.date ? ` · ${new Date(c.date).toLocaleDateString()}` : ""}${quote ? ` · “${quote}”` : ""}`, c.text, (t) => docx.setCommentText(c.id, t), [
        ...(range ? ([["Go to", () => goTo(range.from, range.to)]] as [string, () => void][]) : []),
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

    // Tracked changes (from the saved XML plus the current edits)
    panel.createDiv({ cls: "docx-review-title", text: "Tracked changes" });
    writeBody(docx.body, doc, this.origins, docx);
    const changes = revisions(docx.body);
    if (!changes.length) {
      panel.createDiv({ cls: "docx-review-empty", text: "No tracked changes." });
      return;
    }
    const all = panel.createDiv({ cls: "docx-review-actions" });
    all.createEl("button", { text: "Accept all" }).addEventListener("click", () => this.rewriteBody((b) => acceptRevisions(b)));
    all.createEl("button", { text: "Reject all" }).addEventListener("click", () => this.rewriteBody((b) => rejectRevisions(b)));
    const kinds = { insert: "Inserted", delete: "Deleted", format: "Formatted" };
    for (const r of changes) {
      const c = panel.createDiv({ cls: `docx-review-card docx-change-${r.kind}` });
      c.createDiv({ cls: "docx-review-meta", text: `${kinds[r.kind]}${r.author ? ` by ${r.author}` : ""}` });
      c.createDiv({ cls: "docx-review-text", text: r.text.slice(0, 200) || "(formatting)" });
      const row = c.createDiv({ cls: "docx-review-actions" });
      row.createEl("button", { text: "Accept" }).addEventListener("click", () => this.rewriteBody((b) => acceptRevisions(b, new Set([r.id]))));
      row.createEl("button", { text: "Reject" }).addEventListener("click", () => this.rewriteBody((b) => rejectRevisions(b, new Set([r.id]))));
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

  private readHF(kind: HF): PMNode | null {
    const root = this.docx?.headerFooter(kind);
    return root ? readDocx(root, { imageSrc: () => null, hyperlinkHref: () => null }).doc : null;
  }

  /** Header / footer content for the page view, with page number fields filled in. */
  private renderHF(kind: HF, page: number, total: number): Node | null {
    const doc = this.hf[kind];
    if (!doc || !doc.textContent.trim() && !doc.content.size) return null;
    const frag = DOMSerializer.fromSchema(schema).serializeFragment(doc.content);
    const wrap = createDiv({ cls: `docx-hf docx-hf-${kind}` });
    wrap.appendChild(frag);
    wrap.querySelectorAll<HTMLElement>("[data-field]").forEach((f) => f.setText(String(f.dataset.field === "NUMPAGES" ? total : page)));
    return wrap;
  }

  private pagesPlugin(): Plugin {
    return pagesPlugin({
      setup: () => (this.pageView ? this.setup : null),
      header: (page, total) => this.renderHF("header", page, total),
      footer: (page, total) => this.renderHF("footer", page, total),
    });
  }

  private headerFooterDialog(): void {
    const docx = this.docx;
    if (!docx || !this.ctx) return;
    const modal = new HeaderFooterModal(this.app, this.ctx, docx, () => {
      this.hf = { header: this.readHF("header"), footer: this.readHF("footer") };
      this.dirty = true;
      void this.saveNow();
      if (this.editor) repaginate(this.editor);
    });
    modal.open();
  }

  private tableMenu(): void {
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
    menu.addSeparator();
    item("Delete row", cmd.deleteRow);
    item("Delete column", cmd.deleteColumn);
    item("Delete table", cmd.deleteTable);
    const rect = this.contentEl.querySelector(".dw-toolbar")!.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left + rect.width / 2, y: rect.bottom });
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
  private editors: { kind: HF; view: EditorView; origins: Origins }[] = [];
  private changed = false;

  constructor(app: App, private ctx: EditorContext, private docx: DocxDocument, private onDone: () => void) {
    super(app);
  }

  onOpen(): void {
    this.titleEl.setText("Header and footer");
    this.modalEl.addClass("docx-hf-modal");
    for (const kind of ["header", "footer"] as const) {
      new Setting(this.contentEl).setName(kind === "header" ? "Header" : "Footer").setHeading();
      const root = this.docx.headerFooter(kind);
      const { doc, origins } = root ? readDocx(root, { imageSrc: () => null, hyperlinkHref: (id) => this.docx.hyperlinkHref(id) }) : { doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create()), origins: new Map() as Origins };
      const bar = this.contentEl.createDiv({ cls: "docx-hf-tools" });
      const host = this.contentEl.createDiv({ cls: "docx-hf-editor" });
      const view: EditorView = new EditorView(host, {
        state: EditorState.create({ doc, plugins: docxPlugins(this.ctx) }),
        dispatchTransaction: (tr) => {
          view.updateState(view.state.apply(tr));
          if (tr.docChanged) this.changed = true;
        },
      });
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
      this.editors.push({ kind, view, origins });
    }
  }

  onClose(): void {
    if (this.changed) {
      for (const { kind, view, origins } of this.editors) {
        const empty = !view.state.doc.textContent.trim() && !this.docx.headerFooter(kind);
        if (empty) continue;
        writeBody(this.docx.ensureHeaderFooter(kind), view.state.doc, origins, this.docx);
      }
      this.onDone();
    }
    this.editors.forEach((e) => e.view.destroy());
    this.contentEl.empty();
  }
}
