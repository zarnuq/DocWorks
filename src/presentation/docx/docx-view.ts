import { FileView, Menu, Notice, TFile, WorkspaceLeaf } from "obsidian";
import { Command, EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { toggleMark } from "prosemirror-commands";
import { redo, undo } from "prosemirror-history";
import { NS } from "../../core/utils/xml";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { Origins, readDocx } from "../../domain/docx/docx-reader";
import { writeBody } from "../../domain/docx/docx-writer";
import { DocxStyles } from "../../domain/docx/docx-styles";
import { docxSchema as schema, Align } from "../../domain/docx/schema";
import { FONTS, PaperTheme, promptText, Toolbar } from "../common/toolbar";
import { cmd, docxPlugins, EditorContext } from "./docx-editor";
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

/** Page content width in px from the final w:sectPr (twips / 15), Letter with 1" margins by default. */
function contentWidthPx(body: Element): number {
  const sectPr = Array.from(body.children).filter((c) => c.localName === "sectPr").pop();
  const attr = (local: string, name: string, fallback: number) =>
    Number(sectPr?.getElementsByTagNameNS(NS.w, local)[0]?.getAttributeNS(NS.w, name) ?? fallback);
  return Math.round((attr("pgSz", "w", 12240) - attr("pgMar", "left", 1440) - attr("pgMar", "right", 1440)) / 15);
}

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

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
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
    this.pageWidth = contentWidthPx(docx.body);

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
    const host = this.contentEl.createDiv({ cls: "docx-editor-host" });
    host.style.setProperty("--docx-page-width", `${this.pageWidth}px`);

    this.editor = new EditorView(host, {
      state: EditorState.create({ doc, plugins: docxPlugins(ctx) }),
      nodeViews: {
        image: (node, view, getPos) =>
          new DocxImageView(node, view, getPos, {
            replace: (pos) => this.pickImage((f) => void this.replaceImage(pos, f)),
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
      },
      handleDOMEvents: {
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
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
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

  private pickImage(onPick: (file: File) => void): void {
    const input = document.body.createEl("input", { type: "file", attr: { accept: "image/*" } });
    input.setCssStyles({ display: "none" });
    input.onchange = () => {
      const file = input.files?.[0];
      if (file) onPick(file);
      input.remove();
    };
    input.click();
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
    for (const [label, align] of [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Justify", "justify"]] as [string, Align][]) {
      btn(label, `Align ${align}`, cmd.setAlign(align));
    }
    tb.select("Line spacing", "Spacing", [["1.0", "1"], ["1.15", "1.15"], ["1.5", "1.5"], ["2.0", "2"]], (v) => run(cmd.setLineSpacing(v)));
    btn("Outdent", "Decrease indent / list level (Shift+Tab)", cmd.shiftLevel(-1));
    btn("Indent", "Increase indent / list level (Tab)", cmd.shiftLevel(1));
    btn("• List", "Bulleted list", cmd.toggleList(ctx, "bullet"));
    btn("1. List", "Numbered list", cmd.toggleList(ctx, "ordered"));
    sep();
    action("Table ▾", "Table", () => this.tableMenu());
    action("Image", "Insert image", () => this.pickImage((f) => void this.insertImage(f)));
    btn("Page break", "Insert page break (Ctrl+Enter)", cmd.insertPageBreak);
    btn("Line", "Insert horizontal line", cmd.insertRule);

    const end = tb.end(this.paperTheme, (t) => this.setPaperTheme(t), "paper");
    this.syncTheme = end.syncTheme;
    this.saveEl = end.status;
  }

  private tableMenu(): void {
    const menu = new Menu();
    const item = (title: string, command: Command) =>
      menu.addItem((i) =>
        i.setTitle(title).onClick(() => {
          if (!this.editor) return;
          if (!command(this.editor.state, this.editor.dispatch)) new Notice("Place the cursor inside a table first.");
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
