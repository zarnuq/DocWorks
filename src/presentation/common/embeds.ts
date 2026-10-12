/*
 * Embeds in notes: ![[report.docx]], ![[deck.pptx#3]] (slide 3) and ![[budget.xlsx#Sheet2]] (that sheet) show a
 * read-only preview inside the note, through Obsidian's embed registry (as its own PDF and canvas embeds do).
 * Clicking the preview's title opens the file in its editor.
 */
import { App, Component, Plugin, TFile } from "obsidian";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { IMAGE_MIME } from "../../core/utils/images";
import { readDocx } from "../../domain/docx/docx-reader";
import { DocxStyles } from "../../domain/docx/docx-styles";
import { parseSlide } from "../../domain/pptx/slide-model";
import { Edits } from "../../domain/xlsx/commands";
import { ConditionalFormats } from "../../domain/xlsx/conditional";
import { sheetPictures } from "../../domain/xlsx/drawing";
import { formulaSparklines, readSparklines } from "../../domain/xlsx/sparklines";
import { Workbook } from "../../domain/xlsx/workbook-model";
import { eotToTtf } from "../../infrastructure/fonts/eot";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { docxPlugins } from "../docx/docx-editor";
import { loadEmbeddedFonts } from "../pptx/embedded-fonts";
import { renderSlide } from "../pptx/slide-renderer";
import { Grid } from "../xlsx/grid";

/** Rows and columns an embedded sheet shows at most (the file opens for the rest). */
const MAX_ROWS = 60, MAX_COLS = 20;

/** "#3" as slide index 2; anything else (or nothing) is the first slide. */
export function slideIndex(subpath: string, count: number): number {
  const n = Number(subpath.replace(/^#/, ""));
  return Number.isInteger(n) && n >= 1 && n <= count ? n - 1 : 0;
}

const blobUrl = (bytes: Uint8Array, path: string) => {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return URL.createObjectURL(new Blob([bytes as BlobPart], { type: IMAGE_MIME[ext] ?? `image/${ext}` }));
};

/** A deck's slide, with previous / next buttons. Returns what to undo when the embed goes. */
export async function renderPptxEmbed(el: HTMLElement, bytes: ArrayBuffer | Uint8Array, subpath: string): Promise<() => void> {
  const doc = await PptxDocument.load(bytes);
  const urls = new Map<string, string>();
  for (const path of doc.mediaPaths()) if (IMAGE_MIME[path.slice(path.lastIndexOf(".") + 1).toLowerCase()]) urls.set(path, blobUrl(await doc.media(path), path));
  const fonts = await loadEmbeddedFonts(await doc.embeddedFonts(), eotToTtf).catch(() => ({ unload: () => undefined }));
  const count = doc.slides().length;
  let index = slideIndex(subpath, count);
  const stage = el.createDiv({ cls: "dw-embed-slide" });
  const nav = el.createDiv({ cls: "dw-embed-nav" });
  const prev = nav.createEl("button", { text: "‹", attr: { "aria-label": "Previous slide" } });
  const label = nav.createSpan();
  const next = nav.createEl("button", { text: "›", attr: { "aria-label": "Next slide" } });
  const draw = () => {
    stage.empty();
    const width = el.clientWidth || 640;
    renderSlide(stage, parseSlide(doc, doc.slides()[index].path), doc.slideSize, { width, editable: false, imageUrl: (p) => urls.get(p) ?? null, slideNumber: index + 1 });
    label.setText(`${index + 1} / ${count}`);
  };
  prev.addEventListener("click", () => { index = Math.max(0, index - 1); draw(); });
  next.addEventListener("click", () => { index = Math.min(count - 1, index + 1); draw(); });
  draw();
  // Redrawn at the note's width (the slide is laid out in pixels).
  let last = el.clientWidth;
  const resize = new ResizeObserver(() => {
    if (el.clientWidth && el.clientWidth !== last) {
      last = el.clientWidth;
      draw();
    }
  });
  resize.observe(el);
  return () => {
    resize.disconnect();
    fonts.unload();
    for (const url of urls.values()) URL.revokeObjectURL(url);
  };
}

/** A sheet's used area (up to MAX_ROWS × MAX_COLS) as the grid draws it, with sheet tabs when there are several. */
export async function renderXlsxEmbed(el: HTMLElement, bytes: ArrayBuffer | Uint8Array, subpath: string): Promise<() => void> {
  const edits = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
  const wb = edits.wb;
  const visible = wb.sheets.filter((s) => s.ref.state === "visible");
  const wanted = decodeURIComponent(subpath.replace(/^#/, ""));
  let sheet = visible.find((s) => s.name === wanted) ?? visible[0] ?? wb.sheets[0];
  const urls = new Map<string, string>();
  const area = el.createDiv({ cls: "dw-embed-sheet" });
  const tabs = visible.length > 1 ? el.createDiv({ cls: "dw-embed-tabs" }) : null;
  const draw = async () => {
    area.empty();
    const used = sheet.used();
    if (!used) return void area.createDiv({ cls: "dw-muted", text: "This sheet is empty." });
    const pictures = sheetPictures(wb, sheet);
    for (const p of pictures) if (!urls.has(p.image)) urls.set(p.image, blobUrl(await wb.pkg.media(p.image).catch(() => new Uint8Array()), p.image));
    const cf = new ConditionalFormats(sheet, (r, c, f) => edits.calc.evaluate(sheet, r, c, f));
    const sparks = [...readSparklines(sheet), ...formulaSparklines(sheet, (s, r, c, f) => edits.calc.evaluateArray(s, r, c, f))];
    const painter = new Grid(createDiv(), {
      sheet: () => sheet,
      pictures: () => pictures,
      imageUrl: (path) => urls.get(path) ?? null,
      conditional: (row, col) => (cf.empty ? null : cf.at(row, col)),
      sparklines: () => sparks,
      isStale: () => false,
      isCircular: () => false,
      commit: () => false,
      selectionChanged: () => undefined,
      editing: () => undefined,
      clear: () => undefined,
      fill: () => undefined,
      resize: () => undefined,
      autofit: () => undefined,
      contextMenu: () => undefined,
      filterMenu: () => undefined,
    });
    const r = { r1: 0, c1: 0, r2: Math.min(used.r2, MAX_ROWS - 1), c2: Math.min(used.c2, MAX_COLS - 1) };
    // The painted cells are positioned absolutely; the box takes their size so the frame scrolls over them.
    const grid = area.createDiv({ cls: "dw-embed-grid" });
    grid.setCssProps({ "--dw-grid-w": `${sheet.colX(r.c2 + 1)}px`, "--dw-grid-h": `${sheet.rowY(r.r2 + 1)}px` });
    painter.drawRange(grid, r);
    tabs?.querySelectorAll("button").forEach((b) => b.toggleClass("is-active", b.textContent === sheet.name));
  };
  for (const s of visible) tabs?.createEl("button", { text: s.name }).addEventListener("click", () => {
    sheet = s;
    void draw();
  });
  await draw();
  return () => {
    for (const url of urls.values()) URL.revokeObjectURL(url);
  };
}

/** The document's text flow, read-only (no page view: the note is the page). */
export async function renderDocxEmbed(el: HTMLElement, bytes: ArrayBuffer | Uint8Array): Promise<() => void> {
  const docx = await DocxDocument.load(bytes);
  const urls = await docx.imageUrls();
  const { doc } = readDocx(docx.body, { imageSrc: (id) => urls.get(id) ?? null, hyperlinkHref: (id) => docx.hyperlinkHref(id) });
  const styles = new DocxStyles(docx.stylesXml(), docx.numberingXml(), docx.themeXml());
  const host = el.createDiv({ cls: "docx-editor-host dw-embed-doc" });
  const view = new EditorView(host, {
    state: EditorState.create({ doc, plugins: docxPlugins({ styles, ensureList: () => "", listKind: () => null }) }),
    editable: () => false,
  });
  return () => {
    view.destroy();
    for (const url of urls.values()) URL.revokeObjectURL(url);
  };
}

const RENDER = { docx: renderDocxEmbed, pptx: renderPptxEmbed, xlsx: renderXlsxEmbed } as const;
type Kind = keyof typeof RENDER;

class OfficeEmbed extends Component {
  private cleanup: (() => void) | null = null;

  constructor(private app: App, private containerEl: HTMLElement, private file: TFile, private subpath: string, private kind: Kind) {
    super();
  }

  /** Called by Obsidian once the embed is attached. */
  async loadFile(): Promise<void> {
    const el = this.containerEl;
    el.empty();
    el.addClass("dw-embed");
    const title = el.createDiv({ cls: "dw-embed-title", text: this.file.basename });
    title.addEventListener("click", () => void this.app.workspace.getLeaf("tab").openFile(this.file));
    const body = el.createDiv({ cls: `dw-embed-body dw-embed-${this.kind}` });
    try {
      this.cleanup = await RENDER[this.kind](body, await this.app.vault.readBinary(this.file), this.subpath);
    } catch (err) {
      body.setText(`Couldn't show this file: ${(err as Error).message}`);
    }
  }

  onunload(): void {
    this.cleanup?.();
    this.cleanup = null;
  }
}

interface EmbedRegistry {
  isExtensionRegistered(ext: string): boolean;
  registerExtension(ext: string, create: (ctx: { containerEl: HTMLElement }, file: TFile, subpath: string) => Component): void;
  unregisterExtension(ext: string): void;
}

/** Register embeds for the file types Doc Works opens (the registry isn't in Obsidian's typings; skipped if absent). */
export function registerEmbeds(plugin: Plugin, kinds: Kind[]): void {
  const registry = (plugin.app as unknown as { embedRegistry?: EmbedRegistry }).embedRegistry;
  if (!registry) return;
  for (const kind of kinds) {
    if (registry.isExtensionRegistered(kind)) continue;
    registry.registerExtension(kind, (ctx, file, subpath) => new OfficeEmbed(plugin.app, ctx.containerEl, file, subpath, kind));
    plugin.register(() => registry.unregisterExtension(kind));
  }
}
