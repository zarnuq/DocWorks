/*
 * Office files embedded in a note, made into what a Word document can hold when the note is exported: a sheet (or
 * range) becomes a table of its shown values, a slide a picture (captured as for slide PNG export, desktop only) and a
 * Word file its own content (through DOCX → Markdown, rendered like an embedded note, pictures kept).
 */
import { App, Component, MarkdownRenderer, TFile } from "obsidian";
import { IMAGE_MIME } from "../../core/utils/images";
import { docxToMarkdown } from "../../domain/docx/to-markdown";
import { parseSlide } from "../../domain/pptx/slide-model";
import { rangeRows } from "../../domain/xlsx/csv";
import { Workbook } from "../../domain/xlsx/workbook-model";
import { eotToTtf } from "../../infrastructure/fonts/eot";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { loadEmbeddedFonts } from "../pptx/embedded-fonts";
import { renderSlide } from "../pptx/slide-renderer";
import { embedRange, sectionOf, sheetTarget, slideIndex } from "./embeds";
import { readDocx } from "../../domain/docx/docx-reader";
import { PNG_WIDTH, renderPng } from "./png";

export type ExportPicture = { bytes: Uint8Array; ext: string; width: number; height: number };
export const OFFICE_EXTENSIONS = new Set(["docx", "pptx", "xlsx"]);

/** Folder name the embedded document's pictures are given in its Markdown, to find them again. */
const PICTURES = "dw-embedded-pictures";

/** What replaces the embed of `file` in the exported note; null when the file can't be shown (it stays a link). */
export async function officeEmbedForExport(app: App, file: TFile, subpath: string, component: Component, pictures: Map<HTMLImageElement, ExportPicture>): Promise<HTMLElement | null> {
  const bytes = await app.vault.readBinary(file);
  const ext = file.extension.toLowerCase();
  if (ext === "xlsx") return sheetTable(bytes, subpath);
  if (ext === "pptx") return slidePicture(bytes, subpath, file.basename, pictures);
  if (ext === "docx") return documentContent(app, bytes, file, subpath, component, pictures);
  return null;
}

async function sheetTable(bytes: ArrayBuffer, subpath: string): Promise<HTMLElement | null> {
  const wb = Workbook.open(await XlsxWorkbook.load(bytes));
  const target = sheetTarget(subpath, wb.sheets.map((s) => s.name));
  const sheet = wb.sheets.find((s) => s.name === target.sheet) ?? wb.sheets.find((s) => s.ref.state === "visible") ?? wb.sheets[0];
  const r = sheet && embedRange(sheet, target.range);
  const t = r && rangeRows(sheet, r);
  if (!t) return null;
  const table = createEl("table");
  const head = table.createEl("thead").createEl("tr");
  for (const text of t.rows[0]) head.createEl("th", { text });
  const body = table.createEl("tbody");
  for (const row of t.rows.slice(1)) {
    const tr = body.createEl("tr");
    for (const text of row) tr.createEl("td", { text });
  }
  return table;
}

async function slidePicture(bytes: ArrayBuffer, subpath: string, name: string, pictures: Map<HTMLImageElement, ExportPicture>): Promise<HTMLElement | null> {
  const doc = await PptxDocument.load(bytes);
  const index = slideIndex(subpath, doc.slides().length);
  const urls = new Map<string, string>();
  for (const path of doc.mediaPaths()) {
    const type = IMAGE_MIME[path.slice(path.lastIndexOf(".") + 1).toLowerCase()];
    if (type) urls.set(path, URL.createObjectURL(new Blob([(await doc.media(path)) as BlobPart], { type })));
  }
  const fonts = await loadEmbeddedFonts(await doc.embeddedFonts(), eotToTtf).catch(() => ({ unload: () => undefined }));
  try {
    const { cx, cy } = doc.slideSize;
    const width = Math.floor(Math.min(PNG_WIDTH, window.innerWidth || PNG_WIDTH, ((window.innerHeight || PNG_WIDTH) * cx) / cy));
    const height = Math.floor((width * cy) / cx);
    const png = await renderPng({ width, height }, (host) =>
      renderSlide(host, parseSlide(doc, doc.slides()[index].path), doc.slideSize, { width, editable: false, imageUrl: (p) => urls.get(p) ?? null, slideNumber: index + 1 })
    );
    if (!png) return null;
    const img = createEl("img", { attr: { alt: `${name}, slide ${index + 1}` } });
    pictures.set(img, { bytes: png, ext: "png", width, height });
    return img;
  } finally {
    fonts.unload();
    for (const url of urls.values()) URL.revokeObjectURL(url);
  }
}

async function documentContent(app: App, bytes: ArrayBuffer, file: TFile, subpath: string, component: Component, pictures: Map<HTMLImageElement, ExportPicture>): Promise<HTMLElement> {
  const docx = await DocxDocument.load(bytes);
  // ![[report.docx#Results]]: just that heading's section.
  const section = subpath ? sectionOf(readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: (id) => docx.hyperlinkHref(id) }).doc, subpath) : undefined;
  const result = await docxToMarkdown(docx, PICTURES, section);
  const el = createDiv();
  await MarkdownRenderer.render(app, result.markdown, el, file.path, component);
  // Its pictures aren't in the vault: put the extracted bytes behind the pictures the Markdown points at.
  for (const pic of Array.from(el.querySelectorAll<HTMLElement>("img, .internal-embed"))) {
    if (pic.parentElement?.closest(".internal-embed")) continue;
    const src = decodeURIComponent(pic.getAttribute("src") ?? "");
    if (!src.includes(`${PICTURES}/`)) continue;
    const image = result.images.find((i) => src.endsWith(`/${i.name}`));
    const ext = image?.name.slice(image.name.lastIndexOf(".") + 1).toLowerCase() ?? "";
    const size = image && IMAGE_MIME[ext] ? await imageSize(image.bytes, IMAGE_MIME[ext]) : null;
    if (!image || !size) {
      pic.remove();
      continue;
    }
    const img = createEl("img", { attr: { alt: pic.getAttribute("alt") ?? "" } });
    pictures.set(img, { bytes: image.bytes, ext, ...size });
    pic.replaceWith(img);
  }
  return el;
}

async function imageSize(bytes: Uint8Array, type: string): Promise<{ width: number; height: number } | null> {
  try {
    const bitmap = await createImageBitmap(new Blob([bytes as BlobPart], { type }));
    const size = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return size;
  } catch {
    return null;
  }
}
