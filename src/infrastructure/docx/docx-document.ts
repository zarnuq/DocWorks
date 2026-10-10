import JSZip from "jszip";
import { NS, childNS } from "../../core/utils/xml";
import { IMAGE_MIME } from "../../core/utils/images";
import { OoxmlPackage, parseXml } from "../ooxml/ooxml-package";
import { wVal } from "../../domain/docx/docx-reader";
import { numberingLevel } from "../../domain/docx/docx-styles";

export type ListKind = "bullet" | "ordered";

const DEFAULT_MAIN = "word/document.xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const BULLETS = ["•", "◦", "▪"];
const ORDERED_FORMATS = ["decimal", "lowerLetter", "lowerRoman"];


function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

/** Nine list levels, 0.5" apart, cycling bullet glyphs or decimal / letter / roman. */
function abstractNumXml(id: number, kind: ListKind): string {
  const levels = Array.from({ length: 9 }, (_, i) => {
    const fmt = kind === "bullet" ? "bullet" : ORDERED_FORMATS[i % 3];
    const text = kind === "bullet" ? BULLETS[i % 3] : `%${i + 1}.`;
    return (
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
    );
  }).join("");
  return `<w:abstractNum xmlns:w="${NS.w}" w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`;
}

const heading = (lvl: number, sz: number, extra = "") =>
  [`heading ${lvl + 1}`, `<w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="${lvl}"/>`, `<w:b/>${extra}<w:sz w:val="${sz}"/>`] as const;

/** Word's built-in paragraph styles the editor can apply: [name, pPr, rPr]. */
const BUILTIN_STYLES: Record<string, readonly [string, string, string]> = {
  Title: ["Title", `<w:spacing w:after="120"/>`, `<w:sz w:val="56"/>`],
  Heading1: heading(0, 32),
  Heading2: heading(1, 26),
  Heading3: heading(2, 24),
  Heading4: heading(3, 22, "<w:i/>"),
  Heading5: heading(4, 22),
  Heading6: heading(5, 22, "<w:i/>"),
  Quote: ["Quote", `<w:ind w:left="720" w:right="720"/>`, `<w:i/>`],
  TOCHeading: ["TOC Heading", `<w:keepNext/><w:spacing w:before="240" w:after="120"/>`, `<w:b/><w:sz w:val="28"/>`],
  TOC1: ["toc 1", `<w:spacing w:after="100"/>`, ``],
  TOC2: ["toc 2", `<w:spacing w:after="100"/><w:ind w:left="220"/>`, ``],
  TOC3: ["toc 3", `<w:spacing w:after="100"/><w:ind w:left="440"/>`, ``],
};

/** Tail of CT_Settings' child order, from the settings we add onward. */
const SETTINGS_ORDER = ["evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"];

/** CT_SectPr child order. */
const SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];

/** Which header / footer of a section: every page, the first page, or even pages. */
export type HFType = "default" | "first" | "even";

const nextId = (part: Document, tag: string) =>
  Array.from(part.getElementsByTagNameNS(NS.w, tag)).reduce((m, e) => Math.max(m, Number(wVal(e, "id")) || 0), 0) + 1;

/** Plain text of a comment or footnote: one line per paragraph. */
function noteText(el: Element): string {
  return Array.from(el.getElementsByTagNameNS(NS.w, "p"))
    .map((p) => Array.from(p.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join(""))
    .join("\n")
    .replace(/^\s+/, "");
}

/**
 * Replace a comment's or footnote's text, one paragraph per line. The first
 * paragraph's properties and its reference mark run (w:annotationRef /
 * w:footnoteRef) are kept.
 */
function setNoteText(el: Element, text: string): void {
  const doc = el.ownerDocument;
  const first = el.getElementsByTagNameNS(NS.w, "p")[0];
  const pPr = first && childNS(first, NS.w, "pPr");
  const mark = first && Array.from(first.children).find((r) => r.localName === "r" && Array.from(r.children).some((c) => /^(annotationRef|footnoteRef|endnoteRef)$/.test(c.localName)));
  el.replaceChildren();
  text.split("\n").forEach((line, i) => {
    const p = el.appendChild(doc.createElementNS(NS.w, "w:p"));
    if (pPr) p.appendChild(pPr.cloneNode(true));
    if (i === 0 && mark) {
      p.appendChild(mark.cloneNode(true));
      if (mark.getElementsByTagNameNS(NS.w, "footnoteRef").length) line = ` ${line}`;
    }
    if (!line) return;
    const t = p.appendChild(doc.createElementNS(NS.w, "w:r")).appendChild(doc.createElementNS(NS.w, "w:t"));
    t.setAttributeNS(NS.xml, "xml:space", "preserve");
    t.textContent = line;
  });
}

/**
 * A .docx package. The main document part stays a live DOM; the editor rewrites
 * only <w:body>, so headers, footers, styles, comments etc. pass through.
 */
export class DocxDocument extends OoxmlPackage {
  mainPath = DEFAULT_MAIN;

  static async load(data: ArrayBuffer | Uint8Array): Promise<DocxDocument> {
    const zip = await JSZip.loadAsync(data);
    const doc = new DocxDocument(zip);
    await doc.readParts();
    doc.mainPath = doc.relOfType("", "officeDocument") ?? DEFAULT_MAIN;
    const root = doc.parts.get(doc.mainPath)?.documentElement;
    if (!zip.file(doc.mainPath) || (root && (root.namespaceURI !== NS.w || root.localName !== "document"))) {
      throw new Error("Not a valid .docx file: word/document.xml is missing.");
    }
    if (!root) throw new Error("Failed to parse word/document.xml (file may be corrupt).");
    if (!doc.document.getElementsByTagNameNS(NS.w, "body")[0]) throw new Error("word/document.xml has no <w:body>.");
    return doc;
  }

  get document(): Document {
    return this.parts.get(this.mainPath)!;
  }

  get body(): Element {
    return this.document.getElementsByTagNameNS(NS.w, "body")[0];
  }

  stylesXml(): Document | null {
    const path = this.relOfType(this.mainPath, "styles");
    return path ? this.xml(path) : null;
  }

  /** The theme part (fonts that styles refer to by theme name). */
  themeXml(): Document | null {
    const path = this.relOfType(this.mainPath, "theme");
    return path ? this.xml(path) : null;
  }

  /** Add definitions for built-in paragraph styles the document uses but doesn't define (Word would fall back to Normal). */
  ensureParagraphStyles(ids: Iterable<string>): void {
    const styles = this.stylesXml();
    if (!styles) return;
    const have = new Set(Array.from(styles.getElementsByTagNameNS(NS.w, "style")).map((st) => st.getAttributeNS(NS.w, "styleId")));
    for (const id of new Set(ids)) {
      const def = BUILTIN_STYLES[id];
      if (!def || have.has(id)) continue;
      const [name, pPr, rPr] = def;
      const xml =
        `<w:style xmlns:w="${NS.w}" w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/><w:qFormat/><w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;
      styles.documentElement.appendChild(styles.importNode(parseXml(xml)!.documentElement, true));
      have.add(id);
    }
  }

  numberingXml(): Document | null {
    const path = this.relOfType(this.mainPath, "numbering");
    return path ? this.xml(path) : null;
  }

  /** Every image relationship of the main part as a data URL, keyed by rId. */
  /** rId -> data URL for the images a part (the body by default, or a header / footer) refers to. */
  async imageUrls(part = this.mainPath): Promise<Map<string, string>> {
    const out = new Map<string, string>();
    for (const rel of this.rels(part)) {
      const mime = IMAGE_MIME[rel.target.split(".").pop()?.toLowerCase() ?? ""];
      if (rel.type !== "image" || rel.external || !mime || !this.zip.file(rel.target)) continue;
      out.set(rel.id, `data:${mime};base64,${bytesToBase64(await this.media(rel.target))}`);
    }
    return out;
  }

  hyperlinkHref(rId: string, part = this.mainPath): string | null {
    const rel = this.rels(part).find((r) => r.id === rId && r.type === "hyperlink");
    return rel ? rel.target : null;
  }

  // --- Headers and footers ---------------------------------------------------

  /** The body's final w:sectPr (page setup), created if asked. */
  private finalSectPr(create: boolean): Element | null {
    const last = this.body.lastElementChild;
    if (last?.localName === "sectPr" && last.namespaceURI === NS.w) return last;
    if (!create) return null;
    return this.body.appendChild(this.document.createElementNS(NS.w, "w:sectPr"));
  }

  /** Part path of a header or footer, if the section has one of that type. */
  headerFooterPath(kind: "header" | "footer", type: HFType = "default"): string | null {
    const ref = Array.from(this.finalSectPr(false)?.children ?? []).find(
      (c) => c.localName === `${kind}Reference` && (wVal(c, "type") ?? "default") === type
    );
    return (ref && this.relTarget(this.mainPath, ref.getAttributeNS(NS.r, "id") ?? "")) || null;
  }

  /** Root (w:hdr / w:ftr) of a header or footer, if the document has one. */
  headerFooter(kind: "header" | "footer", type: HFType = "default"): Element | null {
    const path = this.headerFooterPath(kind, type);
    return (path && this.xml(path)?.documentElement) || null;
  }

  /** The header or footer of `type`, created (empty) and referenced from the final section if missing. */
  ensureHeaderFooter(kind: "header" | "footer", type: HFType = "default"): Element {
    const existing = this.headerFooter(kind, type);
    if (existing) return existing;
    const path = this.nextPartPath(`word/${kind}`, ".xml");
    const tag = kind === "header" ? "hdr" : "ftr";
    this.parts.set(path, parseXml(`<w:${tag} xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:p/></w:${tag}>`)!);
    this.addOverride(path, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    const rId = this.addRel(this.mainPath, kind, path);
    const sectPr = this.finalSectPr(true)!;
    const ref = this.document.createElementNS(NS.w, `w:${kind}Reference`);
    ref.setAttributeNS(NS.w, "w:type", type);
    ref.setAttributeNS(NS.r, "r:id", rId);
    // Header / footer references come first in w:sectPr.
    const before = Array.from(sectPr.children).find((c) => !c.localName.endsWith("Reference")) ?? null;
    sectPr.insertBefore(ref, before);
    return this.headerFooter(kind, type)!;
  }

  /** "Different first page" (w:titlePg on the final section). */
  differentFirstPage(): boolean {
    const el = Array.from(this.finalSectPr(false)?.children ?? []).find((c) => c.localName === "titlePg");
    return !!el && !["0", "false", "off"].includes(wVal(el) ?? "");
  }

  setDifferentFirstPage(on: boolean): void {
    const sectPr = this.finalSectPr(true)!;
    Array.from(sectPr.children).filter((c) => c.localName === "titlePg").forEach((c) => c.remove());
    if (!on) return;
    const at = SECT_ORDER.indexOf("titlePg");
    const before = Array.from(sectPr.children).find((c) => SECT_ORDER.indexOf(c.localName) > at) ?? null;
    sectPr.insertBefore(this.document.createElementNS(NS.w, "w:titlePg"), before);
  }

  /** "Different odd & even" pages (w:evenAndOddHeaders in settings.xml). */
  differentOddEven(): boolean {
    const el = this.partOf("settings")?.getElementsByTagNameNS(NS.w, "evenAndOddHeaders")[0];
    return !!el && !["0", "false", "off"].includes(wVal(el) ?? "");
  }

  setDifferentOddEven(on: boolean): void {
    const settings = on ? this.ensurePart("settings", "settings", `<w:settings xmlns:w="${NS.w}"/>`) : this.partOf("settings");
    Array.from(settings?.getElementsByTagNameNS(NS.w, "evenAndOddHeaders") ?? []).forEach((e) => e.remove());
    if (on) this.ensureSetting("evenAndOddHeaders", `<w:evenAndOddHeaders/>`);
  }

  private partOf(type: string): Document | null {
    const path = this.relOfType(this.mainPath, type);
    return path ? this.xml(path) : null;
  }

  /** A part related from the main document, created from `rootXml` if missing. */
  private ensurePart(type: string, file: string, rootXml: string): Document {
    const existing = this.partOf(type);
    if (existing) return existing;
    const path = `word/${file}.xml`;
    this.parts.set(path, parseXml(rootXml)!);
    this.addOverride(path, `application/vnd.openxmlformats-officedocument.wordprocessingml.${file}+xml`);
    this.addRel(this.mainPath, type, path);
    return this.parts.get(path)!;
  }

  private notes(part: Document | null, tag: string): Element[] {
    return Array.from(part?.getElementsByTagNameNS(NS.w, tag) ?? []).filter((n) => !["separator", "continuationSeparator", "continuationNotice"].includes(wVal(n, "type") ?? ""));
  }

  comments(): { id: string; author: string; initials: string; date: string; text: string }[] {
    return this.notes(this.partOf("comments"), "comment").map((c) => ({
      id: wVal(c, "id") ?? "",
      author: wVal(c, "author") ?? "",
      initials: wVal(c, "initials") ?? "",
      date: wVal(c, "date") ?? "",
      text: noteText(c),
    }));
  }

  /** New comment; returns its id (anchor it in the text with the editor's addCommentAnchors). */
  addComment(text: string, author: string): string {
    const part = this.ensurePart("comments", "comments", `<w:comments xmlns:w="${NS.w}"/>`);
    const id = String(nextId(part, "comment"));
    const initials = author.split(/\s+/).filter(Boolean).map((w) => w[0].toUpperCase()).join("").slice(0, 3);
    const date = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    const el = part.createElementNS(NS.w, "w:comment");
    for (const [k, v] of [["id", id], ["author", author], ["date", date], ["initials", initials]]) el.setAttributeNS(NS.w, `w:${k}`, v);
    part.documentElement.appendChild(el);
    setNoteText(el, text);
    return id;
  }

  setCommentText(id: string, text: string): void {
    const el = this.notes(this.partOf("comments"), "comment").find((c) => wVal(c, "id") === id);
    if (el) setNoteText(el, text);
  }

  deleteComment(id: string): void {
    this.notes(this.partOf("comments"), "comment").find((c) => wVal(c, "id") === id)?.remove();
  }

  // --- Footnotes -------------------------------------------------------------------

  footnotes(): { id: string; text: string }[] {
    return this.notes(this.partOf("footnotes"), "footnote").map((f) => ({ id: wVal(f, "id") ?? "", text: noteText(f) }));
  }

  /** New footnote; returns its id (place the reference with the editor's insertFootnote). */
  addFootnote(text: string): string {
    const fresh = !this.partOf("footnotes");
    const sep = (type: string, id: number, mark: string) =>
      `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${mark}/></w:r></w:p></w:footnote>`;
    const part = this.ensurePart(
      "footnotes",
      "footnotes",
      `<w:footnotes xmlns:w="${NS.w}">${sep("separator", -1, "separator")}${sep("continuationSeparator", 0, "continuationSeparator")}</w:footnotes>`
    );
    if (fresh) this.ensureSetting("footnotePr", `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`);
    const id = String(Math.max(1, nextId(part, "footnote")));
    const el = parseXml(`<w:footnote xmlns:w="${NS.w}" w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r></w:p></w:footnote>`)!.documentElement;
    const node = part.documentElement.appendChild(part.importNode(el, true));
    setNoteText(node, text);
    return id;
  }

  setFootnoteText(id: string, text: string): void {
    const el = this.notes(this.partOf("footnotes"), "footnote").find((f) => wVal(f, "id") === id);
    if (el) setNoteText(el, text);
  }

  deleteFootnote(id: string): void {
    this.notes(this.partOf("footnotes"), "footnote").find((f) => wVal(f, "id") === id)?.remove();
  }

  /** Styles a table of contents uses, and a request for Word to refresh fields (page numbers) on open. */
  prepareToc(): void {
    this.ensureParagraphStyles(["TOCHeading", "TOC1", "TOC2", "TOC3"]);
    this.ensureSetting("updateFields", `<w:updateFields w:val="true"/>`);
  }

  /** Add a document setting (settings.xml) at its schema position, if the document has settings and lacks it. */
  ensureSetting(local: string, xml: string): void {
    const settings = this.partOf("settings");
    if (!settings || settings.getElementsByTagNameNS(NS.w, local).length) return;
    const root = settings.documentElement;
    const at = SETTINGS_ORDER.indexOf(local);
    const before = Array.from(root.children).find((c) => SETTINGS_ORDER.indexOf(c.localName) > at) ?? null;
    root.insertBefore(settings.importNode(parseXml(`<r xmlns:w="${NS.w}">${xml}</r>`)!.documentElement.firstElementChild!, true), before);
  }

  hyperlinkRel(href: string): string {
    return this.rels(this.mainPath).find((r) => r.type === "hyperlink" && r.external && r.target === href)?.id ?? this.addRel(this.mainPath, "hyperlink", href, true);
  }

  /** Store an image and relate it to the document. Returns the rId. */
  addImage(bytes: Uint8Array, ext: string, part = this.mainPath): string {
    return this.addRel(part, "image", this.storeImage("word", bytes, ext));
  }

  /** Width of the last section's text area (page width less the side margins), in twips. */
  textWidth(): number {
    const sectPr = this.finalSectPr(false);
    const get = (local: string, attr: string, fallback: number) => {
      const el = Array.from(sectPr?.children ?? []).find((c) => c.localName === local);
      return Number(el ? wVal(el, attr) : null) || fallback;
    };
    return get("pgSz", "w", 12240) - get("pgMar", "left", 1440) - get("pgMar", "right", 1440);
  }

  /** wp:docPr ids must be unique across the document. */
  nextDocPrId(): number {
    const ids = Array.from(this.document.getElementsByTagNameNS(NS.wp, "docPr")).map((e) => Number(e.getAttribute("id")) || 0);
    const next = Math.max(0, this.lastDocPrId, ...ids) + 1;
    this.lastDocPrId = next;
    return next;
  }
  private lastDocPrId = 0;

  /** numFmt of a list level ("bullet", "decimal", ...), or null if undefined. */
  listFormat(numId: string, ilvl: number): string | null {
    const lvl = this.listLevel(numId, ilvl);
    return lvl ? wVal(childNS(lvl, NS.w, "numFmt")) : null;
  }

  listLevel(numId: string, ilvl: number): Element | null {
    return numberingLevel(this.numberingXml(), numId, ilvl);
  }

  /** A numId for a multi-level bullet/ordered list, reusing one we created before or creating it. */
  ensureList(kind: ListKind): string {
    let numbering = this.numberingXml();
    if (!numbering) {
      const path = "word/numbering.xml";
      numbering = parseXml(`<w:numbering xmlns:w="${NS.w}"/>`)!;
      this.parts.set(path, numbering);
      this.addOverride(path, NUMBERING_CT);
      this.addRel(this.mainPath, "numbering", path);
    }
    const wanted = kind === "bullet" ? "bullet" : "decimal";
    const nums = Array.from(numbering.getElementsByTagNameNS(NS.w, "num"));
    // Only reuse multi-level definitions: single-level template lists can't be nested.
    const reuse = nums.find((n) => {
      const id = wVal(n, "numId")!;
      return this.listFormat(id, 0) === wanted && this.listLevel(id, 8) !== null;
    });
    if (reuse) return wVal(reuse, "numId")!;

    const root = numbering.documentElement;
    const maxOf = (tag: string, attr: string) =>
      Array.from(numbering.getElementsByTagNameNS(NS.w, tag)).reduce((m, e) => Math.max(m, Number(wVal(e, attr)) || 0), 0);
    const absId = maxOf("abstractNum", "abstractNumId") + 1;
    const numId = String(maxOf("num", "numId") + 1);

    // Schema order: all abstractNum elements precede all num elements.
    const abs = numbering.importNode(parseXml(abstractNumXml(absId, kind))!.documentElement, true);
    root.insertBefore(abs, nums[0] ?? childNS(root, NS.w, "numIdMacAtCleanup"));
    const num = numbering.createElementNS(NS.w, "w:num");
    num.setAttributeNS(NS.w, "w:numId", numId);
    const ref = numbering.createElementNS(NS.w, "w:abstractNumId");
    ref.setAttributeNS(NS.w, "w:val", String(absId));
    num.appendChild(ref);
    root.insertBefore(num, childNS(root, NS.w, "numIdMacAtCleanup"));
    return numId;
  }
}
