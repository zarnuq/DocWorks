import JSZip from "jszip";
import { NS, childNS } from "../../core/utils/xml";
import { OoxmlPackage, parseXml } from "../ooxml/ooxml-package";

export type ListKind = "bullet" | "ordered";

const DEFAULT_MAIN = "word/document.xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const IMAGE_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
};
const BULLETS = ["•", "◦", "▪"];
const ORDERED_FORMATS = ["decimal", "lowerLetter", "lowerRoman"];

const wVal = (el: Element | null | undefined, attr = "val") => el?.getAttributeNS(NS.w, attr) ?? null;

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
};

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
  async imageUrls(): Promise<Map<string, string>> {
    const out = new Map<string, string>();
    for (const rel of this.rels(this.mainPath)) {
      const mime = IMAGE_MIME[rel.target.split(".").pop()?.toLowerCase() ?? ""];
      if (rel.type !== "image" || rel.external || !mime || !this.zip.file(rel.target)) continue;
      out.set(rel.id, `data:${mime};base64,${bytesToBase64(await this.media(rel.target))}`);
    }
    return out;
  }

  hyperlinkHref(rId: string): string | null {
    const rel = this.rels(this.mainPath).find((r) => r.id === rId && r.type === "hyperlink");
    return rel ? rel.target : null;
  }

  /** Store an image and relate it to the document. Returns the rId. */
  addImage(bytes: Uint8Array, ext: string): string {
    return this.addRel(this.mainPath, "image", this.storeImage("word", bytes, ext));
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
    const numbering = this.numberingXml();
    if (!numbering) return null;
    const num = Array.from(numbering.getElementsByTagNameNS(NS.w, "num")).find((n) => wVal(n, "numId") === numId);
    const absId = wVal(num && childNS(num, NS.w, "abstractNumId"));
    const abs = Array.from(numbering.getElementsByTagNameNS(NS.w, "abstractNum")).find((a) => wVal(a, "abstractNumId") === absId);
    return Array.from(abs?.getElementsByTagNameNS(NS.w, "lvl") ?? []).find((l) => wVal(l, "ilvl") === String(ilvl)) ?? null;
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
