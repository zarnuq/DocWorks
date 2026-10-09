import JSZip from "jszip";
import { NS, escapeXml } from "../../core/utils/xml";
import { nextShapeId, setRect, writePlainParagraphs } from "../../domain/pptx/slide-edit";
import { itemRect, setItemRect, xfrmOf } from "../../domain/pptx/arrange";
import { parseSlide, plainText } from "../../domain/pptx/slide-model";
import { Fill, writeFill } from "../../domain/pptx/shape-style";
import { TABLE_STYLES } from "../../domain/pptx/table";
import { OoxmlPackage, PackageSnapshot, parseXml, relsPathOf } from "../ooxml/ooxml-package";

export interface SlideRef {
  path: string;
  id: string;
  rId: string;
}

export type PptxSnapshot = PackageSnapshot;

export interface LayoutRef {
  path: string;
  name: string;
  /** ST_SlideLayoutType, e.g. "title", "obj", "blank"; "cust" when unset. */
  type: string;
}

export interface Section {
  name: string;
  id: string;
  /** Index of the section's first slide (== next section's start when empty). */
  start: number;
}

interface SectionSpan {
  name: string;
  id: string;
  size: number;
}

const PRESENTATION = "ppt/presentation.xml";
const SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const NOTES_CT = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const SKIPPED_PLACEHOLDERS = new Set(["dt", "ftr", "sldNum"]);
const SECTIONS_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

const phOf = (el: Element) => el.getElementsByTagNameNS(NS.p, "nvPr")[0]?.getElementsByTagNameNS(NS.p, "ph")[0] ?? null;
const phFamily = (ph: Element) => (/^(title|ctrTitle)$/.test(ph.getAttribute("type") ?? "") ? "title" : "body");

/** moveSlide target for dragging slide `from` onto slide `target`'s top (or bottom) half. */
export function dropIndex(from: number, target: number, below: boolean): number {
  const to = target + (below ? 1 : 0);
  return from < to ? to - 1 : to;
}

const newGuid = () => `{${crypto.randomUUID().toUpperCase()}}`;

/** Index of the section holding slide `pos` (the first section for pos < 0, the last past the end). */
function sectionAt(spans: SectionSpan[], pos: number): number {
  let end = 0;
  for (let k = 0; k < spans.length; k++) if (pos < (end += spans[k].size)) return k;
  return spans.length - 1;
}

/** A slide placeholder for layout placeholder `ph`, with no own position or text. */
function placeholderXml(ph: Element, id: number): string {
  const attrs = Array.from(ph.attributes).map((a) => ` ${a.name}="${escapeXml(a.value)}"`).join("");
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

/** A .pptx package: slide list and slide operations on top of the live-DOM OOXML package. */
export class PptxDocument extends OoxmlPackage {
  slideSize = { cx: 9144000, cy: 6858000 };

  static async load(data: ArrayBuffer | Uint8Array): Promise<PptxDocument> {
    const zip = await JSZip.loadAsync(data);
    if (!zip.file(PRESENTATION)) {
      throw new Error("Not a valid .pptx file: ppt/presentation.xml is missing.");
    }
    const doc = new PptxDocument(zip);
    await doc.readParts();
    if (!doc.parts.has(PRESENTATION)) throw new Error("Failed to parse ppt/presentation.xml (file may be corrupt).");
    const sz = doc.parts.get(PRESENTATION)!.getElementsByTagNameNS(NS.p, "sldSz")[0];
    if (sz) doc.slideSize = { cx: Number(sz.getAttribute("cx")), cy: Number(sz.getAttribute("cy")) };
    return doc;
  }

  slides(): SlideRef[] {
    const pres = this.parts.get(PRESENTATION)!;
    return Array.from(pres.getElementsByTagNameNS(NS.p, "sldId")).flatMap((el) => {
      const rId = el.getAttributeNS(NS.r, "id") ?? "";
      const path = this.relTarget(PRESENTATION, rId);
      return path && this.parts.has(path) ? [{ path, id: el.getAttribute("id") ?? "", rId }] : [];
    });
  }

  layoutOf = (slidePath: string) => this.relOfType(slidePath, "slideLayout");
  masterOf = (layoutPath: string) => this.relOfType(layoutPath, "slideMaster");
  themeOf = (masterPath: string) => this.relOfType(masterPath, "theme");
  notesOf = (slidePath: string) => this.relOfType(slidePath, "notesSlide");

  /** Insert a new slide after `after` (-1 = first). Returns the new slide's index. */
  addSlide(after: number, layoutPath?: string): number {
    const slides = this.slides();
    const neighbour = slides[Math.max(0, Math.min(after, slides.length - 1))];
    const layout = layoutPath ?? (neighbour && this.nextLayout(this.layoutOf(neighbour.path))) ?? this.firstLayout();
    if (!layout) throw new Error("Presentation has no slide layouts.");

    const path = this.nextPartPath("ppt/slides/slide", ".xml");
    let spId = 2;
    const placeholders = this.layoutPlaceholders(layout)
      .map((ph) => placeholderXml(ph, spId++))
      .join("");

    this.parts.set(
      path,
      parseXml(
        `<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
          `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
          `${placeholders}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
      )!
    );
    this.addRel(path, "slideLayout", layout);
    return this.registerSlide(path, after);
  }

  /** Copy slide `index` (minus its speaker notes) right after it. Returns the copy's index. */
  duplicateSlide(index: number): number {
    const src = this.slides()[index];
    if (!src) throw new Error(`No slide at index ${index}`);
    const path = this.nextPartPath("ppt/slides/slide", ".xml");
    this.parts.set(path, this.parts.get(src.path)!.cloneNode(true) as Document);

    const rels = this.emptyRels();
    const srcRels = this.parts.get(relsPathOf(src.path));
    for (const r of Array.from(srcRels?.getElementsByTagNameNS(NS.pkgRels, "Relationship") ?? [])) {
      if (!r.getAttribute("Type")?.endsWith("/notesSlide")) rels.documentElement.appendChild(rels.importNode(r, true));
    }
    this.parts.set(relsPathOf(path), rels);
    return this.registerSlide(path, index);
  }

  deleteSlide(index: number): void {
    const slides = this.slides();
    const slide = slides[index];
    if (!slide) throw new Error(`No slide at index ${index}`);
    if (slides.length <= 1) throw new Error("Cannot delete the last slide.");
    const done = this.adjustSections((spans) => spans[sectionAt(spans, index)].size--);

    const pres = this.parts.get(PRESENTATION)!;
    const sldId = Array.from(pres.getElementsByTagNameNS(NS.p, "sldId")).find((e) => e.getAttributeNS(NS.r, "id") === slide.rId);
    sldId?.remove();
    this.removeRel(PRESENTATION, slide.rId);

    const notes = this.notesOf(slide.path);
    if (notes) this.removePart(notes);
    this.removePart(slide.path);
    done();
  }

  moveSlide(from: number, to: number): void {
    const ids = Array.from(this.parts.get(PRESENTATION)!.getElementsByTagNameNS(NS.p, "sldId"));
    if (from === to || from < 0 || to < 0 || from >= ids.length || to >= ids.length) return;
    // The moved slide joins the section of the slide it lands after (the first section at the top).
    const done = this.adjustSections((spans) => {
      spans[sectionAt(spans, from)].size--;
      spans[sectionAt(spans, to - 1)].size++;
    });
    const el = ids[from];
    const list = el.parentNode!;
    el.remove();
    const rest = ids.filter((e) => e !== el);
    list.insertBefore(el, rest[to] ?? null);
    done();
  }

  // --- Embedded fonts --------------------------------------------------------

  /** Fonts the deck carries (Google Slides and PowerPoint's "Embed fonts"), still in their stored form (EOT). */
  async embeddedFonts(): Promise<{ family: string; bold: boolean; italic: boolean; data: Uint8Array }[]> {
    const out: { family: string; bold: boolean; italic: boolean; data: Uint8Array }[] = [];
    const list = this.parts.get(PRESENTATION)?.getElementsByTagNameNS(NS.p, "embeddedFont") ?? [];
    for (const font of Array.from(list)) {
      const family = font.getElementsByTagNameNS(NS.p, "font")[0]?.getAttribute("typeface");
      if (!family) continue;
      for (const [tag, bold, italic] of [["regular", false, false], ["bold", true, false], ["italic", false, true], ["boldItalic", true, true]] as const) {
        const rId = font.getElementsByTagNameNS(NS.p, tag)[0]?.getAttributeNS(NS.r, "id");
        const path = rId && this.relTarget(PRESENTATION, rId);
        if (path && this.zip.file(path)) out.push({ family, bold, italic, data: await this.media(path) });
      }
    }
    return out;
  }

  // --- Speaker notes -------------------------------------------------------

  /** The notes page's body placeholder text body, if the slide has notes. */
  private notesBody(index: number): Element | null {
    const notes = this.notesOf(this.slides()[index]?.path ?? "");
    const sp = Array.from(this.xml(notes ?? "")?.getElementsByTagNameNS(NS.p, "sp") ?? []).find(
      (s) => s.getElementsByTagNameNS(NS.p, "ph")[0]?.getAttribute("type") === "body"
    );
    return Array.from(sp?.children ?? []).find((c) => c.localName === "txBody") ?? null;
  }

  notesText(index: number): string {
    return plainText(this.notesBody(index));
  }

  /** Set a slide's notes as plain text, creating its notes page (and the deck's notes master) if needed. */
  setNotes(index: number, text: string): boolean {
    let body = this.notesBody(index);
    if (!body) {
      if (!text) return true;
      body = this.createNotes(this.slides()[index].path);
    }
    writePlainParagraphs(body, text);
    return true;
  }

  /** The deck's notes master, created with a copy of the deck theme if it has none (new decks, decks never given notes). */
  private ensureNotesMaster(): string {
    const existing = this.relOfType(PRESENTATION, "notesMaster");
    if (existing) return existing;
    const pres = this.parts.get(PRESENTATION)!;
    const sz = pres.getElementsByTagNameNS(NS.p, "notesSz")[0];
    const cx = Number(sz?.getAttribute("cx")) || 6858000;
    const cy = Number(sz?.getAttribute("cy")) || 9144000;
    const box = (x: number, y: number, w: number, h: number) =>
      `<a:xfrm><a:off x="${Math.round(cx * x)}" y="${Math.round(cy * y)}"/><a:ext cx="${Math.round(cx * w)}" cy="${Math.round(cy * h)}"/></a:xfrm>`;
    const clrMap = ["bg1=lt1", "tx1=dk1", "bg2=lt2", "tx2=dk2", ...[1, 2, 3, 4, 5, 6].map((i) => `accent${i}=accent${i}`), "hlink=hlink", "folHlink=folHlink"]
      .map((kv) => kv.replace(/(\w+)=(\w+)/, '$1="$2"'))
      .join(" ");
    const path = this.nextPartPath("ppt/notesMasters/notesMaster", ".xml");
    this.parts.set(
      path,
      parseXml(
        `<p:notesMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>` +
          `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>` +
          `<p:spPr>${box(0.1, 0.125, 0.8, 0.3375)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr>` +
          `<p:spPr>${box(0.1, 0.48125, 0.8, 0.39375)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
          `<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>` +
          `</p:spTree></p:cSld><p:clrMap ${clrMap}/>` +
          `<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200">` +
          `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`
      )!
    );
    this.addOverride(path, "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml");
    // PowerPoint wants each master to have its own theme part.
    const deckTheme = this.xml(this.relOfType(PRESENTATION, "theme") ?? "");
    if (deckTheme) {
      const theme = this.nextPartPath("ppt/theme/theme", ".xml");
      this.parts.set(theme, deckTheme.cloneNode(true) as Document);
      this.addOverride(theme, "application/vnd.openxmlformats-officedocument.theme+xml");
      this.addRel(path, "theme", theme);
    }
    const rId = this.addRel(PRESENTATION, "notesMaster", path);
    const list = pres.createElementNS(NS.p, "p:notesMasterIdLst");
    list.appendChild(pres.createElementNS(NS.p, "p:notesMasterId")).setAttributeNS(NS.r, "r:id", rId);
    const masters = pres.getElementsByTagNameNS(NS.p, "sldMasterIdLst")[0];
    pres.documentElement.insertBefore(list, masters ? masters.nextSibling : pres.documentElement.firstChild);
    return path;
  }

  private createNotes(slidePath: string): Element {
    const path = this.nextPartPath("ppt/notesSlides/notesSlide", ".xml");
    const ph = (id: number, name: string, ph: string, extra = "") =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
      `<p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr><p:spPr/>${extra}</p:sp>`;
    this.parts.set(
      path,
      parseXml(
        `<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
          `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
          ph(2, "Slide Image Placeholder 1", 'type="sldImg"') +
          ph(3, "Notes Placeholder 2", 'type="body" idx="1"', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>`) +
          `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
      )!
    );
    this.addOverride(path, NOTES_CT);
    this.addRel(path, "notesMaster", this.ensureNotesMaster());
    this.addRel(path, "slide", slidePath);
    this.addRel(slidePath, "notesSlide", path);
    return Array.from(this.parts.get(path)!.getElementsByTagNameNS(NS.p, "txBody"))[0];
  }

  // --- Hidden slides -------------------------------------------------------

  isHidden(index: number): boolean {
    return this.xml(this.slides()[index]?.path ?? "")?.documentElement.getAttribute("show") === "0";
  }

  setHidden(index: number, hidden: boolean): void {
    const root = this.xml(this.slides()[index].path)!.documentElement;
    if (hidden) root.setAttribute("show", "0");
    else root.removeAttribute("show");
  }

  /** Add a relationship from a slide (used when pasting shapes). */
  addSlideRel(slidePath: string, type: string, target: string, external = false): string {
    return this.addRel(slidePath, type, target, external);
  }

  /** Relationship id of a hyperlink from `part` to `url`, adding one if needed. */
  linkRel(part: string, url: string): string {
    return this.rels(part).find((r) => r.type === "hyperlink" && r.external && r.target === url)?.id ?? this.addRel(part, "hyperlink", url, true);
  }

  /** Store image bytes in ppt/media and relate them to the slide. Returns the new rId. */
  addImage(slidePath: string, bytes: Uint8Array, ext: string): string {
    return this.addRel(slidePath, "image", this.storeImage("ppt", bytes, ext));
  }

  /** Like PowerPoint, a slide added after a title slide gets "Title and Content" when the master has it. */
  private nextLayout(layout: string | null): string | null {
    if (!layout || this.parts.get(layout)?.documentElement.getAttribute("type") !== "title") return layout;
    const master = this.masterOf(layout);
    const content = master && this.rels(master).find((r) => r.type === "slideLayout" && this.parts.get(r.target)?.documentElement.getAttribute("type") === "obj");
    return content ? content.target : layout;
  }

  private firstLayout(): string | null {
    const masterRel = this.rels(PRESENTATION).find((r) => r.type === "slideMaster");
    return masterRel ? this.relOfType(masterRel.target, "slideLayout") : null;
  }

  // --- Layouts ---------------------------------------------------------------

  /** Layouts of `master` (default: the first master), in the master's order. */
  layouts(master = this.rels(PRESENTATION).find((r) => r.type === "slideMaster")?.target ?? ""): LayoutRef[] {
    const ids = Array.from(this.xml(master)?.getElementsByTagNameNS(NS.p, "sldLayoutId") ?? []);
    return ids.flatMap((id) => {
      const path = this.relTarget(master, id.getAttributeNS(NS.r, "id") ?? "");
      const root = path ? this.xml(path)?.documentElement : null;
      if (!path || !root) return [];
      const name = root.getElementsByTagNameNS(NS.p, "cSld")[0]?.getAttribute("name") ?? "";
      return [{ path, name, type: root.getAttribute("type") ?? "cust" }];
    });
  }

  /**
   * Switch slide `index` to `layout`. Placeholders are matched to the new layout's
   * (titles to titles, the rest by idx, then in order); unmatched empty ones are
   * dropped, unmatched filled ones stay where they were; missing ones are added.
   */
  setLayout(index: number, layout: string): void {
    const slide = this.slides()[index];
    if (!slide || !this.parts.has(layout)) return;
    const xml = this.parts.get(slide.path)!;
    const tree = xml.getElementsByTagNameNS(NS.p, "spTree")[0];
    const rects = new Map(parseSlide(this, slide.path).shapes.map((s) => [s.el, s.rect]));
    const free = this.layoutPlaceholders(layout);
    const take = (ph: Element) => {
      const same = free.filter((l) => phFamily(l) === phFamily(ph));
      const m = (phFamily(ph) === "body" && same.find((l) => (l.getAttribute("idx") ?? "0") === (ph.getAttribute("idx") ?? "0"))) || same[0];
      if (m) free.splice(free.indexOf(m), 1);
      return m ?? null;
    };

    for (const item of Array.from(tree.children)) {
      const ph = phOf(item);
      if (!ph || SKIPPED_PLACEHOLDERS.has(ph.getAttribute("type") ?? "")) continue;
      const match = take(ph);
      if (match) {
        for (const a of Array.from(ph.attributes)) ph.removeAttribute(a.name);
        for (const a of Array.from(match.attributes)) ph.setAttribute(a.name, a.value);
        // Like PowerPoint, a re-laid-out placeholder takes the new layout's position.
        if (item.localName !== "graphicFrame") xfrmOf(item)?.remove();
      } else if (item.localName === "sp" && !item.textContent?.trim()) {
        item.remove();
      } else if (!xfrmOf(item) && rects.has(item)) {
        setRect(item, rects.get(item)!);
      }
    }
    const id = nextShapeId(xml);
    const added = free.map((ph, i) => placeholderXml(ph, id + i)).join("");
    if (added) {
      const frag = parseXml(`<r xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">${added}</r>`)!;
      for (const sp of Array.from(frag.documentElement.children)) tree.appendChild(xml.importNode(sp, true));
    }
    const rel = this.rels(slide.path).find((r) => r.type === "slideLayout");
    if (rel) this.removeRel(slide.path, rel.id);
    this.addRel(slide.path, "slideLayout", layout);
  }

  private layoutPlaceholders(layout: string): Element[] {
    return Array.from(this.parts.get(layout)?.getElementsByTagNameNS(NS.p, "spTree")[0]?.children ?? [])
      .map(phOf)
      .filter((ph): ph is Element => !!ph && !SKIPPED_PLACEHOLDERS.has(ph.getAttribute("type") ?? ""));
  }

  // --- Table styles --------------------------------------------------------------

  /** Make sure ppt/tableStyles.xml defines built-in style `id` (PowerPoint and LibreOffice read it from there). */
  ensureTableStyle(id: string): void {
    const def = TABLE_STYLES.find((s) => s.id === id);
    if (!def) return;
    let path = this.relOfType(PRESENTATION, "tableStyles");
    if (!path || !this.parts.has(path)) {
      path = "ppt/tableStyles.xml";
      this.parts.set(path, parseXml(`<a:tblStyleLst xmlns:a="${NS.a}" def="${TABLE_STYLES[1].id}"/>`)!);
      this.addRel(PRESENTATION, "tableStyles", path);
      this.addOverride(path, "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml");
    }
    const xml = this.parts.get(path)!;
    if (Array.from(xml.getElementsByTagNameNS(NS.a, "tblStyle")).some((s) => s.getAttribute("styleId") === id)) return;
    const el = parseXml(`<r xmlns:a="${NS.a}">${def.xml}</r>`)!.documentElement.firstElementChild!;
    xml.documentElement.appendChild(xml.importNode(el, true));
  }

  // --- Background and size -----------------------------------------------------

  /** Give slide `index` its own background fill; null goes back to the layout / master background. */
  setBackground(index: number, fill: Fill | null): void {
    const xml = this.xml(this.slides()[index]?.path ?? "");
    const cSld = xml?.getElementsByTagNameNS(NS.p, "cSld")[0];
    if (!xml || !cSld) return;
    Array.from(cSld.children)
      .filter((c) => c.localName === "bg")
      .forEach((c) => c.remove());
    if (!fill) return;
    const bg = xml.createElementNS(NS.p, "p:bg");
    const bgPr = bg.appendChild(xml.createElementNS(NS.p, "p:bgPr"));
    bgPr.appendChild(xml.createElementNS(NS.a, "a:effectLst"));
    writeFill(bgPr, fill);
    cSld.insertBefore(bg, cSld.firstChild);
  }

  /**
   * Change the slide size. Masters and layouts stretch to it, so placeholders
   * fill the new shape; slide content is scaled uniformly to fit and centered
   * (PowerPoint's "Ensure fit"). Pictures and frames keep their aspect ratio.
   */
  // ponytail: font sizes are not scaled; shrink text by hand after big downsizes.
  setSlideSize(cx: number, cy: number): void {
    const { cx: ox, cy: oy } = this.slideSize;
    const sx = cx / ox;
    const sy = cy / oy;
    const s = Math.min(sx, sy);
    for (const [path, xml] of this.parts) {
      const kind = /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.exec(path)?.[1];
      if (!kind) continue;
      for (const item of Array.from(xml.getElementsByTagNameNS(NS.p, "spTree")[0]?.children ?? [])) {
        if (!xfrmOf(item)) continue;
        const r = itemRect(item);
        if (kind !== "slides" && item.localName !== "pic" && item.localName !== "graphicFrame") {
          setItemRect(item, { x: r.x * sx, y: r.y * sy, w: r.w * sx, h: r.h * sy });
        } else {
          // Uniform size around the center, which moves with the stretch (slides: with the fit).
          const [mx, my, ox2, oy2] = kind === "slides" ? [s, s, (cx - ox * s) / 2, (cy - oy * s) / 2] : [sx, sy, 0, 0];
          const ccx = (r.x + r.w / 2) * mx + ox2;
          const ccy = (r.y + r.h / 2) * my + oy2;
          setItemRect(item, { x: ccx - (r.w * s) / 2, y: ccy - (r.h * s) / 2, w: r.w * s, h: r.h * s });
        }
      }
    }
    const sz = this.parts.get(PRESENTATION)!.getElementsByTagNameNS(NS.p, "sldSz")[0];
    sz.setAttribute("cx", String(Math.round(cx)));
    sz.setAttribute("cy", String(Math.round(cy)));
    sz.removeAttribute("type");
    this.slideSize = { cx: Math.round(cx), cy: Math.round(cy) };
  }

  // --- Sections ------------------------------------------------------------------

  sections(): Section[] {
    let start = 0;
    return this.sectionSpans().map(({ name, id, size }) => {
      const sec = { name, id, start };
      start += size;
      return sec;
    });
  }

  /** Start a new section at slide `start`; with no sections yet, earlier slides go to "Default Section". */
  addSection(start: number, name: string): void {
    const total = this.slides().length;
    let spans = this.sectionSpans();
    if (!spans.length) spans = [{ name: "Default Section", id: newGuid(), size: total }];
    const k = Math.min(sectionAt(spans, start), spans.length - 1);
    const first = spans.slice(0, k).reduce((n, sp) => n + sp.size, 0);
    const before = Math.max(0, Math.min(start - first, spans[k].size));
    spans.splice(k + 1, 0, { name, id: newGuid(), size: spans[k].size - before });
    spans[k].size = before;
    // A default section that ended up empty at the very top isn't wanted.
    if (k === 0 && before === 0 && !this.sectionSpans().length) spans.shift();
    this.writeSections(spans);
  }

  renameSection(index: number, name: string): void {
    const spans = this.sectionSpans();
    if (!spans[index]) return;
    spans[index].name = name;
    this.writeSections(spans);
  }

  /** Remove a section; its slides join the section before it (or after it, for the first). */
  removeSection(index: number): void {
    const spans = this.sectionSpans();
    if (!spans[index]) return;
    const [gone] = spans.splice(index, 1);
    const into = spans[Math.max(0, index - 1)];
    if (into) into.size += gone.size;
    this.writeSections(spans);
  }

  private sectionList(): Element | null {
    return this.parts.get(PRESENTATION)!.getElementsByTagNameNS(P14, "sectionLst")[0] ?? null;
  }

  private sectionSpans(): SectionSpan[] {
    return Array.from(this.sectionList()?.getElementsByTagNameNS(P14, "section") ?? []).map((sec) => ({
      name: sec.getAttribute("name") ?? "",
      id: sec.getAttribute("id") ?? newGuid(),
      size: sec.getElementsByTagNameNS(P14, "sldId").length,
    }));
  }

  /** Read the spans and let `fn` resize them; call the result after changing the slide list to write them back. */
  private adjustSections(fn: (spans: SectionSpan[]) => void): () => void {
    const spans = this.sectionSpans();
    if (!spans.length) return () => undefined;
    fn(spans);
    return () => this.writeSections(spans);
  }

  /** Rewrite p14:sectionLst so the spans cover the slides in order; no spans removes it. */
  private writeSections(spans: SectionSpan[]): void {
    const pres = this.parts.get(PRESENTATION)!;
    const ext = Array.from(pres.getElementsByTagNameNS(NS.p, "ext")).find((e) => e.getAttribute("uri") === SECTIONS_EXT);
    if (!spans.length) {
      const extLst = ext?.parentElement;
      ext?.remove();
      if (extLst && !extLst.children.length) extLst.remove();
      return;
    }
    const ids = this.slides().map((sl) => sl.id);
    let i = 0;
    const xml = spans
      .map(({ name, id, size }) => {
        const slides = ids.slice(i, (i += size)).map((sid) => `<p14:sldId id="${sid}"/>`).join("");
        return `<p14:section name="${escapeXml(name)}" id="${id}"><p14:sldIdLst>${slides}</p14:sldIdLst></p14:section>`;
      })
      .join("");
    const lst = parseXml(`<p14:sectionLst xmlns:p14="${P14}">${xml}</p14:sectionLst>`)!.documentElement;
    if (ext) {
      ext.replaceChildren(pres.importNode(lst, true));
      return;
    }
    let extLst = Array.from(pres.documentElement.children).find((c) => c.localName === "extLst");
    if (!extLst) extLst = pres.documentElement.appendChild(pres.createElementNS(NS.p, "p:extLst"));
    const e = pres.createElementNS(NS.p, "p:ext");
    e.setAttribute("uri", SECTIONS_EXT);
    e.appendChild(pres.importNode(lst, true));
    extLst.insertBefore(e, extLst.firstChild);
  }

  private registerSlide(path: string, after: number): number {
    this.addOverride(path, SLIDE_CT);

    const rId = this.addRel(PRESENTATION, "slide", path);
    const pres = this.parts.get(PRESENTATION)!;
    const ids = Array.from(pres.getElementsByTagNameNS(NS.p, "sldId"));
    const nextId = ids.reduce((m, e) => Math.max(m, Number(e.getAttribute("id"))), 255) + 1;

    let list = pres.getElementsByTagNameNS(NS.p, "sldIdLst")[0];
    if (!list) {
      // sldIdLst must follow sldMasterIdLst / notesMasterIdLst / handoutMasterIdLst.
      list = pres.createElementNS(NS.p, "p:sldIdLst");
      const anchor = Array.from(pres.documentElement.children).find((c) => c.localName === "sldSz") ?? null;
      pres.documentElement.insertBefore(list, anchor);
    }
    const el = pres.createElementNS(NS.p, "p:sldId");
    el.setAttribute("id", String(nextId));
    el.setAttributeNS(NS.r, "r:id", rId);
    const index = Math.max(-1, Math.min(after, ids.length - 1)) + 1;
    const done = this.adjustSections((spans) => spans[sectionAt(spans, index - 1)].size++);
    list.insertBefore(el, ids[index] ?? null);
    done();
    return index;
  }
}
