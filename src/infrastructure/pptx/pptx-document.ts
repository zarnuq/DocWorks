import JSZip from "jszip";
import { NS, escapeXml } from "../../core/utils/xml";
import { OoxmlPackage, PackageSnapshot, parseXml, relsPathOf } from "../ooxml/ooxml-package";

export interface SlideRef {
  path: string;
  id: string;
  rId: string;
}

export type PptxSnapshot = PackageSnapshot;

const PRESENTATION = "ppt/presentation.xml";
const SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SKIPPED_PLACEHOLDERS = new Set(["dt", "ftr", "sldNum"]);

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
    const layout = layoutPath ?? (neighbour && this.layoutOf(neighbour.path)) ?? this.firstLayout();
    if (!layout) throw new Error("Presentation has no slide layouts.");

    const path = this.nextPartPath("ppt/slides/slide", ".xml");
    let spId = 2;
    const placeholders = Array.from(this.parts.get(layout)?.getElementsByTagNameNS(NS.p, "sp") ?? [])
      .map((sp) => sp.getElementsByTagNameNS(NS.p, "ph")[0])
      .filter((ph) => ph && !SKIPPED_PLACEHOLDERS.has(ph.getAttribute("type") ?? ""))
      .map((ph) => {
        const attrs = Array.from(ph.attributes).map((a) => ` ${a.name}="${escapeXml(a.value)}"`).join("");
        const id = spId++;
        return (
          `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
          `<p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr><p:spPr/>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
        );
      })
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

    const pres = this.parts.get(PRESENTATION)!;
    const sldId = Array.from(pres.getElementsByTagNameNS(NS.p, "sldId")).find((e) => e.getAttributeNS(NS.r, "id") === slide.rId);
    sldId?.remove();
    this.removeRel(PRESENTATION, slide.rId);

    const notes = this.notesOf(slide.path);
    if (notes) this.removePart(notes);
    this.removePart(slide.path);
  }

  moveSlide(from: number, to: number): void {
    const ids = Array.from(this.parts.get(PRESENTATION)!.getElementsByTagNameNS(NS.p, "sldId"));
    if (from === to || from < 0 || to < 0 || from >= ids.length || to >= ids.length) return;
    const el = ids[from];
    const list = el.parentNode!;
    el.remove();
    const rest = ids.filter((e) => e !== el);
    list.insertBefore(el, rest[to] ?? null);
  }

  /** Store image bytes in ppt/media and relate them to the slide. Returns the new rId. */
  addImage(slidePath: string, bytes: Uint8Array, ext: string): string {
    return this.addRel(slidePath, "image", this.storeImage("ppt", bytes, ext));
  }

  private firstLayout(): string | null {
    const masterRel = this.rels(PRESENTATION).find((r) => r.type === "slideMaster");
    return masterRel ? this.relOfType(masterRel.target, "slideLayout") : null;
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
    list.insertBefore(el, ids[index] ?? null);
    return index;
  }
}
