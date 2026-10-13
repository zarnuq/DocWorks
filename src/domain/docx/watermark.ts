/*
 * Watermarks, as Google Docs' Insert › Watermark: text or a picture behind the text on every page. Stored as Word
 * does: a VML shape (PowerPlusWaterMarkObject text path / WordPictureWatermark picture), centred on the margins,
 * in a run of the first paragraph of every header, so Word, LibreOffice and Google Docs show it too.
 */
import { NS } from "../../core/utils/xml";
import type { DocxDocument } from "../../infrastructure/docx/docx-document";

export const VML = "urn:schemas-microsoft-com:vml";
const O = "urn:schemas-microsoft-com:office:office";
const W10 = "urn:schemas-microsoft-com:office:word";

export interface TextWatermark {
  kind: "text";
  text: string;
  font: string;
  /** Box size in points (Word stretches the text to fill it). */
  width: number;
  height: number;
  color: string;
  /** 1 = solid. */
  opacity: number;
  /** 45° up, else horizontal. */
  diagonal: boolean;
}

export interface ImageWatermark {
  kind: "image";
  /** New picture bytes, or `path`: one already in the package. */
  data?: Uint8Array;
  ext?: string;
  path?: string;
  width: number;
  height: number;
  /** Faded, as Word's "Washout". */
  washout: boolean;
}

export type Watermark = TextWatermark | ImageWatermark;

const isWatermark = (shape: Element) => /^(PowerPlusWaterMarkObject|WordPictureWatermark)/.test(shape.getAttribute("id") ?? "");
const COLORS: Record<string, string> = { silver: "#C0C0C0", gray: "#808080", grey: "#808080", black: "#000000", white: "#FFFFFF", red: "#FF0000", blue: "#0000FF" };

function styleOf(el: Element): Map<string, string> {
  return new Map(
    (el.getAttribute("style") ?? "")
      .split(";")
      .map((d) => d.split(":").map((x) => x.trim()) as [string, string])
      .filter(([k]) => k)
  );
}

const pt = (v: string | undefined) => (v ? Number.parseFloat(v) * (v.endsWith("in") ? 72 : v.endsWith("px") ? 0.75 : 1) : 0);

/** VML fractions: ".5", "50%", or 16.16 fixed point "32768f". */
function fraction(v: string | null): number {
  if (!v) return 1;
  if (v.endsWith("f")) return Number.parseFloat(v) / 65536;
  if (v.endsWith("%")) return Number.parseFloat(v) / 100;
  return Number.parseFloat(v);
}

function watermarkShapes(docx: DocxDocument): { path: string; shape: Element }[] {
  return docx.headerFooterPaths("header").flatMap((path) =>
    Array.from(docx.xml(path)?.getElementsByTagNameNS(VML, "shape") ?? []).filter(isWatermark).map((shape) => ({ path, shape }))
  );
}

export function readWatermark(docx: DocxDocument): Watermark | null {
  const found = watermarkShapes(docx)[0];
  if (!found) return null;
  const { path, shape } = found;
  const style = styleOf(shape);
  const width = pt(style.get("width")), height = pt(style.get("height"));
  const textpath = shape.getElementsByTagNameNS(VML, "textpath")[0];
  if (textpath) {
    const fill = shape.getAttribute("fillcolor") ?? "silver";
    const color = (COLORS[fill.toLowerCase()] ?? (/^#[0-9a-f]{6}$/i.test(fill) ? fill : "#C0C0C0")).toUpperCase();
    const font = (styleOf(textpath).get("font-family") ?? "Calibri").replace(/^["']|["']$/g, "");
    const opacity = fraction(shape.getElementsByTagNameNS(VML, "fill")[0]?.getAttribute("opacity") ?? null);
    const rotation = Number(style.get("rotation") ?? 0);
    return { kind: "text", text: textpath.getAttribute("string") ?? "", font, width, height, color, opacity: +opacity.toFixed(3), diagonal: rotation % 360 !== 0 };
  }
  const data = shape.getElementsByTagNameNS(VML, "imagedata")[0];
  const rId = data?.getAttributeNS(NS.r, "id") ?? "";
  return { kind: "image", path: (rId && docx.relTarget(path, rId)) || undefined, width, height, washout: !!data?.getAttribute("gain") };
}

/** Remove the watermark runs from every header. */
function clear(docx: DocxDocument): void {
  for (const { path, shape } of watermarkShapes(docx)) {
    let run: Element | null = shape;
    while (run && !(run.localName === "r" && run.namespaceURI === NS.w)) run = run.parentElement;
    (run ?? shape).remove();
    docx.touch(path);
  }
}

const TEXT_SHAPETYPE =
  `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas>` +
  ["sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3", "if @0 @3 0", "if @0 21600 @1", "if @0 0 @2", "if @0 @4 21600", "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5"].map((e) => `<v:f eqn="${e}"/>`).join("") +
  `</v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>` +
  `<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>`;

const PICTURE_SHAPETYPE =
  `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas>` +
  ["if lineDrawn pixelLineWidth 0", "sum @0 1 0", "sum 0 0 @1", "prod @2 1 2", "prod @3 21600 pixelWidth", "prod @3 21600 pixelHeight", "sum @0 0 1", "prod @6 1 2", "prod @7 21600 pixelWidth", "sum @8 21600 0", "prod @7 21600 pixelHeight", "sum @10 21600 0"].map((e) => `<v:f eqn="${e}"/>`).join("") +
  `</v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>`;

const escapeAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const num = (n: number) => String(+n.toFixed(2));

function runXml(wm: Watermark, n: number, rId: string): string {
  const place = (extra: string) =>
    `position:absolute;margin-left:0;margin-top:0;width:${num(wm.width)}pt;height:${num(wm.height)}pt;${extra}z-index:-251657216;` +
    `mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin`;
  const shape =
    wm.kind === "text"
      ? `${TEXT_SHAPETYPE}<v:shape id="PowerPlusWaterMarkObject${n}" o:spid="_x0000_s${2048 + n}" type="#_x0000_t136" style="${place(wm.diagonal ? "rotation:315;" : "")}" o:allowincell="f" fillcolor="${wm.color}" stroked="f">` +
        `<v:fill opacity="${num(wm.opacity)}"/><v:textpath style="font-family:&quot;${escapeAttr(wm.font)}&quot;;font-size:1pt" string="${escapeAttr(wm.text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape>`
      : `${PICTURE_SHAPETYPE}<v:shape id="WordPictureWatermark${n}" o:spid="_x0000_s${2048 + n}" type="#_x0000_t75" style="${place("")}" o:allowincell="f">` +
        `<v:imagedata r:id="${rId}" o:title=""${wm.washout ? ` gain="19661f" blacklevel="22938f"` : ""}/><w10:wrap anchorx="margin" anchory="margin"/></v:shape>`;
  return `<w:r xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:v="${VML}" xmlns:o="${O}" xmlns:w10="${W10}"><w:rPr><w:noProof/></w:rPr><w:pict>${shape}</w:pict></w:r>`;
}

/** Set (or with null remove) the watermark on every page: every header of every section gets it. */
export function setWatermark(docx: DocxDocument, wm: Watermark | null): void {
  const image = wm?.kind === "image" ? wm.path ?? (wm.data ? docx.storeMedia(wm.data, wm.ext ?? "png") : null) : null;
  clear(docx);
  if (!wm || (wm.kind === "image" && !image)) return;
  // Every page needs a header to carry it.
  docx.sections(true).forEach((_, i) => {
    if (!docx.headerFooterPath("header", "default", i)) docx.ensureHeaderFooter("header", "default", i);
  });
  docx.headerFooterPaths("header").forEach((path, k) => {
    const hdr = docx.xml(path)!.documentElement;
    let p = Array.from(hdr.children).find((c) => c.localName === "p" && c.namespaceURI === NS.w);
    if (!p) p = hdr.insertBefore(hdr.ownerDocument.createElementNS(NS.w, "w:p"), hdr.firstChild);
    const rId = image ? docx.imageRel(path, image) : "";
    const run = hdr.ownerDocument.importNode(new DOMParser().parseFromString(runXml(wm, k + 1, rId), "application/xml").documentElement, true);
    const pPr = Array.from(p.children).find((c) => c.localName === "pPr");
    p.insertBefore(run, pPr ? pPr.nextSibling : p.firstChild);
    docx.touch(path);
  });
}
