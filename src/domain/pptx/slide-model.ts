import { NS, childNS } from "../../core/utils/xml";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { ColorContext, resolveColor } from "./color";
import { customGeometry, GeomPath, Geometry, Handle, presetGeometry, Rect, Site } from "./geometry";
import { Fill, fillCss, Line, readFill, readFillRef, readLine, readShadow, Shadow, ThemeStyles, themeStyles } from "./shape-style";
import { cellRect, readTable, TableCell, TableModel } from "./table";
import { Crop, readCrop } from "./slide-edit";

export type { GeomPath, Rect } from "./geometry";

export type Align = "l" | "ctr" | "r" | "just";

/** Geometry in EMU (914400 per inch). */
export interface RunProps {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  /** Points. */
  sz: number;
  color: string;
  font: string;
  /** Character spacing in points (a:rPr spc). */
  spc: number;
  /** Capitals (a:rPr cap): "all" shows the text in capitals, "small" in small capitals. */
  cap: "none" | "all" | "small";
  /** Hyperlink address (a:hlinkClick). */
  link: string | null;
  /** Smallest size (pt) PowerPoint kerns at (a:rPr kern); absent or 0: no kerning. */
  kern?: number;
}

export interface Run extends RunProps {
  /** "\n" marks an a:br line break. */
  text: string;
  /** a:fld type ("slidenum", "datetime1"...) for field runs. */
  field?: string;
}

/** Spacing as a fraction of single line spacing, or in points. */
export type Spacing = { pct: number } | { pts: number };

export interface Para {
  runs: Run[];
  align: Align;
  lvl: number;
  /** Bullet character, "#scheme[@start]" for automatic numbering (a:buAutoNum), or null for none. */
  bullet: string | null;
  /** EMU */
  marL: number;
  indent: number;
  /** Effective run properties for text without its own formatting. */
  defaults: RunProps;
  lnSpc?: Spacing | null;
  spcBef?: Spacing | null;
  spcAft?: Spacing | null;
  /** Font size (pt) of the end-of-paragraph mark; sizes empty paragraphs. */
  endSz?: number;
}

export interface TextBody {
  el: Element;
  paras: Para[];
  anchor: "t" | "ctr" | "b";
  insets: { l: number; t: number; r: number; b: number };
  /** a:normAutofit = shrink text on overflow, a:spAutoFit = resize shape to fit text. */
  autofit: "none" | "shrink" | "resize";
  /** Wrap text in the shape (bodyPr wrap != "none"). */
  wrap: boolean;
  /** normAutofit fontScale, 1 = 100%. */
  fontScale: number;
  /** normAutofit lnSpcReduction, 0.1 = line spacing 10 points of percentage lower. */
  lnSpcReduction: number;
  style: TextStyle;
}

export interface ShapeModel {
  el: Element;
  /** Zip part the element lives in. */
  part: string;
  /** Comes from the layout or master: render only, never edit. */
  inherited: boolean;
  /** Lives inside a group: text editable, geometry not. */
  grouped: boolean;
  /** "cell": one table cell's text, listed right after its table; el is the a:tc. */
  kind: "sp" | "pic" | "table" | "frame" | "cell";
  name: string;
  rect: Rect;
  /** Degrees clockwise. */
  rot: number;
  flipH: boolean;
  flipV: boolean;
  /** Is a slide placeholder (title, body...). */
  placeholder: boolean;
  /** Preset name, "custom" for a:custGeom, "" when the element has no geometry (pictures without one, frames). */
  geom: string;
  /** Preset adjust values set on the shape (a:avLst). */
  adj: Record<string, number>;
  /** Outline paths in shape coordinates; null = no geometry. */
  paths: GeomPath[] | null;
  /** Adjust handles and connection sites, in shape coordinates. */
  handles: Handle[];
  sites: Site[];
  fill: Fill | null;
  line: Line | null;
  shadow: Shadow | null;
  /** Media path of a picture. */
  image: string | null;
  /** Pictures: the part of the image cut away on each side (a:srcRect). */
  crop?: Crop | null;
  table: TableModel | null;
  /** For kind "cell": the cell, with its style-resolved fill and borders. */
  cell?: TableCell;
  text: TextBody | null;
  /** Text box relative to the shape, when it differs from the shape (SmartArt txXfrm). */
  textRect: Rect | null;
}

export interface SlideModel {
  path: string;
  background: { color: string | null; image: string | null };
  /** The slide's own fill background (p:bg/p:bgPr); null when it inherits or uses a picture / theme style. */
  ownBackground: Fill | null;
  shapes: ShapeModel[];
  notes: string;
}

/**
 * Everything needed to resolve effective text formatting: list-style
 * containers from most to least specific, plus theme colors and fonts.
 */
export interface TextStyle {
  /** Elements holding a:lvlNpPr children (a:lstStyle, p:titleStyle, ...). */
  chain: (Element | null)[];
  /** p:style fontRef color, slotted in before the master text styles. */
  fontRefColor: string | null;
  colors: ColorContext;
  fonts: { major: string; minor: string };
  /** Hyperlink relationships of the part the text lives in. */
  links?: { target(rId: string): string | null; add(url: string): string };
}

const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };
const NS_SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

/** Number of chain entries that outrank the shape's p:style fontRef color. */
const OWN_STYLE_DEPTH = 3;

const kids = (el: Element | null | undefined, ns: string, local: string): Element[] =>
  el ? Array.from(el.children).filter((c) => c.localName === local && c.namespaceURI === ns) : [];
const kid = (el: Element | null | undefined, ns: string, local: string): Element | null => (el ? childNS(el, ns, local) : null);
const num = (el: Element | null | undefined, attr: string): number | null =>
  el?.hasAttribute(attr) ? Number(el.getAttribute(attr)) : null;
const flag = (v: string | null) => v === "1" || v === "true";

// -----------------------------------------------------------------------------
// Text formatting resolution
// -----------------------------------------------------------------------------

function typeface(name: string | null | undefined, fonts: TextStyle["fonts"]): string | null {
  if (!name) return null;
  if (name.startsWith("+mj")) return fonts.major;
  if (name.startsWith("+mn")) return fonts.minor;
  return name;
}

/** Fill `props` from a:rPr / a:defRPr, never overwriting values already set. */
function mergeRunProps(props: Partial<RunProps>, rPr: Element | null, style: TextStyle): void {
  if (!rPr) return;
  if (props.b === undefined && rPr.hasAttribute("b")) props.b = flag(rPr.getAttribute("b"));
  if (props.i === undefined && rPr.hasAttribute("i")) props.i = flag(rPr.getAttribute("i"));
  if (props.u === undefined && rPr.hasAttribute("u")) props.u = rPr.getAttribute("u") !== "none";
  if (props.strike === undefined && rPr.hasAttribute("strike")) props.strike = rPr.getAttribute("strike") !== "noStrike";
  if (props.sz === undefined && rPr.hasAttribute("sz")) props.sz = Number(rPr.getAttribute("sz")) / 100;
  if (props.spc === undefined && rPr.hasAttribute("spc")) props.spc = Number(rPr.getAttribute("spc")) / 100;
  if (props.kern === undefined && rPr.hasAttribute("kern")) props.kern = Number(rPr.getAttribute("kern")) / 100;
  if (props.cap === undefined && rPr.hasAttribute("cap")) props.cap = (["all", "small"] as const).find((c) => c === rPr.getAttribute("cap")) ?? "none";
  const link = kid(rPr, NS.a, "hlinkClick")?.getAttributeNS(NS.r, "id");
  if (props.link === undefined && link) props.link = style.links?.target(link) ?? null;
  if (props.color === undefined) {
    const c = resolveColor(kid(rPr, NS.a, "solidFill"), style.colors);
    if (c) props.color = c;
  }
  if (props.font === undefined) {
    const f = typeface(kid(rPr, NS.a, "latin")?.getAttribute("typeface"), style.fonts);
    if (f) props.font = f;
  }
}

interface ResolvedLevel {
  align: Align;
  marL: number;
  indent: number;
  bullet: string | null;
  lnSpc: Spacing | null;
  spcBef: Spacing | null;
  spcAft: Spacing | null;
  defaults: RunProps;
}

/** Resolve paragraph-level formatting for `lvl`, with an optional a:pPr on top. */
export function resolveLevel(style: TextStyle, lvl: number, pPr: Element | null = null): ResolvedLevel {
  const tag = `lvl${lvl + 1}pPr`;
  const levels = [pPr, ...style.chain.map((c) => kid(c, NS.a, tag))];
  let align: Align | undefined;
  let marL: number | null = null;
  let indent: number | null = null;
  let bullet: string | null | undefined;
  const spacing: Record<"lnSpc" | "spcBef" | "spcAft", Spacing | null> = { lnSpc: null, spcBef: null, spcAft: null };
  const props: Partial<RunProps> = {};

  levels.forEach((l, i) => {
    if (i === OWN_STYLE_DEPTH + 1 && props.color === undefined && style.fontRefColor) props.color = style.fontRefColor;
    if (!l) return;
    align ??= (l.getAttribute("algn") as Align | null) ?? undefined;
    marL ??= num(l, "marL");
    indent ??= num(l, "indent");
    if (bullet === undefined) {
      if (kid(l, NS.a, "buNone")) bullet = null;
      else if (kid(l, NS.a, "buChar")) bullet = kid(l, NS.a, "buChar")!.getAttribute("char") ?? "•";
      else if (kid(l, NS.a, "buAutoNum")) {
        const auto = kid(l, NS.a, "buAutoNum")!;
        const start = auto.getAttribute("startAt");
        bullet = `#${auto.getAttribute("type") ?? "arabicPeriod"}${start ? `@${start}` : ""}`;
      }
    }
    for (const key of ["lnSpc", "spcBef", "spcAft"] as const) spacing[key] ??= readSpacing(kid(l, NS.a, key));
    mergeRunProps(props, kid(l, NS.a, "defRPr"), style);
  });
  if (props.color === undefined && style.fontRefColor) props.color = style.fontRefColor;

  return {
    align: align ?? "l",
    marL: marL ?? 0,
    indent: indent ?? 0,
    bullet: bullet ?? null,
    ...spacing,
    defaults: {
      b: props.b ?? false,
      i: props.i ?? false,
      u: props.u ?? false,
      strike: props.strike ?? false,
      sz: props.sz ?? 18,
      color: props.color ?? "#000000",
      font: props.font ?? style.fonts.minor,
      spc: props.spc ?? 0,
      cap: props.cap ?? "none",
      // Links belong to runs, never to level defaults.
      link: null,
    },
  };
}

function readSpacing(el: Element | null): Spacing | null {
  const pct = num(kid(el, NS.a, "spcPct"), "val");
  if (pct !== null) return { pct: pct / 100000 };
  const pts = num(kid(el, NS.a, "spcPts"), "val");
  return pts === null ? null : { pts: pts / 100 };
}

export function runProps(rPr: Element | null, base: RunProps, style: TextStyle): RunProps {
  const props: Partial<RunProps> = {};
  mergeRunProps(props, rPr, style);
  return { ...base, ...props };
}

export function readParagraphs(txBody: Element, style: TextStyle): Para[] {
  return kids(txBody, NS.a, "p").map((p) => {
    const pPr = kid(p, NS.a, "pPr");
    const lvl = num(pPr, "lvl") ?? 0;
    const level = resolveLevel(style, lvl, pPr);
    const runs: Run[] = [];
    for (const child of Array.from(p.children)) {
      if (child.localName === "r" || child.localName === "fld") {
        const run: Run = { ...runProps(kid(child, NS.a, "rPr"), level.defaults, style), text: kid(child, NS.a, "t")?.textContent ?? "" };
        if (child.localName === "fld") run.field = child.getAttribute("type") ?? "";
        runs.push(run);
      } else if (child.localName === "br") {
        runs.push({ ...runProps(kid(child, NS.a, "rPr"), level.defaults, style), text: "\n" });
      }
    }
    const endSz = runProps(kid(p, NS.a, "endParaRPr"), level.defaults, style).sz;
    return { runs, lvl, ...level, endSz };
  });
}

// -----------------------------------------------------------------------------
// Slide parsing
// -----------------------------------------------------------------------------

interface Ph {
  type: string;
  idx: string | null;
}

const phOf = (sp: Element): Ph | null => {
  const nv = Array.from(sp.children).find((c) => c.localName.startsWith("nv"));
  const ph = kid(kid(nv, NS.p, "nvPr"), NS.p, "ph");
  return ph ? { type: ph.getAttribute("type") ?? "obj", idx: ph.getAttribute("idx") } : null;
};

/** Collapse placeholder types to the families masters define. */
const phFamily = (type: string) =>
  type === "ctrTitle" || type === "title" ? "title" : ["dt", "ftr", "sldNum", "hdr"].includes(type) ? type : "body";

function findPlaceholder(part: Document | null, ph: Ph, master: boolean): Element | null {
  if (!part) return null;
  const sps = Array.from(part.getElementsByTagNameNS(NS.p, "sp")).map((sp) => [sp, phOf(sp)] as const);
  if (!master) {
    const hit =
      (ph.idx !== null && sps.find(([, p]) => p?.idx === ph.idx)) ||
      sps.find(([, p]) => p?.type === ph.type) ||
      sps.find(([, p]) => p && phFamily(p.type) === phFamily(ph.type));
    if (hit) return hit[0];
    return null;
  }
  return sps.find(([, p]) => p && phFamily(p.type) === phFamily(ph.type))?.[0] ?? null;
}

interface Transform {
  sx: number;
  sy: number;
  dx: number;
  dy: number;
}
const IDENTITY: Transform = { sx: 1, sy: 1, dx: 0, dy: 0 };

function readXfrm(xfrm: Element | null): { rect: Rect; rot: number; flipH: boolean; flipV: boolean } | null {
  const off = kid(xfrm, NS.a, "off");
  const ext = kid(xfrm, NS.a, "ext");
  if (!off || !ext) return null;
  return {
    rect: { x: num(off, "x")!, y: num(off, "y")!, w: num(ext, "cx")!, h: num(ext, "cy")! },
    rot: (num(xfrm, "rot") ?? 0) / 60000,
    flipH: flag(xfrm!.getAttribute("flipH")),
    flipV: flag(xfrm!.getAttribute("flipV")),
  };
}

const applyTransform = (r: Rect, t: Transform): Rect => ({
  x: Math.round(r.x * t.sx + t.dx),
  y: Math.round(r.y * t.sy + t.dy),
  w: Math.round(r.w * t.sx),
  h: Math.round(r.h * t.sy),
});

export function plainText(txBody: Element | null): string {
  return kids(txBody, NS.a, "p")
    .map((p) => Array.from(p.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join(""))
    .join("\n");
}

class SlideParser {
  layout: Document | null;
  master: Document | null;
  colors: ColorContext = { scheme: {}, map: {} };
  fonts = { major: "Calibri", minor: "Calibri" };
  shapes: ShapeModel[] = [];
  private bg: SlideModel["background"] = { color: "#FFFFFF", image: null };
  private bgFill: Fill | null = { type: "solid", color: "#FFFFFF" };
  private theme: ThemeStyles;
  /** Fill of the group being walked, for a:grpFill. */
  private groupFill: Fill | null = null;

  constructor(private doc: PptxDocument, private slidePath: string) {
    const layoutPath = doc.layoutOf(slidePath);
    const masterPath = layoutPath && doc.masterOf(layoutPath);
    const themePath = masterPath && doc.themeOf(masterPath);
    this.layout = layoutPath ? doc.xml(layoutPath) : null;
    this.master = masterPath ? doc.xml(masterPath) : null;

    const clrMap = this.master?.getElementsByTagNameNS(NS.p, "clrMap")[0];
    for (const a of Array.from(clrMap?.attributes ?? [])) this.colors.map[a.name] = a.value;
    // A slide (or else its layout) can remap tx1/bg1..., e.g. light text on a dark slide.
    const override = [doc.xml(slidePath), this.layout]
      .map((d) => d?.getElementsByTagNameNS(NS.a, "overrideClrMapping")[0])
      .find((o) => o);
    for (const a of Array.from(override?.attributes ?? [])) this.colors.map[a.name] = a.value;
    const theme = themePath ? doc.xml(themePath) : null;
    const scheme = theme?.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    for (const slot of Array.from(scheme?.children ?? [])) {
      const c = resolveColor(slot, { scheme: {}, map: {} });
      if (c) this.colors.scheme[slot.localName] = c;
    }
    const font = (tag: string) =>
      kid(theme?.getElementsByTagNameNS(NS.a, tag)[0], NS.a, "latin")?.getAttribute("typeface") || "Calibri";
    this.fonts = { major: font("majorFont"), minor: font("minorFont") };
    this.theme = themeStyles(theme);
  }

  parse(): SlideModel {
    const slide = this.doc.xml(this.slidePath)!;
    const layoutPath = this.doc.layoutOf(this.slidePath);
    const masterPath = layoutPath && this.doc.masterOf(layoutPath);
    const showMaster = (d: Document | null) => d?.documentElement.getAttribute("showMasterSp") !== "0";

    this.bg = this.background([
      [slide, this.slidePath],
      [this.layout, layoutPath],
      [this.master, masterPath],
    ]);
    if (this.master && masterPath && showMaster(this.layout) && showMaster(slide)) this.walkPart(this.master, masterPath);
    if (this.layout && layoutPath && showMaster(slide)) this.walkPart(this.layout, layoutPath);
    this.walkPart(slide, this.slidePath);

    const ownBgPr = kid(kid(slide.getElementsByTagNameNS(NS.p, "cSld")[0], NS.p, "bg"), NS.p, "bgPr");
    const notesPath = this.doc.notesOf(this.slidePath);
    const notesSp = Array.from(this.doc.xml(notesPath ?? "")?.getElementsByTagNameNS(NS.p, "sp") ?? []).find(
      (sp) => phOf(sp)?.type === "body"
    );

    return {
      path: this.slidePath,
      background: this.bg,
      ownBackground: ownBgPr ? readFill(ownBgPr, null, this.colors, this.theme) : null,
      shapes: this.shapes,
      notes: notesSp ? plainText(kid(notesSp, NS.p, "txBody")) : "",
    };
  }

  /** A blip's image: its SVG version when it has one (Office 2016+, sometimes with no bitmap at all), else the bitmap. */
  private blipImage(part: string, blip: Element | null): string | null {
    if (!blip) return null;
    const svg = blip.getElementsByTagNameNS(NS_SVG, "svgBlip")[0]?.getAttributeNS(NS.r, "embed");
    return (svg && this.doc.relTarget(part, svg)) || this.doc.relTarget(part, blip.getAttributeNS(NS.r, "embed") ?? "") || null;
  }

  private background(parts: [Document | null, string | null][]): SlideModel["background"] {
    for (const [part, path] of parts) {
      const bg = part && kid(part.getElementsByTagNameNS(NS.p, "cSld")[0], NS.p, "bg");
      if (!bg || !path) continue;
      const bgPr = kid(bg, NS.p, "bgPr");
      const blip = bgPr && kid(kid(bgPr, NS.a, "blipFill"), NS.a, "blip");
      if (blip) {
        const target = this.blipImage(path, blip);
        if (target) return { color: null, image: target };
      }
      const bgRef = kid(bg, NS.p, "bgRef");
      let fill = bgPr ? readFill(bgPr, null, this.colors, this.theme) : readFillRef(bgRef, this.colors, this.theme);
      const refColor = !bgPr && resolveColor(bgRef, this.colors);
      if (!fill && refColor) fill = { type: "solid", color: refColor };
      if (fill) {
        this.bgFill = fill;
        return { color: fillCss(fill), image: null };
      }
    }
    return { color: "#FFFFFF", image: null };
  }

  private walkPart(part: Document, path: string): void {
    const tree = part.getElementsByTagNameNS(NS.p, "spTree")[0];
    if (tree) this.walk(tree, path, IDENTITY, false, path !== this.slidePath);
  }

  private walk(tree: Element, part: string, t: Transform, grouped: boolean, inherited: boolean): void {
    for (const el of Array.from(tree.children)) {
      if (el.localName === "AlternateContent") {
        const fallback = Array.from(el.children).find((c) => c.localName === "Fallback");
        if (fallback) this.walk(fallback, part, t, grouped, inherited);
        continue;
      }
      if (el.namespaceURI !== NS.p && el.namespaceURI !== NS.dsp) continue;
      // Layout/master placeholders are prompts for the slide, not slide content.
      if (inherited && phOf(el)) continue;
      if (el.localName === "graphicFrame" && this.diagram(el, part, t)) continue;

      if (el.localName === "grpSp") {
        const xfrm = kid(kid(el, el.namespaceURI, "grpSpPr"), NS.a, "xfrm");
        const off = kid(xfrm, NS.a, "off");
        const ext = kid(xfrm, NS.a, "ext");
        const chOff = kid(xfrm, NS.a, "chOff");
        const chExt = kid(xfrm, NS.a, "chExt");
        let local = IDENTITY;
        if (off && ext && chOff && chExt) {
          const sx = num(chExt, "cx") ? num(ext, "cx")! / num(chExt, "cx")! : 1;
          const sy = num(chExt, "cy") ? num(ext, "cy")! / num(chExt, "cy")! : 1;
          local = { sx, sy, dx: num(off, "x")! - num(chOff, "x")! * sx, dy: num(off, "y")! - num(chOff, "y")! * sy };
        }
        const outer = this.groupFill;
        this.groupFill = readFill(kid(el, el.namespaceURI, "grpSpPr"), null, this.colors, this.theme, outer);
        this.walk(el, part, { sx: t.sx * local.sx, sy: t.sy * local.sy, dx: t.sx * local.dx + t.dx, dy: t.sy * local.dy + t.dy }, true, inherited);
        this.groupFill = outer;
      } else if (el.localName === "sp" || el.localName === "cxnSp" || el.localName === "pic" || el.localName === "graphicFrame") {
        const shape = this.shape(el, part, inherited);
        if (shape) {
          shape.rect = applyTransform(shape.rect, t);
          shape.grouped = grouped;
          this.shapes.push(shape);
          if (shape.table) this.shapes.push(...this.cells(shape));
        }
      }
    }
  }

  /**
   * SmartArt: draw the cached rendering PowerPoint stores next to the diagram
   * data (read-only, since PowerPoint rebuilds it from the data). False if there is none.
   */
  private diagram(frame: Element, part: string, t: Transform): boolean {
    const dm = frame.getElementsByTagNameNS(NS.dgm, "relIds")[0]?.getAttributeNS(NS.r, "dm");
    const dataPath = dm ? this.doc.relTarget(part, dm) : null;
    const relId = dataPath ? this.doc.xml(dataPath)?.getElementsByTagNameNS(NS.dsp, "dataModelExt")[0]?.getAttribute("relId") : null;
    const drawingPath = relId ? this.doc.relTarget(part, relId) : null;
    const tree = drawingPath ? this.doc.xml(drawingPath)?.getElementsByTagNameNS(NS.dsp, "spTree")[0] : null;
    const at = readXfrm(kid(frame, NS.p, "xfrm"));
    if (!tree || !drawingPath || !at) return false;
    this.walk(tree, drawingPath, { ...t, dx: t.dx + at.rect.x * t.sx, dy: t.dy + at.rect.y * t.sy }, false, true);
    return true;
  }

  /** Text shapes for a table's cells; covered (merged-away) cells have no text. */
  private cells(table: ShapeModel): ShapeModel[] {
    const txStyles = this.master?.getElementsByTagNameNS(NS.p, "txStyles")[0];
    const pres = this.doc.xml("ppt/presentation.xml");
    const t = table.table!;
    return t.cells.flat().map((cell) => {
      const r = cellRect(t, cell);
      const tcPr = kid(cell.el, NS.a, "tcPr");
      const txBody = kid(cell.el, NS.a, "txBody");
      const style: TextStyle = {
        chain: [
          kid(txBody, NS.a, "lstStyle"),
          cellTextStyle(cell),
          kid(txStyles, NS.p, "otherStyle"),
          pres?.getElementsByTagNameNS(NS.p, "defaultTextStyle")[0] ?? null,
        ],
        fontRefColor: null,
        colors: this.colors,
        fonts: this.fonts,
        links: { target: (id) => this.doc.relTarget(table.part, id), add: (url) => this.doc.linkRel(table.part, url) },
      };
      const inset = (attr: string, d: number) => num(tcPr, attr) ?? d;
      const anchor = tcPr?.getAttribute("anchor");
      return {
        ...table,
        el: cell.el,
        kind: "cell",
        name: `${table.name} cell ${cell.r + 1},${cell.c + 1}`,
        grouped: true,
        rect: { x: table.rect.x + r.x, y: table.rect.y + r.y, w: r.w, h: r.h },
        rot: 0,
        placeholder: false,
        paths: null,
        handles: [],
        sites: [],
        table: null,
        cell,
        fill: null,
        line: null,
        shadow: null,
        text:
          txBody && !cell.hidden
            ? {
                el: txBody,
                paras: readParagraphs(txBody, style),
                anchor: anchor === "ctr" || anchor === "b" ? anchor : "t",
                autofit: "none",
                wrap: true,
                insets: { l: inset("marL", 91440), r: inset("marR", 91440), t: inset("marT", 45720), b: inset("marB", 45720) },
                fontScale: 1,
                lnSpcReduction: 0,
                style,
              }
            : null,
      };
    });
  }

  private shape(el: Element, part: string, inherited: boolean): ShapeModel | null {
    const ns = el.namespaceURI!;
    const ph = phOf(el);
    const layoutPh = ph && !inherited ? findPlaceholder(this.layout, ph, false) : null;
    const masterPh = ph && !inherited ? findPlaceholder(this.master, ph, true) : null;
    const spPr = kid(el, ns, "spPr");
    const styleEl = kid(el, ns, "style");
    const nv = Array.from(el.children).find((c) => c.localName.startsWith("nv"));
    const cNvPr = kid(nv, ns, "cNvPr");

    const xfrmEl = el.localName === "graphicFrame" ? kid(el, NS.p, "xfrm") : kid(spPr, NS.a, "xfrm");
    const geo =
      readXfrm(xfrmEl) ??
      readXfrm(kid(kid(layoutPh, NS.p, "spPr"), NS.a, "xfrm")) ??
      readXfrm(kid(kid(masterPh, NS.p, "spPr"), NS.a, "xfrm"));
    if (!geo) return null;
    const prstGeom = kid(spPr, NS.a, "prstGeom");
    const custGeom = kid(spPr, NS.a, "custGeom");
    const geom = prstGeom?.getAttribute("prst") ?? (custGeom ? "custom" : el.localName === "sp" || el.localName === "cxnSp" ? "rect" : "");
    const adj: Record<string, number> = {};
    for (const gd of kids(kid(prstGeom, NS.a, "avLst"), NS.a, "gd")) {
      const v = /^val\s+(-?[\d.]+)/.exec(gd.getAttribute("fmla") ?? "");
      if (v) adj[gd.getAttribute("name") ?? ""] = Number(v[1]);
    }
    const { w, h } = geo.rect;
    const g: Geometry | null = custGeom ? customGeometry(custGeom, w, h) : geom ? presetGeometry(geom, w, h, adj) : null;

    const model: ShapeModel = {
      el,
      part,
      inherited,
      grouped: false,
      kind: "sp",
      name: cNvPr?.getAttribute("name") ?? "",
      rect: geo.rect,
      rot: geo.rot,
      flipH: geo.flipH,
      flipV: geo.flipV,
      placeholder: !!ph,
      geom,
      adj,
      paths: g?.paths ?? null,
      handles: g?.handles ?? [],
      sites: g?.sites ?? [],
      // useBgFill: "fill with the slide background" (PowerPoint Designer layouts use it a lot).
      fill: el.getAttribute("useBgFill") === "1" ? this.bgFill : readFill(spPr, styleEl, this.colors, this.theme, this.groupFill),
      line: readLine(spPr, styleEl, this.colors, this.theme),
      shadow: readShadow(spPr, styleEl, this.colors, this.theme),
      image: null,
      table: null,
      text: null,
      textRect: null,
    };

    if (el.localName === "pic") {
      model.kind = "pic";
      model.image = this.blipImage(part, kid(kid(el, ns, "blipFill"), NS.a, "blip"));
      model.crop = readCrop(el);
      model.fill = null;
      return model;
    }

    // A shape filled with a picture (SmartArt icons, pictures cropped to shapes).
    const pictureFill = kid(kid(spPr, NS.a, "blipFill"), NS.a, "blip");
    if (pictureFill) {
      model.image = this.blipImage(part, pictureFill);
      if (model.image) model.fill = null;
    }

    if (el.localName === "graphicFrame") {
      const tbl = el.getElementsByTagNameNS(NS.a, "tbl")[0];
      model.kind = tbl ? "table" : "frame";
      const styles = this.doc.xml(this.doc.relOfType("ppt/presentation.xml", "tableStyles") ?? "");
      model.table = tbl ? readTable(tbl, styles, this.colors) : null;
      model.fill = null;
      model.line = null;
      return model;
    }

    const txBody = kid(el, ns, "txBody");
    // Preset text area (inside an ellipse, triangle...), unless it is the whole shape.
    const tr = g?.textRect;
    if (tr && (tr.x || tr.y || Math.abs(tr.w - w) > 1 || Math.abs(tr.h - h) > 1)) model.textRect = tr;
    const tx = readXfrm(kid(el, NS.dsp, "txXfrm"));
    if (tx) model.textRect = { x: tx.rect.x - geo.rect.x, y: tx.rect.y - geo.rect.y, w: tx.rect.w, h: tx.rect.h };
    if (txBody) {
      const txStyles = this.master?.getElementsByTagNameNS(NS.p, "txStyles")[0];
      const family = ph ? phFamily(ph.type) : null;
      const txStyle = kid(txStyles, NS.p, family === "title" ? "titleStyle" : family === "body" ? "bodyStyle" : "otherStyle");
      const pres = this.doc.xml("ppt/presentation.xml");
      const style: TextStyle = {
        chain: [
          kid(txBody, NS.a, "lstStyle"),
          kid(kid(layoutPh, NS.p, "txBody"), NS.a, "lstStyle"),
          kid(kid(masterPh, NS.p, "txBody"), NS.a, "lstStyle"),
          txStyle,
          pres?.getElementsByTagNameNS(NS.p, "defaultTextStyle")[0] ?? null,
        ],
        fontRefColor: resolveColor(kid(styleEl, NS.a, "fontRef"), this.colors),
        colors: this.colors,
        fonts: this.fonts,
        links: { target: (id) => this.doc.relTarget(part, id), add: (url) => this.doc.linkRel(part, url) },
      };
      const bodyPrs = [txBody, kid(layoutPh, NS.p, "txBody"), kid(masterPh, NS.p, "txBody")].map((b) => kid(b, NS.a, "bodyPr"));
      const pick = (attr: string) => bodyPrs.find((b) => b?.hasAttribute(attr))?.getAttribute(attr) ?? null;
      const inset = (attr: string, d: number) => (pick(attr) === null ? d : Number(pick(attr)));
      const anchor = pick("anchor");
      model.text = {
        el: txBody,
        paras: readParagraphs(txBody, style),
        anchor: anchor === "ctr" || anchor === "b" ? anchor : "t",
        autofit: ((): TextBody["autofit"] => {
          const fit = bodyPrs.map((b) => Array.from(b?.children ?? []).find((c) => c.localName.endsWith("utofit") || c.localName === "spAutoFit")).find(Boolean);
          return fit?.localName === "normAutofit" ? "shrink" : fit?.localName === "spAutoFit" ? "resize" : "none";
        })(),
        wrap: pick("wrap") !== "none",
        insets: {
          l: inset("lIns", DEFAULT_INSETS.l),
          t: inset("tIns", DEFAULT_INSETS.t),
          r: inset("rIns", DEFAULT_INSETS.r),
          b: inset("bIns", DEFAULT_INSETS.b),
        },
        fontScale: (num(kid(bodyPrs[0], NS.a, "normAutofit"), "fontScale") ?? 100000) / 100000,
        lnSpcReduction: (num(kid(bodyPrs[0], NS.a, "normAutofit"), "lnSpcReduction") ?? 0) / 100000,
        style,
      };
    }
    return model;
  }
}

/** The table style's bold / color for a cell, as a list style slotted into the cell's text style chain. */
function cellTextStyle(cell: TableCell): Element | null {
  const { b, color } = cell.text;
  if (b === null && !color) return null;
  const rPr = `<a:defRPr${b === null ? "" : ` b="${b ? 1 : 0}"`}>${color ? `<a:solidFill><a:srgbClr val="${color.slice(1, 7)}"/></a:solidFill>` : ""}</a:defRPr>`;
  const lvls = Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr>${rPr}</a:lvl${i + 1}pPr>`).join("");
  return new DOMParser().parseFromString(`<a:lstStyle xmlns:a="${NS.a}">${lvls}</a:lstStyle>`, "application/xml").documentElement;
}

export function parseSlide(doc: PptxDocument, slidePath: string): SlideModel {
  return new SlideParser(doc, slidePath).parse();
}
