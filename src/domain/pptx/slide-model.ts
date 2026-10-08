import { NS, childNS } from "../../core/utils/xml";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { ColorContext, resolveColor } from "./color";

export type Align = "l" | "ctr" | "r" | "just";

/** Geometry in EMU (914400 per inch). */
export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface RunProps {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  /** Points. */
  sz: number;
  color: string;
  font: string;
}

export interface Run extends RunProps {
  /** "\n" marks an a:br line break. */
  text: string;
}

/** Spacing as a fraction of single line spacing, or in points. */
export type Spacing = { pct: number } | { pts: number };

export interface Para {
  runs: Run[];
  align: Align;
  lvl: number;
  /** Bullet character, or null for none. */
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
  /** normAutofit fontScale, 1 = 100%. */
  fontScale: number;
  /** normAutofit lnSpcReduction, 0.1 = line spacing 10 points of percentage lower. */
  lnSpcReduction: number;
  style: TextStyle;
}

/** One a:path of a custom (or line) geometry, in its own w x h coordinate space. */
export interface GeomPath {
  w: number;
  h: number;
  /** SVG path data. */
  d: string;
  fill: boolean;
  stroke: boolean;
}

export interface ShapeModel {
  el: Element;
  /** Zip part the element lives in. */
  part: string;
  /** Comes from the layout or master: render only, never edit. */
  inherited: boolean;
  /** Lives inside a group: text editable, geometry not. */
  grouped: boolean;
  kind: "sp" | "pic" | "table" | "frame";
  name: string;
  rect: Rect;
  /** Degrees clockwise. */
  rot: number;
  flipH: boolean;
  flipV: boolean;
  /** Is a slide placeholder (title, body...). */
  placeholder: boolean;
  geom: string;
  /** Vector outline for custom geometry and lines; null = draw the box. */
  paths: GeomPath[] | null;
  /** CSS background: a color ("#RRGGBB[AA]") or a gradient. */
  fill: string | null;
  line: { color: string; width: number } | null;
  /** Media path of a picture. */
  image: string | null;
  table: string[][] | null;
  text: TextBody | null;
  /** Text box relative to the shape, when it differs from the shape (SmartArt txXfrm). */
  textRect: Rect | null;
}

export interface SlideModel {
  path: string;
  background: { color: string | null; image: string | null };
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
}

const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };
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
      else if (kid(l, NS.a, "buAutoNum")) bullet = "1.";
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
        runs.push({ ...runProps(kid(child, NS.a, "rPr"), level.defaults, style), text: kid(child, NS.a, "t")?.textContent ?? "" });
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

const LINE_GEOMS = new Set(["line", "straightConnector1"]);

const fmt = (n: number) => String(Math.round(n * 1000) / 1000);
const poly = (...pts: [number, number][]) => pts.map(([x, y], i) => `${i ? "L" : "M"}${fmt(x)} ${fmt(y)}`).join(" ") + " Z";

/**
 * Outline of common preset shapes in shape coordinates (w x h), using the
 * spec's default adjust values unless given. Null = drawn with CSS (rect, ellipse...).
 */
// ponytail: the common outline presets only; others fall back to a box. Add from presetShapeDefinitions.xml as decks need them.
export function presetPath(prst: string, w: number, h: number, adj: number[]): string | null {
  const ss = Math.min(w, h);
  const a = (i: number, def: number) => (Number.isFinite(adj[i]) ? adj[i] : def) / 100000;
  const hc = w / 2;
  const vc = h / 2;
  switch (prst) {
    case "chevron": {
      const x1 = ss * a(0, 50000);
      return poly([0, 0], [w - x1, 0], [w, vc], [w - x1, h], [0, h], [x1, vc]);
    }
    case "homePlate": {
      const x1 = ss * a(0, 50000);
      return poly([0, 0], [w - x1, 0], [w, vc], [w - x1, h], [0, h]);
    }
    case "triangle":
      return poly([w * a(0, 50000), 0], [w, h], [0, h]);
    case "rtTriangle":
      return poly([0, 0], [w, h], [0, h]);
    case "diamond":
      return poly([hc, 0], [w, vc], [hc, h], [0, vc]);
    case "parallelogram": {
      const x = ss * a(0, 25000);
      return poly([x, 0], [w, 0], [w - x, h], [0, h]);
    }
    case "trapezoid": {
      const x = ss * a(0, 25000);
      return poly([0, h], [x, 0], [w - x, 0], [w, h]);
    }
    case "hexagon": {
      const x = ss * a(0, 25000);
      return poly([0, vc], [x, 0], [w - x, 0], [w, vc], [w - x, h], [x, h]);
    }
    case "octagon": {
      const x = ss * a(0, 29289);
      return poly([x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]);
    }
    case "pentagon":
      return poly([hc, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]);
    case "plus": {
      const x = ss * a(0, 25000);
      return poly([x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]);
    }
    case "rightArrow":
    case "leftArrow": {
      const dy = (h * a(0, 50000)) / 2;
      const x = w - ss * a(1, 50000);
      const pts: [number, number][] = [[0, vc - dy], [x, vc - dy], [x, 0], [w, vc], [x, h], [x, vc + dy], [0, vc + dy]];
      return poly(...(prst === "rightArrow" ? pts : pts.map(([px, py]): [number, number] => [w - px, py])));
    }
    case "downArrow":
    case "upArrow": {
      const dx = (w * a(0, 50000)) / 2;
      const y = h - ss * a(1, 50000);
      const pts: [number, number][] = [[hc - dx, 0], [hc - dx, y], [0, y], [hc, h], [w, y], [hc + dx, y], [hc + dx, 0]];
      return poly(...(prst === "downArrow" ? pts : pts.map(([px, py]): [number, number] => [px, h - py])));
    }
    default:
      return null;
  }
}

// ponytail: arcTo is skipped (drawn as a straight jump); add elliptical arcs if decks look wrong.
/**
 * Evaluate a custGeom's avLst + gdLst formulas ("*\/ a b c", "+- a b c", ...)
 * against the shape size. Returns a lookup for names and literals.
 */
// ponytail: covers the arithmetic operators PowerPoint writes for freeforms; trig ops (at2, cat2, sin...) read as 0.
function guides(custGeom: Element | null, rect: Rect): (token: string | null) => number {
  const { w, h } = rect;
  const vars: Record<string, number> = {
    w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2,
    wd2: w / 2, hd2: h / 2, wd4: w / 4, hd4: h / 4, wd8: w / 8, hd8: h / 8,
    ss: Math.min(w, h), ls: Math.max(w, h), ssd2: Math.min(w, h) / 2, ssd4: Math.min(w, h) / 4,
  };
  const val = (token: string | null): number => (token === null ? 0 : token in vars ? vars[token] : Number(token) || 0);
  const ops: Record<string, (a: number, b: number, c: number) => number> = {
    "*/": (a, b, c) => (c ? (a * b) / c : 0),
    "+-": (a, b, c) => a + b - c,
    "+/": (a, b, c) => (c ? (a + b) / c : 0),
    "?:": (a, b, c) => (a > 0 ? b : c),
    val: (a) => a,
    abs: (a) => Math.abs(a),
    sqrt: (a) => Math.sqrt(Math.max(0, a)),
    min: (a, b) => Math.min(a, b),
    max: (a, b) => Math.max(a, b),
    pin: (a, b, c) => (b < a ? a : b > c ? c : b),
  };
  for (const list of ["avLst", "gdLst"]) {
    for (const gd of kids(kid(custGeom, NS.a, list), NS.a, "gd")) {
      const [op, ...args] = (gd.getAttribute("fmla") ?? "").trim().split(/\s+/);
      const name = gd.getAttribute("name");
      if (name) vars[name] = (ops[op] ?? (() => 0))(val(args[0] ?? null), val(args[1] ?? null), val(args[2] ?? null));
    }
  }
  return val;
}

function geometryPaths(spPr: Element | null, geom: string, rect: Rect): GeomPath[] | null {
  if (LINE_GEOMS.has(geom)) return [{ w: 1, h: 1, d: "M0 0 L1 1", fill: false, stroke: true }];
  const prstGeom = kid(spPr, NS.a, "prstGeom");
  if (prstGeom) {
    const adj = kids(kid(prstGeom, NS.a, "avLst"), NS.a, "gd")
      .sort((x, y) => (x.getAttribute("name") ?? "").localeCompare(y.getAttribute("name") ?? ""))
      .map((gd) => Number(/^val (-?\d+)/.exec(gd.getAttribute("fmla") ?? "")?.[1] ?? NaN));
    const d = presetPath(geom, rect.w, rect.h, adj);
    return d ? [{ w: rect.w, h: rect.h, d, fill: true, stroke: true }] : null;
  }
  const custGeom = kid(spPr, NS.a, "custGeom");
  const pathLst = kid(custGeom, NS.a, "pathLst");
  if (!pathLst) return null;
  const val = guides(custGeom, rect);
  return kids(pathLst, NS.a, "path").map((path) => {
    const pt = (p: Element) => `${Math.round(val(p.getAttribute("x")))} ${Math.round(val(p.getAttribute("y")))}`;
    const cmds = Array.from(path.children).flatMap((c) => {
      const pts = kids(c, NS.a, "pt").map(pt);
      switch (c.localName) {
        case "moveTo":
          return [`M${pts[0]}`];
        case "lnTo":
          return [`L${pts[0]}`];
        case "cubicBezTo":
          return [`C${pts.join(" ")}`];
        case "quadBezTo":
          return [`Q${pts.join(" ")}`];
        case "close":
          return ["Z"];
        default:
          return [];
      }
    });
    return {
      w: num(path, "w") ?? rect.w,
      h: num(path, "h") ?? rect.h,
      d: cmds.join(" "),
      fill: path.getAttribute("fill") !== "none",
      stroke: !["0", "false"].includes(path.getAttribute("stroke") ?? ""),
    };
  });
}

function gradientCss(grad: Element, colors: ColorContext, flipH = false, flipV = false): string | null {
  const stops = Array.from(grad.getElementsByTagNameNS(NS.a, "gs"))
    .map((gs) => ({ pos: (num(gs, "pos") ?? 0) / 1000, color: resolveColor(gs, colors) }))
    .filter((s) => s.color)
    .sort((a, b) => a.pos - b.pos)
    .map((s) => `${s.color} ${s.pos}%`);
  if (!stops.length) return null;
  if (kid(grad, NS.a, "path")) return `radial-gradient(${stops.join(", ")})`;
  // DrawingML angles run clockwise from "left to right"; CSS 0deg points up.
  // A flipped shape mirrors its gradient (CSS rotate() only covers rotation).
  let ang = (num(kid(grad, NS.a, "lin"), "ang") ?? 0) / 60000;
  if (flipH) ang = 180 - ang;
  if (flipV) ang = -ang;
  const deg = (((ang + 90) % 360) + 360) % 360;
  return `linear-gradient(${deg}deg, ${stops.join(", ")})`;
}

const applyTransform = (r: Rect, t: Transform): Rect => ({
  x: Math.round(r.x * t.sx + t.dx),
  y: Math.round(r.y * t.sy + t.dy),
  w: Math.round(r.w * t.sx),
  h: Math.round(r.h * t.sy),
});

function fillColor(spPr: Element | null, styleEl: Element | null, colors: ColorContext, flips = { flipH: false, flipV: false }): string | null {
  if (kid(spPr, NS.a, "noFill")) return null;
  const solid = kid(spPr, NS.a, "solidFill");
  if (solid) return resolveColor(solid, colors);
  const grad = kid(spPr, NS.a, "gradFill");
  if (grad) return gradientCss(grad, colors, flips.flipH, flips.flipV);
  const ref = kid(styleEl, NS.a, "fillRef");
  return ref && ref.getAttribute("idx") !== "0" ? resolveColor(ref, colors) : null;
}

function lineOf(spPr: Element | null, styleEl: Element | null, colors: ColorContext): ShapeModel["line"] {
  const ln = kid(spPr, NS.a, "ln");
  if (kid(ln, NS.a, "noFill")) return null;
  const width = num(ln, "w") ?? 12700;
  const solid = kid(ln, NS.a, "solidFill");
  if (solid) {
    const color = resolveColor(solid, colors);
    return color ? { color, width } : null;
  }
  const ref = kid(styleEl, NS.a, "lnRef");
  const color = ref && ref.getAttribute("idx") !== "0" ? resolveColor(ref, colors) : null;
  return color ? { color, width } : null;
}

function plainText(txBody: Element | null): string {
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

  constructor(private doc: PptxDocument, private slidePath: string) {
    const layoutPath = doc.layoutOf(slidePath);
    const masterPath = layoutPath && doc.masterOf(layoutPath);
    const themePath = masterPath && doc.themeOf(masterPath);
    this.layout = layoutPath ? doc.xml(layoutPath) : null;
    this.master = masterPath ? doc.xml(masterPath) : null;

    const clrMap = this.master?.getElementsByTagNameNS(NS.p, "clrMap")[0];
    for (const a of Array.from(clrMap?.attributes ?? [])) this.colors.map[a.name] = a.value;
    const theme = themePath ? doc.xml(themePath) : null;
    const scheme = theme?.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    for (const slot of Array.from(scheme?.children ?? [])) {
      const c = resolveColor(slot, { scheme: {}, map: {} });
      if (c) this.colors.scheme[slot.localName] = c;
    }
    const font = (tag: string) =>
      kid(theme?.getElementsByTagNameNS(NS.a, tag)[0], NS.a, "latin")?.getAttribute("typeface") || "Calibri";
    this.fonts = { major: font("majorFont"), minor: font("minorFont") };
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

    const notesPath = this.doc.notesOf(this.slidePath);
    const notesSp = Array.from(this.doc.xml(notesPath ?? "")?.getElementsByTagNameNS(NS.p, "sp") ?? []).find(
      (sp) => phOf(sp)?.type === "body"
    );

    return {
      path: this.slidePath,
      background: this.bg,
      shapes: this.shapes,
      notes: notesSp ? plainText(kid(notesSp, NS.p, "txBody")) : "",
    };
  }

  private background(parts: [Document | null, string | null][]): SlideModel["background"] {
    for (const [part, path] of parts) {
      const bg = part && kid(part.getElementsByTagNameNS(NS.p, "cSld")[0], NS.p, "bg");
      if (!bg || !path) continue;
      const bgPr = kid(bg, NS.p, "bgPr");
      const blip = bgPr && kid(kid(bgPr, NS.a, "blipFill"), NS.a, "blip");
      if (blip) {
        const target = this.doc.relTarget(path, blip.getAttributeNS(NS.r, "embed") ?? "");
        if (target) return { color: null, image: target };
      }
      const color = bgPr ? fillColor(bgPr, null, this.colors) : resolveColor(kid(bg, NS.p, "bgRef"), this.colors);
      if (color) return { color, image: null };
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
        this.walk(el, part, { sx: t.sx * local.sx, sy: t.sy * local.sy, dx: t.sx * local.dx + t.dx, dy: t.sy * local.dy + t.dy }, true, inherited);
      } else if (el.localName === "sp" || el.localName === "cxnSp" || el.localName === "pic" || el.localName === "graphicFrame") {
        const shape = this.shape(el, part, inherited);
        if (shape) {
          shape.rect = applyTransform(shape.rect, t);
          shape.grouped = grouped;
          this.shapes.push(shape);
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
    const geom = kid(spPr, NS.a, "prstGeom")?.getAttribute("prst") ?? (kid(spPr, NS.a, "custGeom") ? "custom" : "rect");

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
      paths: geometryPaths(spPr, geom, geo.rect),
      // useBgFill: "fill with the slide background" (PowerPoint Designer layouts use it a lot).
      fill: el.getAttribute("useBgFill") === "1" ? this.bg.color : fillColor(spPr, styleEl, this.colors, geo),
      line: lineOf(spPr, styleEl, this.colors),
      image: null,
      table: null,
      text: null,
      textRect: null,
    };

    if (el.localName === "pic") {
      model.kind = "pic";
      const blip = kid(kid(el, ns, "blipFill"), NS.a, "blip");
      model.image = this.doc.relTarget(part, blip?.getAttributeNS(NS.r, "embed") ?? "");
      model.fill = null;
      return model;
    }

    if (el.localName === "graphicFrame") {
      const tbl = el.getElementsByTagNameNS(NS.a, "tbl")[0];
      model.kind = tbl ? "table" : "frame";
      model.table = tbl
        ? kids(tbl, NS.a, "tr").map((tr) => kids(tr, NS.a, "tc").map((tc) => plainText(kid(tc, NS.a, "txBody"))))
        : null;
      return model;
    }

    const txBody = kid(el, ns, "txBody");
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
      };
      const bodyPrs = [txBody, kid(layoutPh, NS.p, "txBody"), kid(masterPh, NS.p, "txBody")].map((b) => kid(b, NS.a, "bodyPr"));
      const pick = (attr: string) => bodyPrs.find((b) => b?.hasAttribute(attr))?.getAttribute(attr) ?? null;
      const inset = (attr: string, d: number) => (pick(attr) === null ? d : Number(pick(attr)));
      const anchor = pick("anchor");
      model.text = {
        el: txBody,
        paras: readParagraphs(txBody, style),
        anchor: anchor === "ctr" || anchor === "b" ? anchor : "t",
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

export function parseSlide(doc: PptxDocument, slidePath: string): SlideModel {
  return new SlideParser(doc, slidePath).parse();
}
