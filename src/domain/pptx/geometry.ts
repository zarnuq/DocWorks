/*
 * DrawingML geometry: evaluates preset (preset-shapes.ts) and custom (a:custGeom)
 * shape definitions into SVG path data, text rectangles, connection sites and
 * adjust handles, all in shape coordinates (EMU, origin top-left).
 */
import { NS } from "../../core/utils/xml";
import { PRESETS } from "./preset-shapes";

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export type PathFill = "norm" | "none" | "lighten" | "lightenLess" | "darken" | "darkenLess";

/** One path, in shape coordinates (w x h = the shape size). */
export interface GeomPath {
  w: number;
  h: number;
  /** SVG path data. */
  d: string;
  fill: PathFill;
  stroke: boolean;
}

/** Connection site: where connectors attach; `ang` in degrees, the direction a connector leaves. */
export interface Site {
  x: number;
  y: number;
  ang: number;
}

/** Adjust handle: dragging it changes the `refX`/`refY` adjust values within their ranges. */
export interface Handle {
  x: number;
  y: number;
  refX: string | null;
  minX: number;
  maxX: number;
  refY: string | null;
  minY: number;
  maxY: number;
}

export interface Geometry {
  paths: GeomPath[];
  textRect: Rect;
  sites: Site[];
  handles: Handle[];
}

export const PRESET_NAMES = Object.keys(PRESETS);

/** DrawingML angles are in 60000ths of a degree. */
const ANGLE = 60000;
const rad = (a: number) => (a / ANGLE) * (Math.PI / 180);
const deg = (r: number) => (r * 180) / Math.PI;
const fmt = (n: number) => String(Math.round(n * 1000) / 1000);

const OPS: Record<string, (x: number, y: number, z: number) => number> = {
  "*/": (x, y, z) => (z ? (x * y) / z : 0),
  "+-": (x, y, z) => x + y - z,
  "+/": (x, y, z) => (z ? (x + y) / z : 0),
  "?:": (x, y, z) => (x > 0 ? y : z),
  abs: (x) => Math.abs(x),
  at2: (x, y) => deg(Math.atan2(y, x)) * ANGLE,
  cat2: (x, y, z) => x * Math.cos(Math.atan2(z, y)),
  cos: (x, y) => x * Math.cos(rad(y)),
  max: (x, y) => Math.max(x, y),
  min: (x, y) => Math.min(x, y),
  mod: (x, y, z) => Math.sqrt(x * x + y * y + z * z),
  pin: (x, y, z) => (y < x ? x : y > z ? z : y),
  sat2: (x, y, z) => x * Math.sin(Math.atan2(z, y)),
  sin: (x, y) => x * Math.sin(rad(y)),
  sqrt: (x) => Math.sqrt(Math.max(0, x)),
  tan: (x, y) => x * Math.tan(rad(y)),
  val: (x) => x,
};

/** Guide values: the spec's built-in names plus avLst/gdLst formulas, evaluated in order. */
class Guides {
  private vars: Record<string, number>;

  constructor(w: number, h: number) {
    const ss = Math.min(w, h);
    this.vars = {
      w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls: Math.max(w, h),
      cd2: 10800000, cd4: 5400000, cd8: 2700000, "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000,
    };
    // wd2, hd10, ssd32...: w, h or ss divided by n.
    for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
      Object.assign(this.vars, { [`wd${n}`]: w / n, [`hd${n}`]: h / n, [`ssd${n}`]: ss / n });
    }
  }

  get = (token: string | null | undefined): number => {
    if (!token) return 0;
    return token in this.vars ? this.vars[token] : Number(token) || 0;
  };

  define(name: string, fmla: string): void {
    const [op, ...args] = fmla.trim().split(/\s+/);
    this.vars[name] = (OPS[op] ?? (() => 0))(this.get(args[0]), this.get(args[1]), this.get(args[2]));
  }

  has = (name: string) => name in this.vars;
}

/** Path commands: [op, ...tokens], op one of M L A C Q Z. */
type Command = string[];

interface PathSpec {
  w: number | null;
  h: number | null;
  fill: PathFill;
  stroke: boolean;
  cmds: Command[];
}

/** Point on an ellipse at DrawingML angle `a` (a "visual" angle, not the parametric one). */
function ellipsePoint(wR: number, hR: number, a: number): [number, number] {
  const t = Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  return [wR * Math.cos(t), hR * Math.sin(t)];
}

/** SVG path data; points in a path with its own w/h are scaled by sx, sy into shape space. */
function pathData(spec: PathSpec, g: Guides, sx: number, sy: number): string {
  let cx = 0;
  let cy = 0;
  const out: string[] = [];
  for (const [op, ...t] of spec.cmds) {
    // Every command's values alternate x, y (arcTo: wR, hR, then angles).
    const v = t.map((tok, i) => g.get(tok) * (op === "A" && i > 1 ? 1 : i % 2 ? sy : sx));
    if (op === "M" || op === "L") {
      [cx, cy] = v;
      out.push(`${op}${fmt(cx)} ${fmt(cy)}`);
    } else if (op === "C" || op === "Q") {
      out.push(`${op}${v.map(fmt).join(" ")}`);
      [cx, cy] = v.slice(-2);
    } else if (op === "Z") {
      out.push("Z");
    } else if (op === "A") {
      const [wR, hR, st, sw] = v;
      const [sx, sy] = ellipsePoint(wR, hR, rad(st));
      const ox = cx - sx;
      const oy = cy - sy;
      // Split into pieces of at most 90 degrees so SVG never sees an ambiguous arc.
      const n = Math.max(1, Math.ceil(Math.abs(sw) / (90 * ANGLE)));
      for (let i = 1; i <= n; i++) {
        const [px, py] = ellipsePoint(wR, hR, rad(st + (sw * i) / n));
        cx = ox + px;
        cy = oy + py;
        out.push(`A${fmt(wR)} ${fmt(hR)} 0 0 ${sw > 0 ? 1 : 0} ${fmt(cx)} ${fmt(cy)}`);
      }
    }
  }
  return out.join(" ");
}

interface Definition {
  av: [string, string][];
  gd: [string, string][];
  rect: string[] | null;
  sites: string[][];
  handles: string[][];
  paths: PathSpec[];
}

function build(def: Definition, w: number, h: number, adj: Record<string, number>): Geometry {
  const g = new Guides(w, h);
  for (const [name, fmla] of def.av) {
    if (name in adj) g.define(name, `val ${adj[name]}`);
    else g.define(name, fmla);
  }
  for (const [name, fmla] of def.gd) g.define(name, fmla);

  const [l, t, r, b] = def.rect?.map(g.get) ?? [0, 0, w, h];
  const ref = (s: string) => (s && g.has(s) ? s : null);
  return {
    paths: def.paths.map((p) => ({
      w,
      h,
      d: pathData(p, g, p.w ? w / p.w : 1, p.h ? h / p.h : 1),
      fill: p.fill,
      stroke: p.stroke,
    })),
    textRect: { x: l, y: t, w: r - l, h: b - t },
    sites: def.sites.map(([ang, x, y]) => ({ x: g.get(x), y: g.get(y), ang: g.get(ang) / ANGLE })),
    handles: def.handles.map(([refX, minX, maxX, refY, minY, maxY, x, y]) => ({
      x: g.get(x),
      y: g.get(y),
      refX: ref(refX),
      minX: g.get(minX),
      maxX: g.get(maxX),
      refY: ref(refY),
      minY: g.get(minY),
      maxY: g.get(maxY),
    })),
  };
}

const pair = (s: string): [string, string] => {
  const i = s.indexOf("=");
  return [s.slice(0, i), s.slice(i + 1)];
};
const list = (s: string, sep = ";") => (s ? s.split(sep) : []);

const parsed = new Map<string, Definition>();

function presetDefinition(name: string): Definition | null {
  const src = PRESETS[name];
  if (!src) return null;
  if (!parsed.has(name)) {
    const [av, gd, rect, cxn, ah, paths] = src.split("~");
    parsed.set(name, {
      av: list(av).map(pair),
      gd: list(gd).map(pair),
      rect: rect ? rect.split(" ") : null,
      sites: list(cxn).map((c) => c.split(" ")),
      handles: list(ah).map((a) => a.split(",")),
      paths: list(paths, "|").map((p) => {
        const [attrs, ...cmds] = p.split(";");
        const [pw, ph, fill, stroke] = attrs.split(",");
        return {
          w: pw ? Number(pw) : null,
          h: ph ? Number(ph) : null,
          fill: (fill || "norm") as PathFill,
          stroke: stroke !== "false" && stroke !== "0",
          cmds: cmds.map((c) => c.split(" ")),
        };
      }),
    });
  }
  return parsed.get(name)!;
}

export function presetGeometry(name: string, w: number, h: number, adj: Record<string, number> = {}): Geometry | null {
  const def = presetDefinition(name);
  return def ? build(def, w, h, adj) : null;
}

const kids = (el: Element | null | undefined, local: string) =>
  Array.from(el?.children ?? []).filter((c) => c.namespaceURI === NS.a && c.localName === local);
const gds = (el: Element | null | undefined): [string, string][] =>
  kids(el, "gd").map((gd) => [gd.getAttribute("name") ?? "", gd.getAttribute("fmla") ?? ""]);

const TAGS: Record<string, string> = { moveTo: "M", lnTo: "L", cubicBezTo: "C", quadBezTo: "Q", close: "Z", arcTo: "A" };

export function customGeometry(custGeom: Element, w: number, h: number): Geometry {
  const rect = kids(custGeom, "rect")[0];
  return build(
    {
      av: gds(kids(custGeom, "avLst")[0]),
      gd: gds(kids(custGeom, "gdLst")[0]),
      rect: rect ? ["l", "t", "r", "b"].map((k) => rect.getAttribute(k) ?? "0") : null,
      sites: kids(kids(custGeom, "cxnLst")[0], "cxn").map((c) => {
        const pos = kids(c, "pos")[0];
        return [c.getAttribute("ang") ?? "0", pos?.getAttribute("x") ?? "0", pos?.getAttribute("y") ?? "0"];
      }),
      handles: [],
      paths: kids(kids(custGeom, "pathLst")[0], "path").map((p) => ({
        w: p.hasAttribute("w") ? Number(p.getAttribute("w")) : null,
        h: p.hasAttribute("h") ? Number(p.getAttribute("h")) : null,
        fill: (p.getAttribute("fill") ?? "norm") as PathFill,
        stroke: !["0", "false"].includes(p.getAttribute("stroke") ?? ""),
        cmds: Array.from(p.children).map((c) => {
          const op = TAGS[c.localName] ?? "Z";
          if (op === "A") return [op, ...["wR", "hR", "stAng", "swAng"].map((k) => c.getAttribute(k) ?? "0")];
          return [op, ...kids(c, "pt").flatMap((pt) => [pt.getAttribute("x") ?? "0", pt.getAttribute("y") ?? "0"])];
        }),
      })),
    },
    w,
    h,
    {}
  );
}

/**
 * Adjust values that move handle `index` of a preset as close as possible to
 * `target` (shape coordinates). Solved per axis by bisection over the handle's
 * range, assuming the handle moves monotonically with its adjust value.
 */
export function solveAdjust(
  name: string,
  w: number,
  h: number,
  adj: Record<string, number>,
  index: number,
  target: { x: number; y: number }
): Record<string, number> {
  const out = { ...adj };
  const handle = presetGeometry(name, w, h, out)?.handles[index];
  if (!handle) return out;
  for (const axis of ["x", "y"] as const) {
    const ref = axis === "x" ? handle.refX : handle.refY;
    if (!ref) continue;
    const pos = (v: number) => presetGeometry(name, w, h, { ...out, [ref]: v })!.handles[index][axis];
    let lo = axis === "x" ? handle.minX : handle.minY;
    let hi = axis === "x" ? handle.maxX : handle.maxY;
    const rising = pos(hi) >= pos(lo);
    for (let i = 0; i < 48; i++) {
      const mid = (lo + hi) / 2;
      if (pos(mid) < target[axis] === rising) lo = mid;
      else hi = mid;
    }
    out[ref] = Math.round((lo + hi) / 2);
  }
  return out;
}
