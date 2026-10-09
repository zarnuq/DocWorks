/*
 * DrawingML tables (a:tbl in a p:graphicFrame): reading the grid with its table
 * style resolved per cell, and grid edits that keep merges and the frame size valid.
 * Every row holds one a:tc per grid column; merged-away cells carry hMerge / vMerge.
 */
import { NS, escapeXml } from "../../core/utils/xml";
import type { ColorContext } from "./color";
import { resolveColor } from "./color";
import type { Rect } from "./geometry";
import { Fill, Line, lineFrom, readFill, writeFill } from "./shape-style";
import { appendShape, nextShapeId } from "./slide-edit";

export type TableFlags = Record<"firstRow" | "lastRow" | "firstCol" | "lastCol" | "bandRow" | "bandCol", boolean>;
export type Side = "l" | "r" | "t" | "b";

export interface TableCell {
  el: Element;
  r: number;
  c: number;
  rowSpan: number;
  gridSpan: number;
  /** Covered by a merged cell: not drawn. */
  hidden: boolean;
  fill: Fill | null;
  borders: Record<Side, Line | null>;
  /** Bold / color from the table style's text style, when it sets them. */
  text: { b: boolean | null; color: string | null };
}

export interface TableModel {
  /** Column widths and row heights, EMU. */
  cols: number[];
  rows: number[];
  cells: TableCell[][];
  flags: TableFlags;
  styleId: string | null;
}

/** Inclusive cell range [r1, c1, r2, c2]. */
export type CellRange = [number, number, number, number];

const A = NS.a;
const FLAGS: (keyof TableFlags)[] = ["firstRow", "lastRow", "firstCol", "lastCol", "bandRow", "bandCol"];
const TCPR_ORDER = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "fill", "headers", "extLst"];
const SIDE_TAG: Record<Side, string> = { l: "lnL", r: "lnR", t: "lnT", b: "lnB" };

const kids = (el: Element | null | undefined, local: string) => Array.from(el?.children ?? []).filter((c) => c.localName === local && c.namespaceURI === A);
const kid = (el: Element | null | undefined, local: string) => kids(el, local)[0] ?? null;
const int = (el: Element, attr: string) => Number(el.getAttribute(attr)) || 0;

const rowsOf = (tbl: Element) => kids(tbl, "tr");
const gridCols = (tbl: Element) => kids(kid(tbl, "tblGrid"), "gridCol");
const cellAt = (tbl: Element, r: number, c: number): Element | null => kids(rowsOf(tbl)[r], "tc")[c] ?? null;
const flag = (tc: Element | null, attr: string) => tc?.getAttribute(attr) === "1" || tc?.getAttribute(attr) === "true";
const span = (tc: Element, attr: string) => Math.max(1, int(tc, attr));

function setSpan(tc: Element, attr: string, n: number): void {
  if (n > 1) tc.setAttribute(attr, String(n));
  else tc.removeAttribute(attr);
}

function setFlag(tc: Element, attr: string, on: boolean): void {
  if (on) tc.setAttribute(attr, "1");
  else tc.removeAttribute(attr);
}

/** Top-left cell of the merged block covering (r, c). */
function originOf(tbl: Element, r: number, c: number): [number, number] {
  while (r > 0 && flag(cellAt(tbl, r, c), "vMerge")) r--;
  while (c > 0 && flag(cellAt(tbl, r, c), "hMerge")) c--;
  return [r, c];
}

// -----------------------------------------------------------------------------
// Reading
// -----------------------------------------------------------------------------

interface Region {
  part: Element;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/** Style parts that apply to (r, c), lowest priority first (ECMA-376 21.1.3.11). */
function regions(style: Element | null, f: TableFlags, r: number, c: number, nr: number, nc: number): Region[] {
  if (!style) return [];
  const out: Region[] = [];
  const add = (name: string, r0: number, r1: number, c0: number, c1: number) => {
    const part = kid(style, name);
    if (part && r >= r0 && r <= r1 && c >= c0 && c <= c1) out.push({ part, r0, r1, c0, c1 });
  };
  const top = f.firstRow ? 1 : 0;
  const bottom = nr - 1 - (f.lastRow ? 1 : 0);
  const left = f.firstCol ? 1 : 0;
  const right = nc - 1 - (f.lastCol ? 1 : 0);
  add("wholeTbl", 0, nr - 1, 0, nc - 1);
  if (f.bandRow && r >= top && r <= bottom) add((r - top) % 2 ? "band2H" : "band1H", r, r, 0, nc - 1);
  if (f.bandCol && c >= left && c <= right) add((c - left) % 2 ? "band2V" : "band1V", 0, nr - 1, c, c);
  if (f.lastCol) add("lastCol", 0, nr - 1, nc - 1, nc - 1);
  if (f.firstCol) add("firstCol", 0, nr - 1, 0, 0);
  if (f.lastRow) add("lastRow", nr - 1, nr - 1, 0, nc - 1);
  if (f.firstRow) add("firstRow", 0, 0, 0, nc - 1);
  return out;
}

/** The border element a region gives side `s` of cell (r, c): its outer edge, or insideH / insideV within it. */
function regionBorder(g: Region, s: Side, r: number, c: number): Element | null {
  const bdr = kid(kid(g.part, "tcStyle"), "tcBdr");
  const edge = s === "t" ? r === g.r0 : s === "b" ? r === g.r1 : s === "l" ? c === g.c0 : c === g.c1;
  const name = edge ? { l: "left", r: "right", t: "top", b: "bottom" }[s] : s === "t" || s === "b" ? "insideH" : "insideV";
  return kid(bdr, name);
}

export function readTable(tbl: Element, styles: Document | null, ctx: ColorContext): TableModel {
  const tblPr = kid(tbl, "tblPr");
  const flags = Object.fromEntries(FLAGS.map((k) => [k, flag(tblPr, k)])) as TableFlags;
  const styleId = kid(tblPr, "tableStyleId")?.textContent?.trim() || null;
  const style = styleId
    ? (Array.from(styles?.getElementsByTagNameNS(A, "tblStyle") ?? []).find((s) => s.getAttribute("styleId") === styleId) ?? null)
    : null;
  const rows = rowsOf(tbl);
  const cols = gridCols(tbl).map((g) => int(g, "w"));
  const nr = rows.length;
  const nc = cols.length;

  const cells = rows.map((tr, r) =>
    kids(tr, "tc").map((el, c): TableCell => {
      const applied = regions(style, flags, r, c, nr, nc);
      let fill: Fill | null = null;
      const text: TableCell["text"] = { b: null, color: null };
      for (const g of applied) {
        const f = kid(kid(g.part, "tcStyle"), "fill");
        if (f) fill = readFill(f, null, ctx, null);
        const tx = kid(g.part, "tcTxStyle");
        if (tx) {
          const b = tx.getAttribute("b");
          if (b === "on" || b === "off") text.b = b === "on";
          // The direct color child wins over the fontRef's.
          const color = resolveColor(Array.from(tx.children).find((ch) => ch.localName !== "fontRef") ?? null, ctx) ?? resolveColor(kid(tx, "fontRef"), ctx);
          if (color) text.color = color;
        }
      }
      const tcPr = kid(el, "tcPr");
      const own = tcPr && Array.from(tcPr.children).find((ch) => /^(noFill|solidFill|gradFill|pattFill)$/.test(ch.localName));
      if (own) fill = readFill(tcPr, null, ctx, null);

      const borders = {} as TableCell["borders"];
      for (const s of ["l", "r", "t", "b"] as Side[]) {
        let line: Line | null = null;
        for (const g of applied) {
          const ln = kid(regionBorder(g, s, r, c), "ln");
          if (ln) line = lineFrom(ln, ctx);
        }
        const mine = kid(tcPr, SIDE_TAG[s]);
        if (mine) line = lineFrom(mine, ctx);
        borders[s] = line;
      }
      return {
        el,
        r,
        c,
        rowSpan: span(el, "rowSpan"),
        gridSpan: span(el, "gridSpan"),
        hidden: flag(el, "hMerge") || flag(el, "vMerge"),
        fill,
        borders,
        text,
      };
    })
  );
  return { cols, rows: rows.map((tr) => int(tr, "h")), cells, flags, styleId };
}

/** A cell's rectangle relative to the table's top-left corner. */
export function cellRect(t: TableModel, cell: TableCell): Rect {
  const sum = (list: number[], from: number, n: number) => list.slice(from, from + n).reduce((a, b) => a + b, 0);
  return {
    x: sum(t.cols, 0, cell.c),
    y: sum(t.rows, 0, cell.r),
    w: sum(t.cols, cell.c, cell.hidden ? 1 : cell.gridSpan),
    h: sum(t.rows, cell.r, cell.hidden ? 1 : cell.rowSpan),
  };
}

// -----------------------------------------------------------------------------
// Editing
// -----------------------------------------------------------------------------

const EMPTY_BODY = `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>`;

function parseA(doc: Document, xml: string): Element {
  const el = new DOMParser().parseFromString(`<r xmlns:a="${A}">${xml}</r>`, "application/xml").documentElement.firstElementChild!;
  return doc.importNode(el, true);
}

/** An empty cell carrying `like`'s formatting (tcPr), without its merge. */
function newCell(like: Element | null, doc: Document): Element {
  const tc = parseA(doc, `<a:tc>${EMPTY_BODY}</a:tc>`);
  const tcPr = kid(like, "tcPr");
  tc.appendChild(tcPr ? tcPr.cloneNode(true) : parseA(doc, "<a:tcPr/>"));
  return tc;
}

/** Keep the graphic frame's size equal to the grid. */
function syncFrame(tbl: Element): void {
  const frame = tbl.closest("graphicFrame");
  const ext = frame && Array.from(frame.children).find((c) => c.localName === "xfrm")?.getElementsByTagNameNS(A, "ext")[0];
  if (!ext) return;
  ext.setAttribute("cx", String(gridCols(tbl).reduce((n, g) => n + int(g, "w"), 0)));
  ext.setAttribute("cy", String(rowsOf(tbl).reduce((n, tr) => n + int(tr, "h"), 0)));
}

/** Grow or shrink the merge covering cell (r, c) by `by`, once per merge (`seen` holds merges already changed). */
function growMerge(tbl: Element, seen: Set<string>, r: number, c: number, attr: "rowSpan" | "gridSpan", by: number): void {
  const [or, oc] = originOf(tbl, r, c);
  if (seen.has(`${or},${oc}`)) return;
  seen.add(`${or},${oc}`);
  const origin = cellAt(tbl, or, oc)!;
  setSpan(origin, attr, span(origin, attr) + by);
}

/** Insert a row above or below row `at`, copying its formatting; a row inside a vertical merge joins it. */
export function insertRow(tbl: Element, at: number, below: boolean): void {
  const src = rowsOf(tbl)[at];
  if (!src) return;
  const doc = tbl.ownerDocument;
  const next = below ? at + 1 : at;
  const tr = src.cloneNode(false) as Element;
  const grown = new Set<string>();
  kids(src, "tc").forEach((like, c) => {
    const tc = tr.appendChild(newCell(like, doc));
    // Inside a vertical merge when the cell the new row pushes down continues one.
    if (flag(cellAt(tbl, next, c), "vMerge")) {
      setFlag(tc, "vMerge", true);
      setFlag(tc, "hMerge", flag(like, "hMerge"));
      growMerge(tbl, grown, next, c, "rowSpan", 1);
    }
  });
  tbl.insertBefore(tr, rowsOf(tbl)[next] ?? null);
  syncFrame(tbl);
}

export function insertCol(tbl: Element, at: number, right: boolean): void {
  const cols = gridCols(tbl);
  if (!cols[at]) return;
  const doc = tbl.ownerDocument;
  const next = right ? at + 1 : at;
  const col = cols[at].cloneNode(false) as Element;
  cols[at].parentNode!.insertBefore(col, cols[next] ?? null);
  const grown = new Set<string>();
  rowsOf(tbl).forEach((tr, r) => {
    const like = cellAt(tbl, r, at);
    const tc = newCell(like, doc);
    if (flag(cellAt(tbl, r, next), "hMerge")) {
      setFlag(tc, "hMerge", true);
      setFlag(tc, "vMerge", flag(like, "vMerge"));
      growMerge(tbl, grown, r, next, "gridSpan", 1);
    }
    tr.insertBefore(tc, kids(tr, "tc")[next] ?? null);
  });
  syncFrame(tbl);
}

/** Remove row `r`; merges crossing it shrink, and a merge starting in it moves its content down. */
export function deleteRow(tbl: Element, r: number): void {
  const rows = rowsOf(tbl);
  if (!rows[r] || rows.length <= 1) return;
  const shrunk = new Set<string>();
  kids(rows[r], "tc").forEach((tc, c) => {
    if (flag(tc, "hMerge") && !flag(tc, "vMerge")) return;
    if (flag(tc, "vMerge")) growMerge(tbl, shrunk, r, c, "rowSpan", -1);
    else if (span(tc, "rowSpan") > 1) {
      const heir = cellAt(tbl, r + 1, c)!;
      setSpan(heir, "rowSpan", span(tc, "rowSpan") - 1);
      setSpan(heir, "gridSpan", span(tc, "gridSpan"));
      for (let k = 0; k < span(tc, "gridSpan"); k++) setFlag(cellAt(tbl, r + 1, c + k)!, "vMerge", false);
      const body = kid(tc, "txBody");
      if (body) kid(heir, "txBody")?.replaceWith(body);
    }
  });
  rows[r].remove();
  syncFrame(tbl);
}

export function deleteCol(tbl: Element, c: number): void {
  const cols = gridCols(tbl);
  if (!cols[c] || cols.length <= 1) return;
  const shrunk = new Set<string>();
  rowsOf(tbl).forEach((tr, r) => {
    const tc = kids(tr, "tc")[c];
    if (!tc || (flag(tc, "vMerge") && !flag(tc, "hMerge"))) return;
    if (flag(tc, "hMerge")) growMerge(tbl, shrunk, r, c, "gridSpan", -1);
    else if (span(tc, "gridSpan") > 1) {
      const heir = cellAt(tbl, r, c + 1)!;
      setSpan(heir, "gridSpan", span(tc, "gridSpan") - 1);
      setSpan(heir, "rowSpan", span(tc, "rowSpan"));
      for (let k = 0; k < span(tc, "rowSpan"); k++) setFlag(cellAt(tbl, r + k, c + 1)!, "hMerge", false);
      const body = kid(tc, "txBody");
      if (body) kid(heir, "txBody")?.replaceWith(body);
    }
  });
  rowsOf(tbl).forEach((tr) => kids(tr, "tc")[c]?.remove());
  cols[c].remove();
  syncFrame(tbl);
}

const hasText = (body: Element | null): body is Element => !!body?.textContent?.trim();

/** Merge a range into its top-left cell, collecting the other cells' text as extra paragraphs. */
export function mergeCells(tbl: Element, range: CellRange): void {
  const [r1, c1, r2, c2] = normalize(range);
  if (r1 === r2 && c1 === c2) return;
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) splitCell(tbl, r, c);
  const origin = cellAt(tbl, r1, c1)!;
  const body = kid(origin, "txBody");
  if (!body) return;
  if (!hasText(body)) kids(body, "p").forEach((p) => p.remove());
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      if (r === r1 && c === c1) continue;
      const tc = cellAt(tbl, r, c)!;
      const other = kid(tc, "txBody");
      if (hasText(other)) {
        kids(other, "p").forEach((p) => body.appendChild(p));
        other.replaceWith(parseA(tbl.ownerDocument, EMPTY_BODY));
      }
      setFlag(tc, "hMerge", c > c1);
      setFlag(tc, "vMerge", r > r1);
    }
  }
  if (!kids(body, "p").length) body.appendChild(parseA(tbl.ownerDocument, `<a:p><a:endParaRPr lang="en-US"/></a:p>`));
  setSpan(origin, "gridSpan", c2 - c1 + 1);
  setSpan(origin, "rowSpan", r2 - r1 + 1);
}

/** Undo the merge covering (r, c). */
export function splitCell(tbl: Element, r: number, c: number): void {
  const [or, oc] = originOf(tbl, r, c);
  const origin = cellAt(tbl, or, oc);
  if (!origin) return;
  const rs = span(origin, "rowSpan");
  const gs = span(origin, "gridSpan");
  for (let i = 0; i < rs; i++) {
    for (let j = 0; j < gs; j++) {
      const tc = cellAt(tbl, or + i, oc + j);
      if (tc) ["hMerge", "vMerge"].forEach((a) => tc.removeAttribute(a));
    }
  }
  origin.removeAttribute("rowSpan");
  origin.removeAttribute("gridSpan");
}

const normalize = ([r1, c1, r2, c2]: CellRange): CellRange => [Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2)];

function tcPrOf(tc: Element): Element {
  return kid(tc, "tcPr") ?? tc.appendChild(parseA(tc.ownerDocument, "<a:tcPr/>"));
}

/** Give cells their own fill; null goes back to the table style's. */
export function setCellFill(tbl: Element, cells: [number, number][], fill: Fill | null): void {
  for (const [r, c] of cells) {
    const tc = cellAt(tbl, r, c);
    if (tc) writeFill(tcPrOf(tc), fill ?? undefined, TCPR_ORDER);
  }
}

export type BorderMode = "all" | "outer" | "inside";

/** Set borders on a range; `line` null draws none. */
export function setCellBorders(tbl: Element, range: CellRange, mode: BorderMode, line: { color: string; width: number } | null): void {
  const [r1, c1, r2, c2] = normalize(range);
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const tc = cellAt(tbl, r, c);
      if (!tc) continue;
      const outer: Record<Side, boolean> = { t: r === r1, b: r === r2, l: c === c1, r: c === c2 };
      for (const s of ["l", "r", "t", "b"] as Side[]) {
        if (mode === "outer" ? !outer[s] : mode === "inside" ? outer[s] : false) continue;
        const tcPr = tcPrOf(tc);
        kid(tcPr, SIDE_TAG[s])?.remove();
        const xml = line
          ? `<a:${SIDE_TAG[s]} w="${Math.round(line.width)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${line.color.slice(1, 7).toUpperCase()}"/></a:solidFill><a:prstDash val="solid"/></a:${SIDE_TAG[s]}>`
          : `<a:${SIDE_TAG[s]}><a:noFill/></a:${SIDE_TAG[s]}>`;
        const el = parseA(tc.ownerDocument, xml);
        const at = TCPR_ORDER.indexOf(SIDE_TAG[s]);
        const before = Array.from(tcPr.children).find((ch) => {
          const i = TCPR_ORDER.indexOf(/Fill$/.test(ch.localName) ? "fill" : ch.localName);
          return i > at;
        });
        tcPr.insertBefore(el, before ?? null);
      }
    }
  }
}

export function setTableFlags(tbl: Element, flags: Partial<TableFlags>): void {
  const tblPr = kid(tbl, "tblPr") ?? tbl.insertBefore(parseA(tbl.ownerDocument, "<a:tblPr/>"), tbl.firstChild);
  for (const [k, on] of Object.entries(flags)) setFlag(tblPr, k, !!on);
}

export function setTableStyleId(tbl: Element, id: string | null): void {
  const tblPr = kid(tbl, "tblPr") ?? tbl.insertBefore(parseA(tbl.ownerDocument, "<a:tblPr/>"), tbl.firstChild);
  kid(tblPr, "tableStyleId")?.remove();
  if (id) tblPr.appendChild(parseA(tbl.ownerDocument, `<a:tableStyleId>${escapeXml(id)}</a:tableStyleId>`));
}

export function setColWidth(tbl: Element, c: number, w: number): void {
  gridCols(tbl)[c]?.setAttribute("w", String(Math.max(1, Math.round(w))));
  syncFrame(tbl);
}

export function setRowHeight(tbl: Element, r: number, h: number): void {
  rowsOf(tbl)[r]?.setAttribute("h", String(Math.max(1, Math.round(h))));
  syncFrame(tbl);
}

/** Scale column widths and row heights so the grid fills `w` x `h` (frame resizes). */
export function fitTable(tbl: Element, w: number, h: number): void {
  const scale = (els: Element[], attr: string, total: number) => {
    const sum = els.reduce((n, e) => n + int(e, attr), 0);
    if (!sum || Math.round(sum) === Math.round(total)) return;
    let used = 0;
    els.forEach((e, i) => {
      const v = i === els.length - 1 ? Math.round(total) - used : Math.round((int(e, attr) * total) / sum);
      e.setAttribute(attr, String(v));
      used += v;
    });
  };
  scale(gridCols(tbl), "w", w);
  scale(rowsOf(tbl), "h", h);
}

/** A new table frame with an even grid. */
export function addTable(slide: Document, rows: number, cols: number, rect: Rect, styleId: string | null): Element {
  const id = nextShapeId(slide);
  const even = (total: number, n: number) => Array.from({ length: n }, (_, i) => Math.round((total * (i + 1)) / n) - Math.round((total * i) / n));
  const ws = even(rect.w, cols);
  const hs = even(rect.h, rows);
  const style = styleId ? `<a:tableStyleId>${escapeXml(styleId)}</a:tableStyleId>` : "";
  const row = (h: number) => `<a:tr h="${h}">${`<a:tc>${EMPTY_BODY}<a:tcPr/></a:tc>`.repeat(cols)}</a:tr>`;
  return appendShape(
    slide,
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="${Math.round(rect.x)}" y="${Math.round(rect.y)}"/><a:ext cx="${Math.round(rect.w)}" cy="${Math.round(rect.h)}"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1">${style}</a:tblPr>` +
      `<a:tblGrid>${ws.map((w) => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>${hs.map(row).join("")}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

// -----------------------------------------------------------------------------
// Built-in styles written into ppt/tableStyles.xml when used
// -----------------------------------------------------------------------------

const ln = (w: number, clr: string) => `<a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="${clr}"/></a:solidFill></a:ln>`;
const sides = (names: string[], w: number, clr: string) => names.map((n) => `<a:${n}>${ln(w, clr)}</a:${n}>`).join("");
const solid = (clr: string, tint?: number) => `<a:fill><a:solidFill><a:schemeClr val="${clr}">${tint ? `<a:tint val="${tint}"/>` : ""}</a:schemeClr></a:solidFill></a:fill>`;
const txt = (clr: string, bold = false) => `<a:tcTxStyle${bold ? ' b="on"' : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="${clr}"/></a:tcTxStyle>`;
const ALL = ["left", "right", "top", "bottom", "insideH", "insideV"];

/** PowerPoint's "Medium Style 2" with color `c` (accentN or dk1). */
function medium2(id: string, name: string, c: string): string {
  const strong = (part: string, bdr = "") => `<a:${part}>${txt("lt1", true)}<a:tcStyle><a:tcBdr>${bdr}</a:tcBdr>${solid(c)}</a:tcStyle></a:${part}>`;
  const band = (part: string, fill = "") => `<a:${part}><a:tcStyle><a:tcBdr/>${fill}</a:tcStyle></a:${part}>`;
  return (
    `<a:tblStyle styleId="${id}" styleName="${name}"><a:wholeTbl>${txt("dk1")}<a:tcStyle><a:tcBdr>${sides(ALL, 12700, "lt1")}</a:tcBdr>${solid(c, 20000)}</a:tcStyle></a:wholeTbl>` +
    band("band1H", solid(c, 40000)) + band("band2H") + band("band1V", solid(c, 40000)) + band("band2V") +
    strong("lastCol") + strong("firstCol") + strong("lastRow", `<a:top>${ln(38100, "lt1")}</a:top>`) + strong("firstRow", `<a:bottom>${ln(38100, "lt1")}</a:bottom>`) +
    `</a:tblStyle>`
  );
}

export const TABLE_STYLES: { id: string; name: string; xml: string }[] = [
  ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", "Medium Style 2", "dk1"],
  ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", "Medium Style 2 - Accent 1", "accent1"],
  ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}", "Medium Style 2 - Accent 2", "accent2"],
  ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}", "Medium Style 2 - Accent 3", "accent3"],
  ["{00A15C55-8517-42AA-B614-E9B94910E393}", "Medium Style 2 - Accent 4", "accent4"],
  ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}", "Medium Style 2 - Accent 5", "accent5"],
  ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}", "Medium Style 2 - Accent 6", "accent6"],
]
  .map(([id, name, c]) => ({ id, name, xml: medium2(id, name, c) }))
  .concat({
    id: "{5940675A-B579-460E-94D1-54222C63F5DA}",
    name: "No Style, Table Grid",
    xml:
      `<a:tblStyle styleId="{5940675A-B579-460E-94D1-54222C63F5DA}" styleName="No Style, Table Grid"><a:wholeTbl>${txt("tx1")}` +
      `<a:tcStyle><a:tcBdr>${sides(ALL, 12700, "tx1")}</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>`,
  });
