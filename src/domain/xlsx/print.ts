/**
 * Print settings, as Google Sheets' print screen offers them and Excel stores them in the sheet (pageSetup,
 * printOptions, pageMargins, sheetPr/pageSetUpPr), and the pages a range splits into: down, then across.
 */
import type { Edits } from "./commands";
import type { Sheet } from "./workbook-model";
import { child } from "./sheet-edit";

const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export type Paper = "letter" | "a4" | "legal";
export interface PrintSettings {
  paper: Paper;
  landscape: boolean;
  /** "normal": at `scale` %, across as many pages as needed; "width": every column on one page; "page": all on one. */
  fit: "normal" | "width" | "page";
  scale: number;
  /** Inches. */
  margins: { top: number; bottom: number; left: number; right: number };
  gridlines: boolean;
}

/** Excel's defaults for a sheet with no print settings. */
export const DEFAULT_PRINT: PrintSettings = { paper: "letter", landscape: false, fit: "normal", scale: 100, margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 }, gridlines: false };

/** Paper sizes in inches, by Excel's paperSize number. */
const PAPERS: Record<Paper, { code: number; w: number; h: number }> = { letter: { code: 1, w: 8.5, h: 11 }, legal: { code: 5, w: 8.5, h: 14 }, a4: { code: 9, w: 8.27, h: 11.69 } };

const el = (sheet: Sheet, tag: string) => sheet.first(sheet.doc.documentElement, tag);
const num = (e: Element | null, attr: string) => (e?.hasAttribute(attr) ? Number(e.getAttribute(attr)) : null);

export function printSettings(sheet: Sheet): PrintSettings {
  const setup = el(sheet, "pageSetup"), options = el(sheet, "printOptions"), margins = el(sheet, "pageMargins");
  const fitToPage = sheet.first(el(sheet, "sheetPr") ?? sheet.doc.documentElement, "pageSetUpPr")?.getAttribute("fitToPage");
  const paper = (Object.keys(PAPERS) as Paper[]).find((p) => PAPERS[p].code === (num(setup, "paperSize") ?? 1)) ?? "letter";
  const fit = !["1", "true"].includes(fitToPage ?? "") ? "normal" : (num(setup, "fitToHeight") ?? 1) === 0 ? "width" : "page";
  return {
    paper,
    landscape: setup?.getAttribute("orientation") === "landscape",
    fit,
    scale: num(setup, "scale") ?? 100,
    margins: {
      top: num(margins, "top") ?? DEFAULT_PRINT.margins.top,
      bottom: num(margins, "bottom") ?? DEFAULT_PRINT.margins.bottom,
      left: num(margins, "left") ?? DEFAULT_PRINT.margins.left,
      right: num(margins, "right") ?? DEFAULT_PRINT.margins.right,
    },
    gridlines: ["1", "true"].includes(options?.getAttribute("gridLines") ?? ""),
  };
}

/** Save print settings into the sheet (one undo step). */
export function setPrintSettings(e: Edits, sheet: Sheet, ps: PrintSettings): void {
  const path = sheet.ref.path;
  e.structural(() => {
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    const options = child(s, "printOptions", true)!;
    if (ps.gridlines) options.setAttribute("gridLines", "1");
    else options.removeAttribute("gridLines");
    if (!options.attributes.length) options.remove();
    const margins = child(s, "pageMargins", true)!;
    for (const side of ["left", "right", "top", "bottom"] as const) margins.setAttribute(side, String(ps.margins[side]));
    if (!margins.hasAttribute("header")) margins.setAttribute("header", "0.3");
    if (!margins.hasAttribute("footer")) margins.setAttribute("footer", "0.3");
    const setup = child(s, "pageSetup", true)!;
    setup.setAttribute("paperSize", String(PAPERS[ps.paper].code));
    if (ps.fit === "normal" && ps.scale !== 100) setup.setAttribute("scale", String(ps.scale));
    else setup.removeAttribute("scale");
    if (ps.fit === "normal") {
      setup.removeAttribute("fitToWidth");
      setup.removeAttribute("fitToHeight");
    } else {
      setup.setAttribute("fitToWidth", "1");
      setup.setAttribute("fitToHeight", ps.fit === "width" ? "0" : "1");
    }
    setup.setAttribute("orientation", ps.landscape ? "landscape" : "portrait");
    // Fitting is switched on in sheetPr (its pageSetUpPr goes after tabColor and outlinePr).
    const pr = child(s, "sheetPr", ps.fit !== "normal");
    if (pr) {
      let up = s.first(pr, "pageSetUpPr");
      if (ps.fit !== "normal") {
        if (!up) up = pr.appendChild(s.doc.createElementNS(S, "pageSetUpPr"));
        up.setAttribute("fitToPage", "1");
      } else up?.removeAttribute("fitToPage");
    }
    s.touch();
  });
}

export interface PrintPage {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Pages for a range of columns / rows (their sizes in px): the scale used, and each page's columns and rows. */
export function pagePlan(cols: number[], rows: number[], ps: PrintSettings): { scale: number; width: number; height: number; pages: PrintPage[] } {
  const paper = PAPERS[ps.paper];
  const width = (ps.landscape ? paper.h : paper.w) * 96, height = (ps.landscape ? paper.w : paper.h) * 96;
  const roomW = width - (ps.margins.left + ps.margins.right) * 96, roomH = height - (ps.margins.top + ps.margins.bottom) * 96;
  const totalW = cols.reduce((a, b) => a + b, 0), totalH = rows.reduce((a, b) => a + b, 0);
  const scale = ps.fit === "width" ? Math.min(1, roomW / (totalW || 1)) : ps.fit === "page" ? Math.min(1, roomW / (totalW || 1), roomH / (totalH || 1)) : ps.scale / 100;
  /** Runs of sizes that fit `room` at the scale (at least one each). */
  const runs = (sizes: number[], room: number) => {
    const out: [number, number][] = [];
    let start = 0;
    while (start < sizes.length) {
      let end = start, used = 0;
      while (end < sizes.length && (used + sizes[end]) * scale <= room + 0.01) used += sizes[end++];
      if (end === start) end++;
      out.push([start, end - 1]);
      start = end;
    }
    return out;
  };
  const across = ps.fit === "normal" ? runs(cols, roomW) : [[0, cols.length - 1] as [number, number]];
  const down = ps.fit === "page" ? [[0, rows.length - 1] as [number, number]] : runs(rows, roomH);
  const pages: PrintPage[] = [];
  for (const [c1, c2] of across) for (const [r1, r2] of down) pages.push({ r1, c1, r2, c2 });
  return { scale, width, height, pages };
}
