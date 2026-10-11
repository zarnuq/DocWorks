/*
 * Pictures on a sheet: xdr:pic elements in the sheet's drawing part, placed by their anchor (one-cell, two-cell or
 * absolute) in sheet pixels. Charts in the same part are drawn by charts.ts; shapes are kept but not drawn.
 */
import { Sheet, Workbook } from "./workbook-model";

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_PX = 9525;

export interface Picture {
  name: string;
  alt: string;
  /** Zip path of the image. */
  image: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

const child = (el: Element, tag: string) => Array.from(el.children).find((c) => c.namespaceURI === XDR && c.localName === tag) ?? null;
const num = (el: Element | null, tag: string) => Number(el ? child(el, tag)?.textContent ?? 0 : 0);

/** Sheet pixel position of a from / to marker. */
function point(sheet: Sheet, m: Element | null) {
  return { x: m ? sheet.colX(num(m, "col")) + num(m, "colOff") / EMU_PER_PX : 0, y: m ? sheet.rowY(num(m, "row")) + num(m, "rowOff") / EMU_PER_PX : 0 };
}

/** Where an anchor (two-cell, one-cell or absolute) puts its object, in sheet pixels. */
export function anchorBox(sheet: Sheet, anchor: Element, content: Element | null = null): { x: number; y: number; w: number; h: number } | null {
  // A one-cell or absolute anchor sizes its object by xdr:ext, or by the object's own a:ext.
  const ext = child(anchor, "ext") ?? content?.getElementsByTagNameNS(A, "ext")[0] ?? null;
  const size = { w: Number(ext?.getAttribute("cx") ?? 0) / EMU_PER_PX, h: Number(ext?.getAttribute("cy") ?? 0) / EMU_PER_PX };
  if (anchor.localName === "twoCellAnchor") {
    const a = point(sheet, child(anchor, "from")), b = point(sheet, child(anchor, "to"));
    return { x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y };
  }
  if (anchor.localName === "oneCellAnchor") return { ...point(sheet, child(anchor, "from")), ...size };
  if (anchor.localName === "absoluteAnchor") {
    const pos = child(anchor, "pos");
    return { x: Number(pos?.getAttribute("x") ?? 0) / EMU_PER_PX, y: Number(pos?.getAttribute("y") ?? 0) / EMU_PER_PX, ...size };
  }
  return null;
}

export function sheetPictures(wb: Workbook, sheet: Sheet): Picture[] {
  const rel = wb.pkg.rels(sheet.ref.path).find((r) => r.type === "drawing" && !r.external);
  const doc = rel && wb.pkg.xml(rel.target);
  if (!rel || !doc) return [];
  const rels = wb.pkg.rels(rel.target);
  const out: Picture[] = [];
  for (const anchor of Array.from(doc.documentElement.children)) {
    const pic = child(anchor, "pic");
    const embed = pic?.getElementsByTagNameNS(A, "blip")[0]?.getAttributeNS(R, "embed");
    const image = embed ? rels.find((r) => r.id === embed && !r.external)?.target : undefined;
    if (!pic || !image) continue;
    const box = anchorBox(sheet, anchor, child(anchor, "ext") ? null : pic);
    if (!box) continue;
    let { x, y, w, h } = box;
    if (anchor.localName === "oneCellAnchor") {
      // Google Sheets exports an image in a cell this way: at the cell's corner, at the image's own (smaller) size,
      // while Sheets shows it fitted to the cell. Draw it fitted and centered, keeping its proportions.
      const from = child(anchor, "from")!;
      const row = num(from, "row"), col = num(from, "col");
      if (!num(from, "rowOff") && !num(from, "colOff") && w > 0 && h > 0) {
        const cellBox = sheet.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
        const bw = sheet.colX(cellBox.c2 + 1) - sheet.colX(cellBox.c1), bh = sheet.rowY(cellBox.r2 + 1) - sheet.rowY(cellBox.r1);
        if (w <= bw + 1 && h <= bh + 1) {
          const k = Math.min(bw / w, bh / h);
          [w, h] = [w * k, h * k];
          [x, y] = [box.x + (bw - w) / 2, box.y + (bh - h) / 2];
        }
      }
    }
    const nv = pic.getElementsByTagNameNS(XDR, "cNvPr")[0];
    out.push({ name: nv?.getAttribute("name") ?? "", alt: nv?.getAttribute("descr") ?? nv?.getAttribute("title") ?? "", image, x, y, w, h });
  }
  return out;
}
