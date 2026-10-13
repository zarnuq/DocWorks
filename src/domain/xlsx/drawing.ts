/*
 * Pictures on a sheet: xdr:pic elements in the sheet's drawing part, placed by their anchor (one-cell, two-cell or
 * absolute) in sheet pixels. Charts in the same part are drawn by charts.ts; shapes are kept but not drawn.
 */
import { Sheet, Workbook } from "./workbook-model";
import type { Edits } from "./commands";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { assertAction } from "./protection";
import { child as sheetChild } from "./sheet-edit";

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_PX = 9525;

export interface Picture {
  /** Identifies the picture for moving / deleting: its drawing part and its anchor's index (Google exports every cNvPr id as 0). */
  key: string;
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
  for (const [index, anchor] of Array.from(doc.documentElement.children).entries()) {
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
    out.push({ key: `${rel.target}#${index}`, name: nv?.getAttribute("name") ?? "", alt: nv?.getAttribute("descr") ?? nv?.getAttribute("title") ?? "", image, x, y, w, h });
  }
  return out;
}

/** A drawing anchor marker (xdr:from / xdr:to) at sheet pixel (x, y). */
export function marker(sheet: Sheet, tag: string, x: number, y: number): string {
  const col = sheet.colAt(x), row = sheet.rowAt(y);
  const off = (v: number) => Math.max(0, Math.round(v * EMU_PER_PX));
  return `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${off(x - sheet.colX(col))}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${off(y - sheet.rowY(row))}</xdr:rowOff></xdr:${tag}>`;
}

/**
 * Google Sheets' Insert › Image › Over cells: a picture at a cell's top-left corner, `size` in sheet pixels, moving
 * with its cells but not resized by them (twoCellAnchor editAs="oneCell", as Excel inserts pictures).
 */
export function insertPicture(e: Edits, sheet: Sheet, bytes: Uint8Array, ext: string, at: { row: number; col: number }, size: { w: number; h: number }, name: string): void {
  assertAction(sheet, "objects");
  e.structural(() => {
    const pkg = e.wb.pkg;
    const drawing = pkg.ensureDrawing(sheet.ref.path);
    if (!sheetChild(sheet, "drawing")) {
      sheetChild(sheet, "drawing", true)!.setAttributeNS(R, "r:id", drawing.rId);
      pkg.touch(sheet.ref.path);
    }
    const image = pkg.addImagePart(drawing.path, bytes, ext);
    const doc = pkg.xml(drawing.path)!;
    const ids = Array.from(doc.getElementsByTagNameNS(XDR, "cNvPr")).map((c) => Number(c.getAttribute("id")) || 0);
    const id = Math.max(0, ...ids) + 1;
    const x = sheet.colX(at.col), y = sheet.rowY(at.row);
    const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    const anchor =
      `<xdr:twoCellAnchor xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}" editAs="oneCell">${marker(sheet, "from", x, y)}${marker(sheet, "to", x + size.w, y + size.h)}` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${esc(name)}" descr=""/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip r:embed="${image.rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(size.w * EMU_PER_PX)}" cy="${Math.round(size.h * EMU_PER_PX)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
    doc.documentElement.appendChild(doc.importNode(parseXml(anchor)!.documentElement, true));
    pkg.touch(drawing.path);
  });
}

/** Put a drawing anchor at `box` (sheet pixels): its from / to markers, or an absolute anchor's position and size. */
export function placeAnchor(sheet: Sheet, anchor: Element, box: { x: number; y: number; w: number; h: number }): void {
  const doc = anchor.ownerDocument;
  const x = Math.max(0, box.x), y = Math.max(0, box.y), w = Math.max(20, box.w), h = Math.max(20, box.h);
  const kid = (name: string) => Array.from(anchor.children).find((c) => c.namespaceURI === XDR && c.localName === name);
  const swap = (tag: string, px: number, py: number) => {
    const old = kid(tag);
    const next = doc.importNode(parseXml(marker(sheet, tag, px, py).replace(`<xdr:${tag}>`, `<xdr:${tag} xmlns:xdr="${XDR}">`))!.documentElement, true);
    if (old) old.replaceWith(next);
    return !!old;
  };
  const ext = (el: Element | undefined) => {
    el?.setAttribute("cx", String(Math.round(w * EMU_PER_PX)));
    el?.setAttribute("cy", String(Math.round(h * EMU_PER_PX)));
  };
  if (anchor.localName === "absoluteAnchor") {
    kid("pos")?.setAttribute("x", String(Math.round(x * EMU_PER_PX)));
    kid("pos")?.setAttribute("y", String(Math.round(y * EMU_PER_PX)));
    ext(kid("ext"));
  } else {
    swap("from", x, y);
    if (!swap("to", x + w, y + h)) ext(kid("ext"));
  }
  // A picture's own size goes with it.
  ext(anchor.getElementsByTagNameNS(A, "ext")[0]);
}

/** The anchor holding picture `key` (see Picture.key), in its live drawing part. */
function pictureAnchor(e: Edits, key: string): { drawing: string; anchor: Element } | null {
  const at = key.lastIndexOf("#");
  const drawing = key.slice(0, at);
  const anchor = e.wb.pkg.xml(drawing)?.documentElement.children[Number(key.slice(at + 1))];
  return anchor && child(anchor, "pic") ? { drawing, anchor } : null;
}

/** Move / resize a picture (dragged on the grid). */
export function movePicture(e: Edits, sheet: Sheet, key: string, box: { x: number; y: number; w: number; h: number }): void {
  assertAction(sheet, "objects");
  e.structural(() => {
    const found = pictureAnchor(e, key);
    if (!found) return;
    placeAnchor(e.wb.sheets.find((s) => s.ref.path === sheet.ref.path) ?? sheet, found.anchor, box);
    e.wb.pkg.touch(found.drawing);
  });
}

/** Delete a picture and, when nothing else shows its image, the drawing's relationship to it. */
export function deletePicture(e: Edits, sheet: Sheet, key: string): void {
  assertAction(sheet, "objects");
  e.structural(() => {
    const found = pictureAnchor(e, key);
    if (!found) return;
    const rId = found.anchor.getElementsByTagNameNS(A, "blip")[0]?.getAttributeNS(R, "embed");
    const doc = found.anchor.ownerDocument;
    found.anchor.remove();
    if (rId && !Array.from(doc.getElementsByTagNameNS(A, "blip")).some((b) => b.getAttributeNS(R, "embed") === rId)) e.wb.pkg.dropRel(found.drawing, rId);
    e.wb.pkg.touch(found.drawing);
  });
}
