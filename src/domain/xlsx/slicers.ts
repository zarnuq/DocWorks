/*
 * Table slicers (Excel 2013+; Google Sheets' Data › Add a slicer): a box of a table column's values whose choice
 * filters the table. Excel keeps a slicer in five places: a slicer cache part (workbook extLst x15:slicerCaches) that
 * names the table and column, a defined name Slicer_<name> = #N/A, a slicer part (sheet extLst x14:slicerList) with
 * its caption, a drawing anchor (mc:AlternateContent, sle:slicer inside the sle15 choice) for where it sits, and the
 * choice itself as the table's autoFilter. Pivot-table slicers are left as they are.
 * Written from ECMA-376 / [MS-XLSX]; not checked against a slicer Excel made.
 */
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { escapeXml } from "../../core/utils/xml";
import type { Edits } from "./commands";
import { anchorBox, marker, placeAnchor } from "./drawing";
import { assertAction } from "./protection";
import { parseRange, quoteSheet, Rect, rangeName } from "./references";
import { child as sheetChild, setFilter } from "./sheet-edit";
import type { Sheet, Workbook } from "./workbook-model";

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer";
const SLE15 = "http://schemas.microsoft.com/office/drawing/2012/slicer";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.microsoft.com/office/2007/relationships/";
/** extLst entries: the workbook's table slicer caches, a sheet's table slicers, a cache's table and column. */
const WORKBOOK_EXT = "{46BE6895-7355-4a93-B00E-2C351335B9C9}";
const SHEET_EXT = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";
const CACHE_EXT = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";
const NA = "#N/A";
const TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

export interface SheetTable {
  path: string;
  id: number;
  name: string;
  ref: Rect;
  /** tableColumn id and name, in order (the first is the table's first column). */
  columns: { id: number; name: string }[];
}

export interface SheetSlicer {
  name: string;
  caption: string;
  table: SheetTable;
  /** Index of its column in the table. */
  column: number;
  /** Where it's drawn (sheet pixels), and its anchor in the drawing part; null when it has no drawing. */
  box: { x: number; y: number; w: number; h: number } | null;
  drawing: string | null;
}

const kids = (el: Element | null | undefined, ns: string, tag: string) => Array.from(el?.children ?? []).filter((c) => c.namespaceURI === ns && c.localName === tag);

export function sheetTables(wb: Workbook, sheet: Sheet): SheetTable[] {
  const out: SheetTable[] = [];
  for (const rel of wb.pkg.rels(sheet.ref.path)) {
    if (rel.type !== "table" || rel.external) continue;
    const t = wb.pkg.xml(rel.target)?.documentElement;
    const ref = parseRange(t?.getAttribute("ref") ?? "");
    if (!t || !ref) continue;
    const cols = kids(kids(t, NS_S, "tableColumns")[0], NS_S, "tableColumn").map((c) => ({ id: Number(c.getAttribute("id")), name: c.getAttribute("name") ?? "" }));
    out.push({ path: rel.target, id: Number(t.getAttribute("id")), name: t.getAttribute("displayName") ?? t.getAttribute("name") ?? "", ref, columns: cols });
  }
  return out;
}

export function tableAt(wb: Workbook, sheet: Sheet, row: number, col: number): SheetTable | null {
  return sheetTables(wb, sheet).find((t) => row >= t.ref.r1 && row <= t.ref.r2 && col >= t.ref.c1 && col <= t.ref.c2) ?? null;
}

/** Table slicer caches by name: the table id and tableColumn id they filter. */
function caches(wb: Workbook): Map<string, { path: string; rId: string; tableId: number; column: number }> {
  const out = new Map<string, { path: string; rId: string; tableId: number; column: number }>();
  for (const rel of wb.pkg.rels(wb.pkg.workbookPath)) {
    if (rel.type !== "slicerCache" || rel.external) continue;
    const doc = wb.pkg.xml(rel.target);
    const t = doc?.getElementsByTagNameNS(X15, "tableSlicerCache")[0];
    const name = doc?.documentElement.getAttribute("name");
    if (t && name) out.set(name, { path: rel.target, rId: rel.id, tableId: Number(t.getAttribute("tableId")), column: Number(t.getAttribute("column")) });
  }
  return out;
}

/** Drawing anchors of slicers by name (the sle:slicer in an anchor's AlternateContent). */
function slicerAnchors(wb: Workbook, sheet: Sheet): { drawing: string; byName: Map<string, Element> } | null {
  const rel = wb.pkg.rels(sheet.ref.path).find((r) => r.type === "drawing" && !r.external);
  const doc = rel && wb.pkg.xml(rel.target);
  if (!rel || !doc) return null;
  const byName = new Map<string, Element>();
  for (const anchor of Array.from(doc.documentElement.children)) {
    for (const s of Array.from(anchor.getElementsByTagNameNS(SLE, "slicer"))) byName.set(s.getAttribute("name") ?? "", anchor);
  }
  return { drawing: rel.target, byName };
}

export function sheetSlicers(wb: Workbook, sheet: Sheet): SheetSlicer[] {
  const tables = new Map<number, SheetTable>();
  for (const s of wb.sheets) for (const t of sheetTables(wb, s)) tables.set(t.id, t);
  const own = new Set(sheetTables(wb, sheet).map((t) => t.id));
  const cacheByName = caches(wb);
  const anchors = slicerAnchors(wb, sheet);
  const out: SheetSlicer[] = [];
  for (const rel of wb.pkg.rels(sheet.ref.path)) {
    if (rel.type !== "slicer" || rel.external) continue;
    for (const el of Array.from(wb.pkg.xml(rel.target)?.getElementsByTagNameNS(X14, "slicer") ?? [])) {
      const cache = cacheByName.get(el.getAttribute("cache") ?? "");
      const table = cache && tables.get(cache.tableId);
      const column = table ? table.columns.findIndex((c) => c.id === cache.column) : -1;
      // Only slicers of this sheet's tables can filter it.
      if (!table || column < 0 || !own.has(table.id)) continue;
      const name = el.getAttribute("name") ?? "";
      const anchor = anchors?.byName.get(name) ?? null;
      out.push({ name, caption: el.getAttribute("caption") ?? table.columns[column].name, table, column, box: anchor ? anchorBox(sheet, anchor) : null, drawing: anchor ? anchors!.drawing : null });
    }
  }
  return out;
}

// -----------------------------------------------------------------------------
// The choice: the table's autoFilter
// -----------------------------------------------------------------------------

/** The table column's distinct shown values (sorted, blank last as "") and whether each is chosen. */
export function slicerItems(wb: Workbook, sheet: Sheet, s: SheetSlicer): { value: string; selected: boolean }[] {
  const col = s.table.ref.c1 + s.column;
  const values = new Set<string>();
  for (let row = s.table.ref.r1 + 1; row <= s.table.ref.r2; row++) values.add(sheet.display(row, col).text);
  const allowed = tableFilter(wb, s.table, s.column);
  return [...values].sort((a, b) => (a === "" ? 1 : b === "" ? -1 : a.localeCompare(b, undefined, { numeric: true }))).map((value) => ({ value, selected: !allowed || allowed.has(value) }));
}

/** The values a table column's filter lets through, or null when it isn't filtered. */
export function tableFilter(wb: Workbook, table: SheetTable, column: number): Set<string> | null {
  const af = kids(wb.pkg.xml(table.path)?.documentElement, NS_S, "autoFilter")[0];
  const fc = kids(af, NS_S, "filterColumn").find((x) => Number(x.getAttribute("colId")) === column);
  const filters = kids(fc, NS_S, "filters")[0];
  if (!filters) return null;
  const out = new Set(kids(filters, NS_S, "filter").map((x) => x.getAttribute("val") ?? ""));
  if (filters.getAttribute("blank") === "1") out.add("");
  return out;
}

/** Filter a table column to `allowed` (null: everything) and hide / show its rows by all its column filters. */
function setTableFilter(e: Edits, sheet: Sheet, table: SheetTable, column: number, allowed: Set<string> | null): void {
  const doc = e.wb.pkg.xml(table.path)!;
  const root = doc.documentElement;
  let af = kids(root, NS_S, "autoFilter")[0];
  if (!af) {
    af = doc.createElementNS(NS_S, "autoFilter");
    af.setAttribute("ref", rangeName(table.ref));
    root.insertBefore(af, kids(root, NS_S, "sortState")[0] ?? kids(root, NS_S, "tableColumns")[0] ?? null);
  }
  for (const fc of kids(af, NS_S, "filterColumn")) if (Number(fc.getAttribute("colId")) === column) fc.remove();
  if (allowed) {
    const fc = doc.createElementNS(NS_S, "filterColumn");
    fc.setAttribute("colId", String(column));
    const filters = fc.appendChild(doc.createElementNS(NS_S, "filters"));
    if (allowed.has("")) filters.setAttribute("blank", "1");
    for (const v of [...allowed].filter(Boolean).sort()) filters.appendChild(doc.createElementNS(NS_S, "filter")).setAttribute("val", v);
    af.insertBefore(fc, kids(af, NS_S, "filterColumn").find((x) => Number(x.getAttribute("colId")) > column) ?? kids(af, NS_S, "sortState")[0] ?? null);
  }
  e.wb.pkg.touch(table.path);
  const live: SheetTable = { ...table };
  const filters = table.columns.map((_, i) => tableFilter(e.wb, live, i));
  for (let row = table.ref.r1 + 1; row <= table.ref.r2; row++) {
    const show = filters.every((f, i) => !f || f.has(sheet.display(row, table.ref.c1 + i).text));
    const el = sheet.rowEl(row, !show);
    if (!el) continue;
    if (show) el.removeAttribute("hidden");
    else el.setAttribute("hidden", "1");
  }
  sheet.touch();
}

/** Clicking slicer items: show only `allowed` (null clears the slicer's filter). */
export function setSlicerSelection(e: Edits, sheet: Sheet, s: SheetSlicer, allowed: Set<string> | null): void {
  assertAction(sheet, "autoFilter");
  const path = sheet.ref.path;
  e.structural(() => setTableFilter(e, e.wb.sheets.find((x) => x.ref.path === path)!, s.table, s.column, allowed));
}

// -----------------------------------------------------------------------------
// Adding, moving and deleting
// -----------------------------------------------------------------------------

/** The workbook's (or sheet's) extLst entry `uri`, made (with its list element) when missing. */
function extList(doc: Document, uri: string, prefix: string, ns: string, listTag: string): Element {
  const root = doc.documentElement;
  let extLst = kids(root, NS_S, "extLst")[0];
  if (!extLst) extLst = root.appendChild(doc.createElementNS(NS_S, "extLst"));
  let ext = kids(extLst, NS_S, "ext").find((x) => x.getAttribute("uri") === uri);
  if (!ext) {
    ext = extLst.appendChild(doc.createElementNS(NS_S, "ext"));
    ext.setAttribute("uri", uri);
    ext.setAttributeNS("http://www.w3.org/2000/xmlns/", `xmlns:${prefix}`, ns);
  }
  let list = Array.from(ext.children).find((c) => c.localName === listTag);
  if (!list) list = ext.appendChild(doc.createElementNS(listTag === "slicerList" ? X14 : X15, `${listTag === "slicerList" ? "x14" : "x15"}:${listTag}`));
  return list;
}

/** Drop an extLst entry pointing at `rId`, and the entry / list / extLst left empty. */
function dropExtEntry(doc: Document, uri: string, rId: string): void {
  const extLst = kids(doc.documentElement, NS_S, "extLst")[0];
  const ext = kids(extLst, NS_S, "ext").find((x) => x.getAttribute("uri") === uri);
  for (const el of Array.from(ext?.getElementsByTagNameNS(X14, "*") ?? [])) if (el.getAttributeNS(R, "id") === rId) el.remove();
  const list = ext?.firstElementChild;
  if (ext && list && !list.children.length) ext.remove();
  if (extLst && !extLst.children.length) extLst.remove();
}

const definedNamesEl = (doc: Document, create: boolean) => {
  let names = doc.getElementsByTagNameNS(NS_S, "definedNames")[0];
  if (!names && create) {
    names = doc.createElementNS(NS_S, "definedNames");
    const sheets = doc.getElementsByTagNameNS(NS_S, "sheets")[0];
    sheets.parentNode!.insertBefore(names, sheets.nextSibling);
  }
  return names ?? null;
};

/** Data › Add a slicer: a slicer for column `column` of `table`, drawn at `at` (sheet pixels). Returns its name. */
export function addSlicer(e: Edits, sheet: Sheet, table: SheetTable, column: number, at: { x: number; y: number }): string {
  assertAction(sheet, "objects");
  const caption = table.columns[column].name;
  const takenNames = new Set<string>();
  for (const s of e.wb.sheets) for (const rel of e.wb.pkg.rels(s.ref.path)) {
    if (rel.type === "slicer") for (const el of Array.from(e.wb.pkg.xml(rel.target)?.getElementsByTagNameNS(X14, "slicer") ?? [])) takenNames.add(el.getAttribute("name") ?? "");
  }
  let name = caption;
  for (let n = 1; takenNames.has(name); n++) name = `${caption} ${n}`;
  const definedNames = new Set(e.wb.names().map((n) => n.name.toLowerCase()));
  const base = `Slicer_${caption.replace(/[^A-Za-z0-9_]/g, "_")}`;
  let cacheName = base;
  for (let n = 1; definedNames.has(cacheName.toLowerCase()) || caches(e.wb).has(cacheName); n++) cacheName = `${base}${n}`;
  const path = sheet.ref.path;
  e.structural(() => {
    const pkg = e.wb.pkg;
    const book = pkg.workbook;
    // The cache, from the workbook.
    const cachePath = pkg.uniquePath("xl/slicerCaches/slicerCache", ".xml");
    pkg.putXml(cachePath, parseXml(
      `<slicerCacheDefinition xmlns="${X14}" xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${NS_S}" name="${escapeXml(cacheName)}" sourceName="${escapeXml(caption)}">` +
      `<extLst><x:ext uri="${CACHE_EXT}" xmlns:x15="${X15}"><x15:tableSlicerCache tableId="${table.id}" column="${table.columns[column].id}"/></x:ext></extLst></slicerCacheDefinition>`
    )!, "application/vnd.ms-excel.slicerCache+xml");
    const cacheRel = pkg.relate(pkg.workbookPath, `${REL}slicerCache`, cachePath);
    const cacheList = extList(book, WORKBOOK_EXT, "x15", X15, "slicerCaches");
    cacheList.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:x14", X14);
    cacheList.appendChild(book.createElementNS(X14, "x14:slicerCache")).setAttributeNS(R, "r:id", cacheRel);
    const dn = definedNamesEl(book, true).appendChild(book.createElementNS(NS_S, "definedName"));
    dn.setAttribute("name", cacheName);
    // Excel's placeholder formula for a slicer cache name.
    dn.textContent = NA;
    pkg.touch(pkg.workbookPath);
    // The slicer, from the sheet.
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    let slicerPath = pkg.rels(path).find((r) => r.type === "slicer" && !r.external)?.target;
    if (!slicerPath) {
      slicerPath = pkg.uniquePath("xl/slicers/slicer", ".xml");
      pkg.putXml(slicerPath, parseXml(`<slicers xmlns="${X14}" xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${NS_S}"/>`)!, "application/vnd.ms-excel.slicer+xml");
      const rel = pkg.relate(path, `${REL}slicer`, slicerPath);
      extList(s.doc, SHEET_EXT, "x15", X15, "slicerList").appendChild(s.doc.createElementNS(X14, "x14:slicer")).setAttributeNS(R, "r:id", rel);
    }
    const slicers = pkg.xml(slicerPath)!;
    const el = slicers.documentElement.appendChild(slicers.createElementNS(X14, "slicer"));
    el.setAttribute("name", name);
    el.setAttribute("cache", cacheName);
    el.setAttribute("caption", caption);
    el.setAttribute("rowHeight", "241300");
    pkg.touch(slicerPath);
    // Where it's drawn.
    const drawing = pkg.ensureDrawing(path);
    if (!sheetChild(s, "drawing")) sheetChild(s, "drawing", true)!.setAttributeNS(R, "r:id", drawing.rId);
    const doc = pkg.xml(drawing.path)!;
    const id = Math.max(0, ...Array.from(doc.getElementsByTagNameNS(XDR, "cNvPr")).map((c) => Number(c.getAttribute("id")) || 0)) + 1;
    const w = 180, h = 240;
    const anchor =
      `<xdr:twoCellAnchor xmlns:xdr="${XDR}" xmlns:a="${A}" editAs="oneCell">${marker(s, "from", at.x, at.y)}${marker(s, "to", at.x + w, at.y + h)}` +
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:sle15="${SLE15}" Requires="sle15">` +
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${SLE}"><sle:slicer xmlns:sle="${SLE}" name="${escapeXml(name)}"/></a:graphicData></a:graphic></xdr:graphicFrame>` +
      `</mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * 9525}" cy="${h * 9525}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
      `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>This shape represents a table slicer. Table slicers are supported in Excel 2013 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp>` +
      `</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
    doc.documentElement.appendChild(doc.importNode(parseXml(anchor)!.documentElement, true));
    pkg.touch(drawing.path);
    s.touch();
  });
  return name;
}

/** Move / resize a slicer's box. */
export function moveSlicer(e: Edits, sheet: Sheet, name: string, box: { x: number; y: number; w: number; h: number }): void {
  assertAction(sheet, "objects");
  const path = sheet.ref.path;
  e.structural(() => {
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    const anchors = slicerAnchors(e.wb, s);
    const anchor = anchors?.byName.get(name);
    if (!anchors || !anchor) return;
    placeAnchor(s, anchor, box);
    e.wb.pkg.touch(anchors.drawing);
  });
}

/** Delete a slicer: its anchor and entry, its cache and name when nothing else uses them, and its filter (as Excel). */
export function deleteSlicer(e: Edits, sheet: Sheet, slicer: SheetSlicer): void {
  assertAction(sheet, "objects");
  const path = sheet.ref.path;
  e.structural(() => {
    const pkg = e.wb.pkg;
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    const anchors = slicerAnchors(e.wb, s);
    const anchor = anchors?.byName.get(slicer.name);
    if (anchor) {
      anchor.remove();
      pkg.touch(anchors!.drawing);
    }
    let cacheName = "";
    for (const rel of pkg.rels(path)) {
      if (rel.type !== "slicer" || rel.external) continue;
      const doc = pkg.xml(rel.target)!;
      const el = Array.from(doc.getElementsByTagNameNS(X14, "slicer")).find((x) => x.getAttribute("name") === slicer.name);
      if (!el) continue;
      cacheName = el.getAttribute("cache") ?? "";
      el.remove();
      pkg.touch(rel.target);
      if (!doc.documentElement.children.length) {
        dropExtEntry(s.doc, SHEET_EXT, rel.id);
        pkg.removeXml(rel.target);
        pkg.dropRel(path, rel.id);
      }
    }
    const stillUsed = e.wb.sheets.some((x) => pkg.rels(x.ref.path).some((r) => r.type === "slicer" && Array.from(pkg.xml(r.target)?.getElementsByTagNameNS(X14, "slicer") ?? []).some((el) => el.getAttribute("cache") === cacheName)));
    const cache = caches(e.wb).get(cacheName);
    if (cache && !stillUsed) {
      dropExtEntry(pkg.workbook, WORKBOOK_EXT, cache.rId);
      pkg.removeXml(cache.path);
      pkg.dropRel(pkg.workbookPath, cache.rId);
      const names = definedNamesEl(pkg.workbook, false);
      for (const n of Array.from(names?.children ?? [])) if (n.getAttribute("name") === cacheName) n.remove();
      if (names && !names.children.length) names.remove();
      pkg.touch(pkg.workbookPath);
    }
    setTableFilter(e, s, slicer.table, slicer.column, null);
  });
}

/** Excel's current region: the block of filled cells around (row, col), grown while a neighbouring row or column has any. */
export function regionAround(sheet: Sheet, row: number, col: number): Rect {
  const r = { r1: row, c1: col, r2: row, c2: col };
  const filled = (r1: number, c1: number, r2: number, c2: number) => {
    if (r1 < 0 || c1 < 0) return false;
    for (const cell of sheet.cellsIn({ r1, c1, r2, c2 })) if (sheet.display(cell.row, cell.col).text !== "") return true;
    return false;
  };
  for (let grew = true; grew; ) {
    grew = false;
    const top = Math.max(0, r.r1 - 1), left = Math.max(0, r.c1 - 1);
    if (r.r1 > 0 && filled(r.r1 - 1, left, r.r1 - 1, r.c2 + 1)) {
      r.r1--;
      grew = true;
    }
    if (filled(r.r2 + 1, left, r.r2 + 1, r.c2 + 1)) {
      r.r2++;
      grew = true;
    }
    if (r.c1 > 0 && filled(top, r.c1 - 1, r.r2 + 1, r.c1 - 1)) {
      r.c1--;
      grew = true;
    }
    if (filled(top, r.c2 + 1, r.r2 + 1, r.c2 + 1)) {
      r.c2++;
      grew = true;
    }
  }
  return r;
}

/**
 * Turn `r` (header row first) into a table, as Excel's Format as Table, so it can have slicers. Each column needs a
 * heading, all different; a sheet filter on the range goes (a table has its own).
 */
export function makeTable(e: Edits, sheet: Sheet, r: Rect): SheetTable {
  assertAction(sheet, "sheet");
  if (r.r2 <= r.r1) throw new Error("A table needs a header row and at least one row of data.");
  const names: string[] = [];
  for (let col = r.c1; col <= r.c2; col++) {
    const text = sheet.display(r.r1, col).text.trim();
    if (!text) throw new Error("Each column needs a heading in the first row.");
    if (names.some((n) => n.toLowerCase() === text.toLowerCase())) throw new Error(`Two columns are called "${text}".`);
    names.push(text);
  }
  for (const s of e.wb.sheets) for (const t of sheetTables(e.wb, s)) {
    if (s === sheet && !(t.ref.r2 < r.r1 || t.ref.r1 > r.r2 || t.ref.c2 < r.c1 || t.ref.c1 > r.c2)) throw new Error(`This range overlaps ${t.name}.`);
  }
  if (sheet.filter && !(sheet.filter.r2 < r.r1 || sheet.filter.r1 > r.r2 || sheet.filter.c2 < r.c1 || sheet.filter.c1 > r.c2)) setFilter(e, sheet, null);
  const ids = e.wb.sheets.flatMap((s) => sheetTables(e.wb, s).map((t) => t.id));
  const id = Math.max(0, ...ids) + 1;
  const taken = new Set(e.wb.names().map((n) => n.name.toLowerCase()));
  for (const s of e.wb.sheets) for (const t of sheetTables(e.wb, s)) taken.add(t.name.toLowerCase());
  let n = id;
  while (taken.has(`table${n}`)) n++;
  const name = `Table${n}`;
  const ref = rangeName(r);
  const path = sheet.ref.path;
  let tablePath = "";
  e.structural(() => {
    const pkg = e.wb.pkg;
    tablePath = pkg.uniquePath("xl/tables/table", ".xml");
    pkg.putXml(tablePath, parseXml(
      `<table xmlns="${NS_S}" id="${id}" name="${name}" displayName="${name}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/>` +
      `<tableColumns count="${names.length}">${names.map((c, i) => `<tableColumn id="${i + 1}" name="${escapeXml(c)}"/>`).join("")}</tableColumns>` +
      `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`
    )!, TABLE_CT);
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    const rId = pkg.relate(path, "table", tablePath);
    const parts = sheetChild(s, "tableParts", true)!;
    parts.appendChild(s.doc.createElementNS(NS_S, "tablePart")).setAttributeNS(R, "r:id", rId);
    parts.setAttribute("count", String(parts.children.length));
    s.touch();
  });
  return { path: tablePath, id, name, ref: r, columns: names.map((c, i) => ({ id: i + 1, name: c })) };
}

/** "Table1" with its sheet, for messages. */
export const tableLabel = (sheet: Sheet, t: SheetTable) => `${t.name} (${quoteSheet(sheet.name)}!${rangeName(t.ref)})`;
