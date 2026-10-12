import JSZip from "jszip";
import { OoxmlPackage, parseXml } from "../ooxml/ooxml-package";
import { escapeXml } from "../../core/utils/xml";

export const NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const STRINGS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const COMMENTS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const VML_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing";
/** Sheet children that come after w:legacyDrawing. */
const LEGACY_AFTER = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];

/** Column letters / row number of an A1 reference, 0-based. */
function cellOf(ref: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref) ?? ["", "A", "1"];
  return { row: Number(m[2]) - 1, col: [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1 };
}

/** The VML drawing Excel shows notes with: one hidden yellow box per note, beside its cell. */
function noteVml(refs: string[]): string {
  const shapes = refs.map((ref, i) => {
    const { row, col } = cellOf(ref);
    return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
      `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1}, 15, ${row}, 2, ${col + 3}, 15, ${row + 3}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
  });
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shapes.join("") + `</xml>`;
}
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface SheetRef {
  name: string;
  /** Stable id from workbook.xml; independent of tab order, name and relationship id. */
  sheetId: number;
  rId: string;
  path: string;
  state: "visible" | "hidden" | "veryHidden";
  /** The workbook.xml <sheet> element. */
  el: Element;
}

/**
 * An .xlsx package. Saves only the parts edits `touch`, so opening and saving a workbook never rewrites a worksheet
 * nobody changed (unknown parts, drawings and extensions keep their bytes).
 */
export class XlsxWorkbook extends OoxmlPackage {
  workbookPath = "xl/workbook.xml";
  stylesPath: string | null = null;
  stringsPath: string | null = null;

  static async load(data: ArrayBuffer | Uint8Array): Promise<XlsxWorkbook> {
    const zip = await JSZip.loadAsync(data);
    const wb = new XlsxWorkbook(zip);
    wb.dirtyOnly = true;
    await wb.readParts();
    wb.workbookPath = wb.relOfType("", "officeDocument") ?? wb.workbookPath;
    const root = wb.xml(wb.workbookPath)?.documentElement;
    if (!root || root.localName !== "workbook") throw new Error("Not a valid .xlsx file: the workbook part is missing.");
    if (root.namespaceURI !== NS_S) throw new Error("This workbook uses the Strict Open XML format, which isn't supported yet.");
    wb.stylesPath = wb.relOfType(wb.workbookPath, "styles");
    wb.stringsPath = wb.relOfType(wb.workbookPath, "sharedStrings");
    return wb;
  }

  get workbook(): Document {
    return this.xml(this.workbookPath)!;
  }

  sheets(): SheetRef[] {
    const rels = this.rels(this.workbookPath);
    return Array.from(this.workbook.getElementsByTagNameNS(NS_S, "sheet")).flatMap((el) => {
      const rId = el.getAttributeNS(NS_R, "id") ?? "";
      const path = rels.find((r) => r.id === rId)?.target;
      if (!path || !this.xml(path)) return [];
      const state = el.getAttribute("state");
      return [{ name: el.getAttribute("name") ?? "", sheetId: Number(el.getAttribute("sheetId")), rId, path, state: state === "hidden" || state === "veryHidden" ? state : "visible", el }];
    });
  }

  /** The shared string table, created on first use. */
  strings(): Document {
    if (!this.stringsPath) {
      this.stringsPath = this.uniquePath("xl/sharedStrings", ".xml");
      this.putXml(this.stringsPath, parseXml(`<sst xmlns="${NS_S}" count="0" uniqueCount="0"/>`)!, STRINGS_CT);
      this.addRel(this.workbookPath, "sharedStrings", this.stringsPath);
    }
    return this.xml(this.stringsPath)!;
  }

  /** Append a sheet part with `xml` named `name` after tab `after` (end by default). */
  addSheet(name: string, xml: string, after = this.sheets().length - 1): SheetRef {
    const path = this.uniquePath("xl/worksheets/sheet", ".xml");
    this.putXml(path, parseXml(xml)!, SHEET_CT);
    const rId = this.addRel(this.workbookPath, "worksheet", path);
    const sheetsEl = this.workbook.getElementsByTagNameNS(NS_S, "sheets")[0];
    const el = this.workbook.createElementNS(NS_S, "sheet");
    el.setAttribute("name", name);
    el.setAttribute("sheetId", String(Math.max(0, ...this.sheets().map((s) => s.sheetId)) + 1));
    el.setAttributeNS(NS_R, "r:id", rId);
    sheetsEl.insertBefore(el, this.sheets()[after + 1]?.el ?? null);
    this.touch(this.workbookPath);
    return this.sheets().find((s) => s.path === path)!;
  }

  /**
   * A new pivot table on `sheetPath` over `ref` on sheet `source`: an empty definition, cache and records (refreshing
   * the pivot table fills them in), linked from the sheet and listed in the workbook's pivotCaches.
   */
  addPivotTable(sheetPath: string, source: string, ref: string): string {
    const S = `xmlns="${NS_S}" xmlns:r="${NS_R}"`;
    const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.";
    const caches = this.workbook.getElementsByTagNameNS(NS_S, "pivotCache");
    const cacheId = Math.max(0, ...Array.from(caches).map((c) => Number(c.getAttribute("cacheId")) || 0)) + 1;
    const records = this.uniquePath("xl/pivotCache/pivotCacheRecords", ".xml");
    this.putXml(records, parseXml(`<pivotCacheRecords ${S} count="0"/>`)!, `${CT}pivotCacheRecords+xml`);
    const cache = this.uniquePath("xl/pivotCache/pivotCacheDefinition", ".xml");
    this.putXml(cache, parseXml(`<pivotCacheDefinition ${S} r:id="rId1" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="${ref}" sheet="${escapeXml(source)}"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>`)!, `${CT}pivotCacheDefinition+xml`);
    this.addRel(cache, "pivotCacheRecords", records);
    const table = this.uniquePath("xl/pivotTables/pivotTable", ".xml");
    this.putXml(table, parseXml(`<pivotTableDefinition ${S} name="PivotTable${cacheId}" cacheId="${cacheId}" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" useAutoFormatting="1" itemPrintTitles="1" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="0"/><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`)!, `${CT}pivotTable+xml`);
    this.addRel(table, "pivotCacheDefinition", cache);
    this.addRel(sheetPath, "pivotTable", table);
    // CT_Workbook: pivotCaches after calcPr / oleSize / customWorkbookViews, before smartTagPr and the rest.
    const root = this.workbook.documentElement;
    let list = Array.from(root.children).find((c) => c.localName === "pivotCaches");
    if (!list) {
      const after = ["smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
      list = root.insertBefore(this.workbook.createElementNS(NS_S, "pivotCaches"), Array.from(root.children).find((c) => after.includes(c.localName)) ?? null);
    }
    const entry = list.appendChild(this.workbook.createElementNS(NS_S, "pivotCache"));
    entry.setAttribute("cacheId", String(cacheId));
    entry.setAttributeNS(NS_R, "r:id", this.addRel(this.workbookPath, "pivotCacheDefinition", cache));
    this.touch(this.workbookPath);
    this.touch(sheetPath);
    return table;
  }

  /** Remove a sheet's part, relationship and tab (its own rels and parts only it uses stay unreferenced). */
  deleteSheet(sheet: SheetRef): void {
    sheet.el.remove();
    this.removeRel(this.workbookPath, sheet.rId);
    this.removePart(sheet.path);
    this.touch(this.workbookPath);
  }

  /** The sheet's drawing part and the sheet's relationship to it, created if missing. */
  /** Reuse or append Excel's dynamic-array metadata record; existing metadata indexes remain stable. */
  ensureDynamicMetadata(): number {
    const rel = this.rels(this.workbookPath).find(r => r.type === "sheetMetadata");
    const path = rel?.target ?? "xl/metadata.xml";
    let doc = this.xml(path);
    if (!doc) {
      doc = parseXml(`<metadata xmlns="${NS_S}"/>`)!;
      this.putXml(path, doc, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
      this.addRel(this.workbookPath, "sheetMetadata", path);
    }
    const root = doc.documentElement;
    const order = ["metadataTypes", "metadataStrings", "mdxMetadata", "futureMetadata", "cellMetadata", "valueMetadata", "extLst"];
    const child = (parent: Element, tag: string) => Array.from(parent.children).find(c => c.localName === tag);
    const add = (tag: string) => {
      const el = doc.createElementNS(NS_S,tag);
      root.insertBefore(el,Array.from(root.children).find(c => order.indexOf(c.localName) > order.indexOf(tag)) ?? null);
      return el;
    };
    const types = child(root,"metadataTypes") ?? add("metadataTypes");
    let type = Array.from(types.children).findIndex(t => t.getAttribute("name") === "XLDAPR");
    if (type < 0) {
      const el = doc.createElementNS(NS_S,"metadataType");
      for (const [k,v] of Object.entries({name:"XLDAPR",minSupportedVersion:"120000",copy:"1",pasteAll:"1",pasteValues:"1",merge:"1",splitFirst:"1",rowColShift:"1",clearFormats:"1",clearComments:"1",assign:"1",coerce:"1",cellMeta:"1"})) el.setAttribute(k,v);
      types.appendChild(el); types.setAttribute("count",String(types.children.length));type=types.children.length-1;
    }
    let future = Array.from(root.children).find(c => c.localName === "futureMetadata" && c.getAttribute("name") === "XLDAPR");
    if (!future) {future=add("futureMetadata");future.setAttribute("name","XLDAPR");}
    let record = Array.from(future.children).findIndex(b => Array.from(b.getElementsByTagNameNS("http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray","dynamicArrayProperties")).some(p => p.getAttribute("fDynamic") === "1"));
    if (record < 0) {
      const bk = parseXml(`<bk xmlns="${NS_S}"><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" fDynamic="1" fCollapsed="0"/></ext></extLst></bk>`)!;
      future.appendChild(doc.importNode(bk.documentElement,true));future.setAttribute("count",String(future.children.length));record=future.children.length-1;
    }
    const cells = child(root,"cellMetadata") ?? add("cellMetadata");
    let index = Array.from(cells.children).findIndex(b => Array.from(b.children).some(r => r.getAttribute("t") === String(type+1) && r.getAttribute("v") === String(record)));
    if (index < 0) {
      const bk=doc.createElementNS(NS_S,"bk"), rc=doc.createElementNS(NS_S,"rc");
      rc.setAttribute("t",String(type+1));rc.setAttribute("v",String(record));bk.appendChild(rc);cells.appendChild(bk);cells.setAttribute("count",String(cells.children.length));index=cells.children.length-1;
    }
    this.touch(path);return index+1;
  }

  /** Whether a cell's metadata marks it as a dynamic-array anchor. */
  isDynamicMetadata(index: number): boolean {
    if (!(index > 0)) return false;
    const rel = this.rels(this.workbookPath).find(r => r.type === "sheetMetadata");
    const root = rel && this.xml(rel.target)?.documentElement;
    if (!root) return false;
    const child = (parent: Element, tag: string) => Array.from(parent.children).find(c => c.localName === tag);
    const bk=child(root,"cellMetadata")?.children[index-1], types=child(root,"metadataTypes");
    return Array.from(bk?.children ?? []).some(r => {
      const name=types?.children[Number(r.getAttribute("t"))-1]?.getAttribute("name");
      if(name !== "XLDAPR") return false;
      const future=Array.from(root.children).find(c => c.localName === "futureMetadata" && c.getAttribute("name") === name);
      const record=future?.children[Number(r.getAttribute("v"))];
      return Array.from(record?.getElementsByTagNameNS("http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray","dynamicArrayProperties") ?? []).some(p => ["1","true"].includes(p.getAttribute("fDynamic") ?? ""));
    });
  }

  ensureDrawing(sheetPath: string): { path: string; rId: string } {
    const rel = this.rels(sheetPath).find((r) => r.type === "drawing" && !r.external);
    if (rel) return { path: rel.target, rId: rel.id };
    const path = this.uniquePath("xl/drawings/drawing", ".xml");
    this.putXml(path, parseXml(`<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}"/>`)!, DRAWING_CT);
    return { path, rId: this.addRel(sheetPath, "drawing", path) };
  }

  /** Worksheet comments and their workbook-wide author identities. */
  ensureThreadedComments(sheetPath: string): string {
    const existing = this.rels(sheetPath).find(r => r.type === "threadedComment" && !r.external);
    if (existing) return existing.target;
    const path = this.uniquePath("xl/threadedComments/threadedComment", ".xml");
    this.putXml(path, parseXml('<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>')!, "application/vnd.ms-excel.threadedcomments+xml");
    this.addRel(sheetPath, "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", path);
    return path;
  }
  ensurePersons(): string {
    const existing = this.rels(this.workbookPath).find(r => r.type === "person" && !r.external);
    if (existing) return existing.target;
    const path = this.uniquePath("xl/persons/person", ".xml");
    this.putXml(path, parseXml('<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>')!, "application/vnd.ms-excel.person+xml");
    this.addRel(this.workbookPath, "http://schemas.microsoft.com/office/2017/10/relationships/person", path);
    return path;
  }

  // --- Notes (Excel's legacy comments) --------------------------------------------

  /** The sheet's comments part, if it has one. */
  commentsPath(sheetPath: string): string | null {
    return this.rels(sheetPath).find((r) => r.type === "comments" && !r.external)?.target ?? null;
  }

  /**
   * The sheet's comments part, created if missing, with the VML drawing Excel shows notes with (rebuilt from the
   * comments when saving) and the sheet's w:legacyDrawing link to it.
   */
  ensureComments(sheetPath: string): Document {
    const existing = this.commentsPath(sheetPath);
    if (existing) return this.xml(existing)!;
    const path = this.uniquePath("xl/comments", ".xml");
    this.putXml(path, parseXml(`<comments xmlns="${NS_S}"><authors/><commentList/></comments>`)!, COMMENTS_CT);
    this.addRel(sheetPath, "comments", path);
    if (!this.vmlPath(sheetPath)) {
      const vml = this.uniquePath("xl/drawings/vmlDrawing", ".vml");
      this.putBinary(vml, new TextEncoder().encode(noteVml([])), VML_CT);
      const rId = this.addRel(sheetPath, "vmlDrawing", vml);
      const sheet = this.xml(sheetPath)!;
      const el = sheet.createElementNS(NS_S, "legacyDrawing");
      el.setAttributeNS(NS_R, "r:id", rId);
      const root = sheet.documentElement;
      root.insertBefore(el, Array.from(root.children).find((c) => LEGACY_AFTER.includes(c.localName)) ?? null);
      this.touch(sheetPath);
    }
    return this.xml(path)!;
  }

  /** Remove the sheet's comments part, its VML drawing and the sheet's link to it. */
  removeComments(sheetPath: string): void {
    const comments = this.rels(sheetPath).find((r) => r.type === "comments");
    if (comments) {
      this.removeRel(sheetPath, comments.id);
      this.removePart(comments.target);
    }
    const vml = this.rels(sheetPath).find((r) => r.type === "vmlDrawing");
    if (vml) {
      this.removeRel(sheetPath, vml.id);
      this.dropBinary(vml.target);
      const sheet = this.xml(sheetPath)!;
      Array.from(sheet.getElementsByTagNameNS(NS_S, "legacyDrawing")).forEach((e) => e.remove());
      this.touch(sheetPath);
    }
  }

  private vmlPath(sheetPath: string): string | null {
    return this.rels(sheetPath).find((r) => r.type === "vmlDrawing" && !r.external)?.target ?? null;
  }

  /** Before saving: notes that changed get their VML drawing rebuilt, so Excel shows the same notes. */
  async save(): Promise<Uint8Array> {
    for (const sheet of this.sheets()) {
      const comments = this.commentsPath(sheet.path);
      const vml = this.vmlPath(sheet.path);
      if (!comments || !vml || !this.isTouched(comments)) continue;
      const refs = Array.from(this.xml(comments)?.getElementsByTagNameNS(NS_S, "comment") ?? []).map((c) => c.getAttribute("ref") ?? "");
      this.putBinary(vml, new TextEncoder().encode(noteVml(refs)), VML_CT);
    }
    return super.save();
  }

  /** Add a chart part linked from `drawingPath`; returns its path and the drawing's relationship id. */
  addChartPart(drawingPath: string, xml: Document): { path: string; rId: string } {
    const path = this.uniquePath("xl/charts/chart", ".xml");
    this.putXml(path, xml, CHART_CT);
    return { path, rId: this.addRel(drawingPath, "chart", path) };
  }

  /** Remove a chart part and the drawing's relationship to it. */
  removeChartPart(drawingPath: string, rId: string): void {
    const target = this.relTarget(drawingPath, rId);
    this.removeRel(drawingPath, rId);
    if (target) this.removePart(target);
  }

  xmlPaths(): string[] {
    return [...this.parts.keys()];
  }

  /** Forget a part (undo of a structural edit restores its rels and content types separately). */
  removeXml(path: string): void {
    this.parts.delete(path);
  }

  /** Drop calcChain.xml: it lists formula cells, and Excel repairs the file when it no longer matches. Excel rebuilds it. */
  dropCalcChain(): void {
    const rel = this.rels(this.workbookPath).find((r) => r.type === "calcChain");
    if (!rel) return;
    this.removeRel(this.workbookPath, rel.id);
    this.removePart(rel.target);
  }
}
