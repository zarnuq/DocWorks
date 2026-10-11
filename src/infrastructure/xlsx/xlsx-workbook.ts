import JSZip from "jszip";
import { OoxmlPackage, parseXml } from "../ooxml/ooxml-package";

export const NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const STRINGS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

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

  /** Remove a sheet's part, relationship and tab (its own rels and parts only it uses stay unreferenced). */
  deleteSheet(sheet: SheetRef): void {
    sheet.el.remove();
    this.removeRel(this.workbookPath, sheet.rId);
    this.removePart(sheet.path);
    this.touch(this.workbookPath);
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
