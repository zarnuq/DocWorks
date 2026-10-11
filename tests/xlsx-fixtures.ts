/*
 * Hand-authored .xlsx workbooks for the spreadsheet tests. Each part is plain XML, so a test can see exactly
 * what the reader gets; `xlsx({ sheets })` swaps in other sheet XML for focused cases.
 */
import JSZip from "jszip";

export const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const CT = "application/vnd.openxmlformats-officedocument.";

export const sheetXml = (data: string, extra = { before: "", after: "" }) =>
  `<worksheet xmlns="${S}" xmlns:r="${R}">${extra.before}<sheetData>${data}</sheetData>${extra.after}</worksheet>`;

/** Sheet 1 ("Data"): styled literals, strings, a cached formula, a shared formula group, a merge, a date, a drawing. */
export const DATA_SHEET =
  `<worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="A1:D6"/>` +
  `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
  `<sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="20" customWidth="1"/><col min="4" max="4" width="9" hidden="1"/></cols>` +
  `<sheetData>` +
  `<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="s" s="1"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>` +
  `<row r="2"><c r="A2"><v>10</v></c><c r="B2" t="s"><v>3</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">A2*2</f><v>20</v></c></row>` +
  `<row r="3"><c r="A3"><v>20</v></c><c r="B3" t="inlineStr"><is><t>inline</t></is></c><c r="C3"><f t="shared" si="0"/><v>40</v></c></row>` +
  `<row r="4" ht="30" customHeight="1"><c r="A4"><v>30</v></c><c r="B4" t="b"><v>1</v></c><c r="C4"><f t="shared" si="0"/><v>60</v></c></row>` +
  `<row r="5"><c r="A5"><f>SUM(A2:A4)</f><v>60</v></c><c r="B5" s="2"><v>45292</v></c><c r="C5" t="e"><v>#DIV/0!</v></c></row>` +
  `<row r="6" hidden="1"><c r="A6" t="str"><f>Totals!A1&amp;"!"</f><v>x!</v></c></row>` +
  `</sheetData><mergeCells count="1"><mergeCell ref="B6:C6"/></mergeCells>` +
  `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></worksheet>`;

const TOTALS_SHEET = sheetXml(`<row r="1"><c r="A1" t="s"><v>4</v></c><c r="B1"><f>SUM(Data!A2:A4)</f><v>60</v></c></row>`);

export const STRINGS =
  `<sst xmlns="${S}" count="5" uniqueCount="5"><si><t>Name</t></si><si><t>Notes</t></si><si><t>Double</t></si>` +
  `<si><r><rPr><b/></rPr><t>rich</t></r><r><t xml:space="preserve"> text</t></r></si><si><t>x</t></si></sst>`;

export const STYLES =
  `<styleSheet xmlns="${S}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>` +
  `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font></fonts>` +
  `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>` +
  `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>` +
  `<border><left/><right/><top/><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>` +
  `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

// Spacing and an explicit end tag a serializer would normalize: proves untouched parts keep their bytes.
const DRAWING = `<xdr:wsDr  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>`;

const types = (sheets: number) =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/xl/workbook.xml" ContentType="${CT}spreadsheetml.sheet.main+xml"/>` +
  Array.from({ length: sheets }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT}spreadsheetml.worksheet+xml"/>`).join("") +
  `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}spreadsheetml.sharedStrings+xml"/>` +
  `<Override PartName="/xl/styles.xml" ContentType="${CT}spreadsheetml.styles+xml"/>` +
  `<Override PartName="/xl/calcChain.xml" ContentType="${CT}spreadsheetml.calcChain+xml"/>` +
  `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CT}drawing+xml"/></Types>`;

const rels = (list: [string, string][]) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  list.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${REL}${type}" Target="${target}"/>`).join("") +
  `</Relationships>`;

export interface FixtureOptions {
  /** Sheet XML per sheet; names default to Data, Totals, Sheet3... */
  sheets?: string[];
  names?: string[];
  workbookExtra?: string;
  strings?: string;
  styles?: string;
  /** Sheet 1's drawing part, with its relationships (rId -> media path) and media files. */
  drawing?: string;
  drawingRels?: [string, string][];
  media?: Record<string, Uint8Array>;
}

export async function xlsx(o: FixtureOptions = {}): Promise<Uint8Array> {
  const sheets = o.sheets ?? [DATA_SHEET, TOTALS_SHEET];
  const names = o.names ?? ["Data", "Totals", "Sheet3", "Sheet4"];
  const z = new JSZip();
  const put = (path: string, xml: string) => z.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`);
  put("[Content_Types].xml", types(sheets.length));
  put("_rels/.rels", rels([["officeDocument", "xl/workbook.xml"]]));
  // Sheets are listed in reverse relationship order so code can't rely on rId order matching tab order.
  put(
    "xl/workbook.xml",
    `<workbook xmlns="${S}" xmlns:r="${R}"><workbookPr/><bookViews><workbookView activeTab="0"/></bookViews><sheets>` +
      sheets.map((_, i) => `<sheet name="${names[i]}" sheetId="${i + 7}" r:id="rId${sheets.length - i}"/>`).join("") +
      `</sheets>${o.workbookExtra ?? ""}<calcPr calcId="191029"/></workbook>`
  );
  put(
    "xl/_rels/workbook.xml.rels",
    rels([
      ...sheets.map((_, i): [string, string] => ["worksheet", `worksheets/sheet${sheets.length - i}.xml`]),
      ["sharedStrings", "sharedStrings.xml"],
      ["styles", "styles.xml"],
      ["calcChain", "calcChain.xml"],
    ])
  );
  sheets.forEach((xml, i) => put(`xl/worksheets/sheet${i + 1}.xml`, xml));
  put("xl/worksheets/_rels/sheet1.xml.rels", rels([["drawing", "../drawings/drawing1.xml"]]));
  put("xl/drawings/drawing1.xml", o.drawing ?? DRAWING);
  if (o.drawingRels) put("xl/drawings/_rels/drawing1.xml.rels", rels(o.drawingRels.map(([, target]) => ["image", target])));
  for (const [path, bytes] of Object.entries(o.media ?? {})) z.file(path, bytes);
  put("xl/sharedStrings.xml", o.strings ?? STRINGS);
  put("xl/styles.xml", o.styles ?? STYLES);
  put("xl/calcChain.xml", `<calcChain xmlns="${S}"><c r="C2" i="1"/></calcChain>`);
  return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
