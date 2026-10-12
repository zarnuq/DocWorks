import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { insertPivot, pivotAt, readPivots, refreshPivot } from "../src/domain/xlsx/pivot";
import { sheetXml, STYLES, xlsx } from "./xlsx-fixtures";

const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const num = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;
const ROWS: [string, string, number][] = [["East", "A", 10], ["East", "B", 20], ["West", "A", 5], ["West", "B", 15], ["East", "A", 1], ["North", "B", 3]];
const DATA = `<row r="1">${str("A1", "Region")}${str("B1", "Product")}${str("C1", "Amount")}</row>` +
  ROWS.map(([r, p, a], i) => `<row r="${i + 2}">${str(`A${i + 2}`, r)}${str(`B${i + 2}`, p)}${num(`C${i + 2}`, a)}</row>`).join("");
const regions = ["East", "West", "North"], products = ["A", "B"];
const CACHE = `<pivotCacheDefinition xmlns="${S}" xmlns:r="${R}" r:id="rId1" recordCount="6"><cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Data"/></cacheSource><cacheFields count="3">` +
  `<cacheField name="Region" numFmtId="0"><sharedItems count="3">${regions.map((v) => `<s v="${v}"/>`).join("")}</sharedItems></cacheField>` +
  `<cacheField name="Product" numFmtId="0"><sharedItems count="2">${products.map((v) => `<s v="${v}"/>`).join("")}</sharedItems></cacheField>` +
  `<cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="20"/></cacheField></cacheFields></pivotCacheDefinition>`;
const RECORDS = `<pivotCacheRecords xmlns="${S}" count="6">${ROWS.map(([r, p, a]) => `<r><x v="${regions.indexOf(r)}"/><x v="${products.indexOf(p)}"/><n v="${a}"/></r>`).join("")}</pivotCacheRecords>`;
const TABLE = `<pivotTableDefinition xmlns="${S}" name="PivotTable1" cacheId="1" dataCaption="Values" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" outline="1" outlineData="1" multipleFieldFilters="0">` +
  `<location ref="A3:B7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="4"><item x="0"/><item x="2"/><item x="1"/><item t="default"/></items></pivotField><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>` +
  `<rowFields count="1"><field x="0"/></rowFields><rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems>` +
  `<dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
const OUT = `<row r="3">${str("A3", "Row Labels")}${str("B3", "Sum of Amount")}</row><row r="4">${str("A4", "East")}${num("B4", 31)}</row><row r="5">${str("A5", "North")}${num("B5", 3)}</row><row r="6">${str("A6", "West")}${num("B6", 20)}</row><row r="7">${str("A7", "Grand Total")}${num("B7", 54)}</row>`;
const relsXml = (list: [string, string][]) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([t, target], i) => `<Relationship Id="rId${i + 1}" Type="${REL}${t}" Target="${target}"/>`).join("")}</Relationships>`;

async function setup() {
  const bytes = await xlsx({
    sheets: [sheetXml(DATA), sheetXml(OUT)],
    names: ["Data", "Pivot"],
    parts: {
      "xl/worksheets/_rels/sheet2.xml.rels": relsXml([["pivotTable", "../pivotTables/pivotTable1.xml"]]),
      "xl/pivotTables/pivotTable1.xml": TABLE,
      "xl/pivotTables/_rels/pivotTable1.xml.rels": relsXml([["pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml"]]),
      "xl/pivotCache/pivotCacheDefinition1.xml": CACHE,
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels": relsXml([["pivotCacheRecords", "pivotCacheRecords1.xml"]]),
      "xl/pivotCache/pivotCacheRecords1.xml": RECORDS,
    },
  });
  return new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
}
const grid = (e: Edits, sheet = 1) => {
  const s = e.wb.sheets[sheet];
  const used = s.used()!;
  return Array.from({ length: used.r2 - used.r1 + 1 }, (_, r) => Array.from({ length: used.c2 - used.c1 + 1 }, (_, c) => s.at(used.r1 + r, used.c1 + c)?.value ?? null));
};

describe("pivot tables", () => {
  it("reads a pivot table's fields, source and place", async () => {
    const e = await setup();
    const [p] = readPivots(e.wb.sheets[1]);
    expect(p).toMatchObject({ name: "PivotTable1", fields: ["Region", "Product", "Amount"], rows: [0], cols: [], values: [{ field: 2, fn: "sum", name: "Sum of Amount" }], source: { sheet: "Data", ref: { r1: 0, c1: 0, r2: 6, c2: 2 } }, location: { r1: 2, c1: 0, r2: 6, c2: 1 } });
    expect(pivotAt(e.wb.sheets[1], 4, 1)?.name).toBe("PivotTable1");
    expect(pivotAt(e.wb.sheets[1], 10, 1)).toBeNull();
  });

  it("refreshes from the source's current values", async () => {
    const e = await setup();
    e.input(e.wb.sheets[0], 6, 2, "100"); // North's amount
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0]);
    expect(grid(e)).toEqual([["Row Labels", "Sum of Amount"], ["East", 31], ["North", 100], ["West", 20], ["Grand Total", 151]]);
    // Excel refreshes it again on open; the cache's records match the data.
    const cache = e.wb.pkg.xml("xl/pivotCache/pivotCacheDefinition1.xml")!.documentElement;
    expect(cache.getAttribute("refreshOnLoad")).toBe("1");
    const recs = e.wb.pkg.xml("xl/pivotCache/pivotCacheRecords1.xml")!;
    expect(recs.getElementsByTagNameNS(S, "r")).toHaveLength(6);
    expect(recs.getElementsByTagNameNS(S, "n")[5].getAttribute("v")).toBe("100");
    e.undo();
    expect(grid(e)[2]).toEqual(["North", 3]);
  });

  it("changes rows, columns and values, writing the definition and the grown output", async () => {
    const e = await setup();
    const p = readPivots(e.wb.sheets[1])[0];
    refreshPivot(e, e.wb.sheets[1], p, { rows: [0], cols: [1], values: [{ field: 2, fn: "sum" }, { field: 2, fn: "count" }] });
    expect(grid(e)).toEqual([
      [null, "Column Labels", null, null, null, null, null],
      [null, "A", null, "B", null, "Total Sum of Amount", "Total Count of Amount"],
      ["Row Labels", "Sum of Amount", "Count of Amount", "Sum of Amount", "Count of Amount", null, null],
      ["East", 11, 2, 20, 1, 31, 3],
      ["North", null, null, 3, 1, 3, 1],
      ["West", 5, 1, 15, 1, 20, 2],
      ["Grand Total", 16, 3, 38, 3, 54, 6],
    ]);
    const q = readPivots(e.wb.sheets[1])[0];
    expect(q).toMatchObject({ rows: [0], cols: [1], values: [{ field: 2, fn: "sum", name: "Sum of Amount" }, { field: 2, fn: "count", name: "Count of Amount" }] });
    const def = e.wb.pkg.xml("xl/pivotTables/pivotTable1.xml")!;
    const fields = def.getElementsByTagNameNS(S, "pivotField");
    expect(fields[1].getAttribute("axis")).toBe("axisCol");
    expect(fields[1].getElementsByTagNameNS(S, "item")).toHaveLength(3); // A, B and the subtotal
    expect(Array.from(def.getElementsByTagNameNS(S, "colFields")[0].children).map((f) => f.getAttribute("x"))).toEqual(["1", "-2"]);
    expect(def.getElementsByTagNameNS(S, "location")[0].getAttribute("ref")).toBe("A3:G9");
  });

  it("nests a second row field under the first with subtotals, and shrinks back, clearing what it no longer covers", async () => {
    const e = await setup();
    const p = readPivots(e.wb.sheets[1])[0];
    refreshPivot(e, e.wb.sheets[1], p, { rows: [0, 1], cols: [], values: [{ field: 2, fn: "sum" }] });
    expect(grid(e)).toEqual([["Row Labels", "Sum of Amount"], ["East", 31], ["A", 11], ["B", 20], ["North", 3], ["B", 3], ["West", 20], ["A", 5], ["B", 15], ["Grand Total", 54]]);
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { rows: [1], cols: [], values: [{ field: 2, fn: "average" }] });
    expect(grid(e)).toEqual([["Row Labels", "Average of Amount"], ["A", 16 / 3], ["B", 38 / 3], ["Grand Total", 9]]);
  });
});

describe("pivot table editor form", () => {
  it("edits rows, the column field and values, and hands back the new layout", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    let got: unknown = null;
    pivotForm(host, { fields: ["Region", "Product", "Amount"], layout: { rows: [0], cols: [], values: [{ field: 2, fn: "sum" }] }, apply: (l) => (got = l) });
    const boxes = host.querySelectorAll<HTMLInputElement>("input[type=checkbox]");
    expect(Array.from(boxes).map((b) => b.checked)).toEqual([true, false, false]);
    boxes[1].click(); // Product as a second row field
    const col = host.querySelector<HTMLSelectElement>("select[aria-label='Column field']")!;
    col.value = "0";
    col.dispatchEvent(new Event("change"));
    host.querySelector<HTMLButtonElement>("button.xlsx-pivot-add")!.click();
    const fns = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Summarize by']");
    fns[1].value = "count";
    fns[1].dispatchEvent(new Event("change"));
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    // A field is on one axis: Region moved from the rows to the columns.
    expect(boxes[0].checked).toBe(false);
    expect(got).toEqual({ rows: [1], cols: [0], values: [{ field: 2, fn: "sum" }, { field: 0, fn: "count" }] });
  });
});

describe("insert pivot table", () => {
  it("makes a pivot table on a new sheet from a range (rows from the first text column, sum of the first number column), as one undo step", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(DATA)], names: ["Data"] });
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    const name = insertPivot(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 6, c2: 2 });
    expect(name).toBe("Pivot Table 1");
    const s = e.wb.sheets.find((x) => x.name === name)!;
    const [p] = readPivots(s);
    expect(p).toMatchObject({ rows: [0], cols: [], values: [{ field: 2, fn: "sum" }], source: { sheet: "Data" } });
    expect(s.at(0, 0)?.value).toBe("Row Labels");
    expect(s.at(4, 1)?.value).toBe(54);
    // The workbook lists the cache, and the saved file opens with the pivot table.
    const caches = e.wb.pkg.workbook.getElementsByTagNameNS(S, "pivotCache");
    expect(caches).toHaveLength(1);
    const again = Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
    expect(readPivots(again.sheets.find((x) => x.name === name)!)).toHaveLength(1);
    e.undo();
    expect(e.wb.sheets.map((x) => x.name)).toEqual(["Data"]);
    expect(e.wb.pkg.workbook.getElementsByTagNameNS(S, "pivotCache")).toHaveLength(0);
  });

  it("refuses a range without a header row and data", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(DATA)], names: ["Data"] });
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    expect(() => insertPivot(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 2 })).toThrow(/header row/i);
  });

  it("shows summed values in the source column's number format (counts stay plain)", async () => {
    const styled = DATA.replace(/<c r="(C\d+)"><v>/g, '<c r="$1" s="2"><v>');
    const bytes = await xlsx({ sheets: [sheetXml(styled)], names: ["Data"], styles: STYLES });
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    const name = insertPivot(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 6, c2: 2 });
    const s = e.wb.sheets.find((x) => x.name === name)!;
    expect(e.wb.styles.numFmt(s.at(1, 1)!.style)).toBe("yyyy-mm-dd");
    refreshPivot(e, e.wb.sheets.find((x) => x.name === name)!, readPivots(e.wb.sheets.find((x) => x.name === name)!)[0], { rows: [0], cols: [], values: [{ field: 2, fn: "count" }] });
    const t = e.wb.sheets.find((x) => x.name === name)!;
    expect(e.wb.styles.numFmt(t.at(1, 1)!.style)).toBe("General");
  });
});
