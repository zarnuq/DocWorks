import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { insertPivot, pivotAt, pivotItems, readPivots, refreshPivot } from "../src/domain/xlsx/pivot";
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

describe("unknown pivot subtotal functions", () => {
  it("treats an unrecognized subtotal like sum rather than min", async () => {
    const e = await setup();
    // Simulate an Excel file with subtotal="product" (not in our PivotFn union).
    const def = e.wb.pkg.xml("xl/pivotTables/pivotTable1.xml")!;
    def.getElementsByTagNameNS(S, "dataField")[0].setAttribute("subtotal", "product");
    const q = readPivots(e.wb.sheets[1])[0];
    expect(q.values[0].fn).toBe("sum");
    refreshPivot(e, e.wb.sheets[1], q);
    // Should produce the same totals as sum (the fallback), not min.
    expect(grid(e)).toEqual([["Row Labels", "Sum of Amount"], ["East", 31], ["North", 3], ["West", 20], ["Grand Total", 54]]);
  });
});

describe("pivot table column fields", () => {
  it("nests a second column field under the first, with a total after each group", async () => {
    const e = await setup();
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { rows: [], cols: [0, 1], values: [{ field: 2, fn: "sum" }] });
    expect(grid(e)).toEqual([
      ["Sum of Amount", "Column Labels", null, null, null, null, null, null, null, null],
      [null, "East", null, "East Total", "North", "North Total", "West", null, "West Total", "Grand Total"],
      ["Row Labels", "A", "B", null, "B", null, "A", "B", null, null],
      ["Grand Total", 11, 20, 31, 3, 3, 5, 15, 20, 54],
    ]);
    const def = e.wb.pkg.xml("xl/pivotTables/pivotTable1.xml")!;
    expect(Array.from(def.getElementsByTagNameNS(S, "colFields")[0].children).map((f) => f.getAttribute("x"))).toEqual(["0", "1"]);
    expect(Array.from(def.getElementsByTagNameNS(S, "pivotField")).map((f) => f.getAttribute("axis"))).toEqual(["axisCol", "axisCol", null]);
    const loc = def.getElementsByTagNameNS(S, "location")[0];
    expect([loc.getAttribute("ref"), loc.getAttribute("firstDataRow")]).toEqual(["A3:J6", "3"]);
    expect(readPivots(e.wb.sheets[1])[0]).toMatchObject({ rows: [], cols: [0, 1] });
  });

  it("puts several values under each column of two column fields", async () => {
    const e = await setup();
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { rows: [], cols: [0, 1], values: [{ field: 2, fn: "sum" }, { field: 2, fn: "count" }] });
    const g = grid(e);
    expect(g[0].slice(0, 2)).toEqual([null, "Column Labels"]);
    expect(g[1].slice(0, 8)).toEqual([null, "East", null, null, null, "East Total", null, "North"]);
    expect(g[1].slice(-2)).toEqual(["Total Sum of Amount", "Total Count of Amount"]);
    expect(g[2].slice(0, 6)).toEqual([null, "A", null, "B", null, null]);
    expect(g[3].slice(0, 7)).toEqual(["Row Labels", "Sum of Amount", "Count of Amount", "Sum of Amount", "Count of Amount", "Sum of Amount", "Count of Amount"]);
    expect(g[4]).toEqual(["Grand Total", 11, 2, 20, 1, 31, 3, 3, 1, 3, 1, 5, 1, 15, 1, 20, 2, 54, 6]);
  });
});

describe("pivot table filters", () => {
  const base = { rows: [0], cols: [], values: [{ field: 2, fn: "sum" as const }] };
  const def = (e: Edits) => e.wb.pkg.xml("xl/pivotTables/pivotTable1.xml")!;
  const attrs = (el: Element, names: string[]) => names.map((n) => el.getAttribute(n));

  it("lists a field's items for the filter's checklist", async () => {
    const e = await setup();
    const p = readPivots(e.wb.sheets[1])[0];
    expect(p.filters).toEqual([]);
    expect(pivotItems(e.wb.sheets[1], p, 0)).toEqual(["East", "North", "West"]);
    expect(pivotItems(e.wb.sheets[1], p, 2)).toEqual([1, 3, 5, 10, 15, 20]);
  });

  it("filters by a field off the table: a report filter above it, the table moved down, one item picked", async () => {
    const e = await setup();
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { ...base, filters: [{ field: 1, hidden: ["B"] }] });
    expect(grid(e)).toEqual([["Product", "A"], [null, null], ["Row Labels", "Sum of Amount"], ["East", 11], ["West", 5], ["Grand Total", 16]]);
    const loc = def(e).getElementsByTagNameNS(S, "location")[0];
    expect(attrs(loc, ["ref", "rowPageCount", "colPageCount"])).toEqual(["A5:B8", "1", "1"]);
    const field = def(e).getElementsByTagNameNS(S, "pivotField")[1];
    expect(field.getAttribute("axis")).toBe("axisPage");
    const page = def(e).getElementsByTagNameNS(S, "pageField");
    expect(page).toHaveLength(1);
    expect(attrs(page[0], ["fld", "item", "hier"])).toEqual(["1", "0", "-1"]);
    // pageFields sits between the column items and the data fields (CT_pivotTableDefinition's order).
    expect(page[0].parentElement!.nextElementSibling!.localName).toBe("dataFields");
    // The cache keeps every record: Excel filters when it draws.
    expect(e.wb.pkg.xml("xl/pivotCache/pivotCacheRecords1.xml")!.getElementsByTagNameNS(S, "r")).toHaveLength(6);
    const p = readPivots(e.wb.sheets[1])[0];
    expect(p.filters).toEqual([{ field: 1, hidden: ["B"] }]);
    expect(p.location).toEqual({ r1: 4, c1: 0, r2: 7, c2: 1 });
    // The filter's own cells belong to the pivot table (a click there offers its editor).
    expect(pivotAt(e.wb.sheets[1], 2, 0)?.name).toBe("PivotTable1");
    // A plain refresh keeps the filter and its place.
    e.input(e.wb.sheets[0], 1, 2, "110"); // East A: 10 -> 110
    refreshPivot(e, e.wb.sheets[1], p);
    expect(grid(e)).toEqual([["Product", "A"], [null, null], ["Row Labels", "Sum of Amount"], ["East", 111], ["West", 5], ["Grand Total", 116]]);
    // Without the filter the table is back where it was and the filter's cells are gone.
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { ...base, filters: [] });
    expect(grid(e)).toEqual([["Row Labels", "Sum of Amount"], ["East", 131], ["North", 3], ["West", 20], ["Grand Total", 154]]);
    expect(e.wb.sheets[1].used()).toMatchObject({ r1: 2, r2: 6 });
    expect(def(e).getElementsByTagNameNS(S, "pageFields")).toHaveLength(0);
    expect(def(e).getElementsByTagNameNS(S, "location")[0].hasAttribute("rowPageCount")).toBe(false);
  });

  it("shows several of a report filter's items, or all of them", async () => {
    const e = await setup();
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { rows: [1], cols: [], values: base.values, filters: [{ field: 0, hidden: ["North"] }] });
    expect(grid(e)).toEqual([["Region", "(Multiple Items)"], [null, null], ["Row Labels", "Sum of Amount"], ["A", 16], ["B", 35], ["Grand Total", 51]]);
    const field = def(e).getElementsByTagNameNS(S, "pivotField")[0];
    expect(field.getAttribute("multipleItemSelectionAllowed")).toBe("1");
    expect(Array.from(field.getElementsByTagNameNS(S, "item")).map((i) => attrs(i, ["x", "h", "t"]))).toEqual([["0", null, null], ["1", "1", null], ["2", null, null], [null, null, "default"]]);
    expect(def(e).getElementsByTagNameNS(S, "pageField")[0].hasAttribute("item")).toBe(false);
    expect(readPivots(e.wb.sheets[1])[0].filters).toEqual([{ field: 0, hidden: ["North"] }]);
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { rows: [1], cols: [], values: base.values, filters: [{ field: 0, hidden: [] }] });
    expect(grid(e).slice(0, 2)).toEqual([["Region", "(All)"], [null, null]]);
    expect(readPivots(e.wb.sheets[1])[0].filters).toEqual([{ field: 0, hidden: [] }]);
  });

  it("filters a field that is on the rows by hiding its items, with no report filter", async () => {
    const e = await setup();
    refreshPivot(e, e.wb.sheets[1], readPivots(e.wb.sheets[1])[0], { ...base, filters: [{ field: 0, hidden: ["East"] }] });
    expect(grid(e)).toEqual([["Row Labels", "Sum of Amount"], ["North", 3], ["West", 20], ["Grand Total", 23]]);
    expect(def(e).getElementsByTagNameNS(S, "pageFields")).toHaveLength(0);
    const items = def(e).getElementsByTagNameNS(S, "pivotField")[0].getElementsByTagNameNS(S, "item");
    expect(Array.from(items).map((i) => i.getAttribute("h"))).toEqual(["1", null, null, null]);
    expect(readPivots(e.wb.sheets[1])[0].filters).toEqual([{ field: 0, hidden: ["East"] }]);
    // Undo restores the unfiltered table.
    e.undo();
    expect(grid(e)[1]).toEqual(["East", 31]);
    expect(readPivots(e.wb.sheets[1])[0].filters).toEqual([]);
  });

  it("matches a date field's cache items (ISO dates) with the serial numbers in the cells", async () => {
    const e = await setup();
    // Product becomes a date column: 2024-01-01 (45292) for A, 2024-01-02 (45293) for B.
    e.run(() => { ROWS.forEach(([, p], i) => e.input(e.wb.sheets[0], i + 1, 1, p === "A" ? "45292" : "45293")); return []; });
    const cache = e.wb.pkg.xml("xl/pivotCache/pivotCacheDefinition1.xml")!;
    cache.getElementsByTagNameNS(S, "sharedItems")[1].innerHTML = `<d xmlns="${S}" v="2024-01-01T00:00:00"/><d xmlns="${S}" v="2024-01-02T00:00:00"/>`;
    const root = def(e).documentElement;
    const fields = root.getElementsByTagNameNS(S, "pivotField");
    fields[1].setAttribute("axis", "axisPage");
    fields[1].innerHTML = `<items xmlns="${S}" count="3"><item x="0"/><item x="1"/><item t="default"/></items>`;
    const page = def(e).createElementNS(S, "pageFields");
    page.innerHTML = `<pageField xmlns="${S}" fld="1" item="1" hier="-1"/>`;
    root.insertBefore(page, root.getElementsByTagNameNS(S, "dataFields")[0]);
    const p = readPivots(e.wb.sheets[1])[0];
    expect(p.filters).toEqual([{ field: 1, hidden: [45292] }]);
    refreshPivot(e, e.wb.sheets[1], p);
    expect(grid(e).slice(0, 1).concat(grid(e).slice(3))).toEqual([["Product", 45293], ["East", 20], ["North", 3], ["West", 15], ["Grand Total", 38]]);
  });

  it("finds the top of report filters laid out across columns from the rows they take, not their number", async () => {
    const e = await setup();
    const root = def(e).documentElement;
    const page = def(e).createElementNS(S, "pageFields");
    page.innerHTML = `<pageField xmlns="${S}" fld="1" hier="-1"/><pageField xmlns="${S}" fld="2" hier="-1"/>`;
    root.insertBefore(page, root.getElementsByTagNameNS(S, "dataFields")[0]);
    const loc = root.getElementsByTagNameNS(S, "location")[0];
    loc.setAttribute("ref", "A6:B10");
    // "Over, then down": both filters on one row, two rows above the table.
    loc.setAttribute("rowPageCount", "1");
    loc.setAttribute("colPageCount", "2");
    expect(readPivots(e.wb.sheets[1])[0].top).toBe(3);
    loc.removeAttribute("rowPageCount");
    expect(readPivots(e.wb.sheets[1])[0].top).toBe(2);
  });

  it("reads and refreshes a report filter Excel wrote", async () => {
    const e = await setup();
    const doc = def(e);
    const root = doc.documentElement;
    // Excel: Product on the page axis with "B" picked, shown in A1:B1 above the table.
    const fields = root.getElementsByTagNameNS(S, "pivotField");
    fields[1].setAttribute("axis", "axisPage");
    fields[1].innerHTML = `<items xmlns="${S}" count="3"><item x="0"/><item x="1"/><item t="default"/></items>`;
    const page = doc.createElementNS(S, "pageFields");
    page.setAttribute("count", "1");
    page.innerHTML = `<pageField xmlns="${S}" fld="1" item="1" hier="-1"/>`;
    root.insertBefore(page, root.getElementsByTagNameNS(S, "dataFields")[0]);
    root.getElementsByTagNameNS(S, "location")[0].setAttribute("rowPageCount", "1");
    const p = readPivots(e.wb.sheets[1])[0];
    expect(p.filters).toEqual([{ field: 1, hidden: ["A"] }]);
    refreshPivot(e, e.wb.sheets[1], p);
    expect(grid(e)).toEqual([["Product", "B"], [null, null], ["Row Labels", "Sum of Amount"], ["East", 20], ["North", 3], ["West", 15], ["Grand Total", 38]]);
    expect(def(e).getElementsByTagNameNS(S, "location")[0].getAttribute("ref")).toBe("A3:B7");
  });
});

describe("pivot table editor form", () => {
  it("edits rows, the column field and values, and hands back the new layout", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    let got: unknown = null;
    pivotForm(host, { fields: ["Region", "Product", "Amount"], layout: { rows: [0], cols: [], values: [{ field: 2, fn: "sum" }] }, apply: (l) => (got = l) });
    const boxes = host.querySelectorAll<HTMLInputElement>("input[aria-label^='Rows']");
    expect(Array.from(boxes).map((b) => b.checked)).toEqual([true, false, false]);
    boxes[1].click(); // Product as a second row field
    host.querySelector<HTMLInputElement>("input[aria-label='Columns: Region']")!.click();
    host.querySelector<HTMLButtonElement>("button.xlsx-pivot-add")!.click();
    // A new value starts on a field not used yet... (Region is now a column, Product a row, Amount a value: none left)
    const fields = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Value field']");
    expect(fields[1].value).toBe("0");
    fields[1].value = "0";
    const fns = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Summarize by']");
    fns[1].value = "count";
    fns[1].dispatchEvent(new Event("change"));
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    // A field is on one axis: Region moved from the rows to the columns.
    expect(boxes[0].checked).toBe(false);
    expect(got).toEqual({ rows: [1], cols: [0], values: [{ field: 2, fn: "sum" }, { field: 0, fn: "count" }], filters: [] });
  });

  it("adds filters: a field and the items it shows", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    let got: unknown = null;
    const items = [["East", "North", "West"], ["A", "B"], [1, 5]];
    pivotForm(host, { fields: ["Region", "Product", "Amount"], layout: { rows: [0], cols: [], values: [{ field: 2, fn: "sum" }], filters: [{ field: 0, hidden: ["North"] }] }, items: (f) => items[f], apply: (l) => (got = l) });
    const status = () => Array.from(host.querySelectorAll(".xlsx-pivot-filter-status")).map((b) => b.textContent);
    expect(status()).toEqual(["Showing 2 of 3 items"]);
    host.querySelector<HTMLButtonElement>("button.xlsx-pivot-add-filter")!.click();
    // A new filter starts on the first field without one, showing everything.
    const fields = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Filter field']");
    expect(Array.from(fields).map((f) => f.value)).toEqual(["0", "1"]);
    expect(status()).toEqual(["Showing 2 of 3 items", "Showing all items"]);
    host.querySelectorAll<HTMLButtonElement>(".xlsx-pivot-filter-status")[1].click();
    const boxes = host.querySelectorAll<HTMLInputElement>(".xlsx-pivot-filter-items input[type=checkbox]");
    expect(Array.from(boxes).map((b) => [b.parentElement!.textContent, b.checked])).toEqual([["A", true], ["B", true]]);
    boxes[1].click();
    expect(status()[1]).toBe("Showing 1 of 2 items");
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    expect(got).toEqual({ rows: [0], cols: [], values: [{ field: 2, fn: "sum" }], filters: [{ field: 0, hidden: ["North"] }, { field: 1, hidden: ["B"] }] });
    // Removing a filter shows everything again.
    host.querySelector<HTMLButtonElement>("button[aria-label='Remove filter']")!.click();
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    expect(got).toMatchObject({ filters: [{ field: 1, hidden: ["B"] }] });
  });

  it("keeps the fields' own order on the rows and columns, new ones last, and opens when the items can't be read", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    let got: { rows?: number[]; cols?: number[] } = {};
    const fields = ["Region", "Product", "Year", "Amount"];
    pivotForm(host, { fields, layout: { rows: [1], cols: [2, 0], values: [{ field: 3, fn: "sum" }], filters: [{ field: 1, hidden: ["B"] }] }, items: () => { throw new Error("no source"); }, apply: (l) => (got = l) });
    expect(host.querySelector(".xlsx-pivot-filter-status")!.textContent).toBe("Some items hidden");
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    expect([got.rows, got.cols]).toEqual([[1], [2, 0]]);
    // Region moves from the columns to the end of the rows.
    host.querySelector<HTMLInputElement>("input[aria-label='Rows: Region']")!.click();
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    expect([got.rows, got.cols]).toEqual([[1, 0], [2]]);
  });

  it("ticks a filter's items by a condition", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    let got: { filters?: unknown } = {};
    pivotForm(host, { fields: ["Region", "Amount"], layout: { rows: [0], cols: [], values: [{ field: 1, fn: "sum" }], filters: [{ field: 1, hidden: [] }] }, items: () => [1, 5, 10, 20], apply: (l) => (got = l) });
    host.querySelector<HTMLButtonElement>(".xlsx-pivot-filter-status")!.click();
    const kind = host.querySelector<HTMLSelectElement>("select[aria-label='Condition']")!;
    kind.value = "greater";
    host.querySelector<HTMLInputElement>("input[aria-label='Condition value']")!.value = "5";
    host.querySelector<HTMLButtonElement>("button.xlsx-pivot-filter-match")!.click();
    expect(Array.from(host.querySelectorAll<HTMLInputElement>(".xlsx-pivot-filter-items input[type=checkbox]")).map((b) => b.checked)).toEqual([false, false, true, true]);
    expect(host.querySelector(".xlsx-pivot-filter-status")!.textContent).toBe("Showing 2 of 4 items");
    host.querySelector<HTMLButtonElement>("button.mod-cta")!.click();
    expect(got.filters).toEqual([{ field: 1, hidden: [1, 5] }]);
  });

  it("adds a value on the first unused field, counting it when every field is taken", async () => {
    const { pivotForm } = await import("../src/presentation/xlsx/rule-forms");
    const host = document.body.createDiv();
    pivotForm(host, { fields: ["Region", "Product", "Amount"], layout: { rows: [0], cols: [], values: [{ field: 2, fn: "sum" }] }, apply: () => undefined });
    host.querySelector<HTMLButtonElement>("button.xlsx-pivot-add")!.click();
    const field = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Value field']")[1];
    const fn = host.querySelectorAll<HTMLSelectElement>("select[aria-label='Summarize by']")[1];
    expect([field.value, fn.value]).toEqual(["1", "count"]);
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
