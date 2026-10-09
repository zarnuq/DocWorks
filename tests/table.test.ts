import { describe, it, expect } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { parseSlide, ShapeModel } from "../src/domain/pptx/slide-model";
import { setRect } from "../src/domain/pptx/slide-edit";
import {
  addTable,
  deleteCol,
  deleteRow,
  insertCol,
  insertRow,
  mergeCells,
  setCellBorders,
  setCellFill,
  setColWidth,
  setRowHeight,
  setTableFlags,
  splitCell,
  TABLE_STYLES,
} from "../src/domain/pptx/table";
import { loadSample, roundTrip, texts } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const MEDIUM_2 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

async function deck(rows = 3, cols = 3) {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  const frame = addTable(doc.xml(path)!, rows, cols, { x: 914400, y: 914400, w: 3 * 914400 * cols, h: 457200 * rows }, MEDIUM_2);
  doc.ensureTableStyle(MEDIUM_2);
  const model = () => parseSlide(doc, path);
  const table = () => model().shapes.find((s) => s.el === frame)!;
  const cell = (r: number, c: number) => model().shapes.find((s) => s.kind === "cell" && s.cell!.r === r && s.cell!.c === c)!;
  const tbl = () => frame.getElementsByTagNameNS(A, "tbl")[0];
  const grid = () => Array.from(tbl().getElementsByTagNameNS(A, "tr")).map((tr) => Array.from(tr.getElementsByTagNameNS(A, "tc")).map((tc) => texts(tc).join("")));
  return { doc, path, frame, model, table, cell, tbl, grid };
}

const fillTextIn = (tbl: Element) =>
  Array.from(tbl.getElementsByTagNameNS(A, "tr")).forEach((tr, r) =>
    Array.from(tr.getElementsByTagNameNS(A, "tc")).forEach((tc, c) => {
      const p = tc.getElementsByTagNameNS(A, "p")[0];
      p.innerHTML = `<a:r xmlns:a="${A}"><a:t>${r}${c}</a:t></a:r>`;
    })
  );

describe("reading tables", () => {
  it("models the grid and gives each cell its own text shape after the table", async () => {
    const doc = await loadSample();
    const shapes = parseSlide(doc, doc.slides()[2].path).shapes;
    const i = shapes.findIndex((s) => s.name === "Table 5");
    const t = shapes[i];
    expect(t.kind).toBe("table");
    expect(t.table!.cols).toEqual([1828800, 1828800]);
    expect(t.table!.rows).toEqual([457200, 457200]);
    expect(t.table!.flags).toMatchObject({ firstRow: true, bandRow: true, firstCol: false });
    const cells = shapes.slice(i + 1, i + 5);
    expect(cells.map((s) => [s.kind, s.cell!.r, s.cell!.c, s.text!.paras[0].runs[0].text])).toEqual([
      ["cell", 0, 0, "A1"],
      ["cell", 0, 1, "B1"],
      ["cell", 1, 0, "A2"],
      ["cell", 1, 1, "B2"],
    ]);
    expect(cells[3].rect).toEqual({ x: 914400 + 1828800, y: 4572000 + 457200, w: 1828800, h: 457200 });
    expect(cells.every((c) => c.grouped && c.text!.insets.l === 91440)).toBe(true);
  });

  it("applies the table style by region: header, bands and borders", async () => {
    const { cell } = await deck();
    const header = cell(0, 0);
    expect(header.cell!.fill).toEqual({ type: "solid", color: "#156082" });
    expect(header.text!.paras[0].defaults).toMatchObject({ b: true, color: "#FFFFFF" });
    expect(header.cell!.borders.b).toMatchObject({ color: "#FFFFFF", width: 38100 });

    const band1 = cell(1, 1).cell!.fill as { color: string };
    const band2 = cell(2, 1).cell!.fill as { color: string };
    expect(band1.color).not.toBe(band2.color);
    expect(cell(1, 1).text!.paras[0].defaults).toMatchObject({ b: false });
    expect(cell(1, 1).cell!.borders.t).toMatchObject({ color: "#FFFFFF", width: 12700 });
  });

  it("lets explicit cell formatting override the style", async () => {
    const { tbl, cell } = await deck();
    setCellFill(tbl(), [[1, 1]], { type: "solid", color: "#00FF00" });
    setCellBorders(tbl(), [1, 1, 1, 1], "all", { color: "#FF0000", width: 25400 });
    expect(cell(1, 1).cell!.fill).toEqual({ type: "solid", color: "#00FF00" });
    expect(cell(1, 1).cell!.borders.l).toMatchObject({ color: "#FF0000", width: 25400 });
    setCellBorders(tbl(), [1, 1, 1, 1], "all", null);
    expect(cell(1, 1).cell!.borders.l).toBeNull();
  });
});

describe("editing tables", () => {
  it("inserts a styled table sized to its grid", async () => {
    const { doc, table, frame } = await deck(2, 4);
    expect(table().table!.cols).toHaveLength(4);
    expect(table().rect.h).toBe(2 * 457200);
    expect(doc.xml("ppt/tableStyles.xml")!.getElementsByTagNameNS(A, "tblStyle")).toHaveLength(1);
    doc.ensureTableStyle(MEDIUM_2);
    expect(doc.xml("ppt/tableStyles.xml")!.getElementsByTagNameNS(A, "tblStyle")).toHaveLength(1);
    const back = await roundTrip(doc);
    expect(parseSlide(back, back.slides()[0].path).shapes.filter((s) => s.kind === "cell")).toHaveLength(8);
    expect(frame.getElementsByTagNameNS(A, "tableStyleId")[0].textContent).toBe(MEDIUM_2);
    expect(TABLE_STYLES.some((s) => s.id === MEDIUM_2)).toBe(true);
  });

  it("inserts and deletes rows and columns, keeping the frame size in step", async () => {
    const { tbl, grid, table } = await deck(2, 2);
    fillTextIn(tbl());
    insertRow(tbl(), 0, true);
    expect(grid()).toEqual([["00", "01"], ["", ""], ["10", "11"]]);
    insertCol(tbl(), 1, false);
    expect(grid()).toEqual([["00", "", "01"], ["", "", ""], ["10", "", "11"]]);
    expect(table().rect.w).toBe(table().table!.cols.reduce((a, b) => a + b));
    expect(table().rect.h).toBe(3 * 457200);
    deleteRow(tbl(), 1);
    deleteCol(tbl(), 1);
    expect(grid()).toEqual([["00", "01"], ["10", "11"]]);
    expect(table().rect.h).toBe(2 * 457200);
  });

  it("merges a range into its top-left cell and splits it again", async () => {
    const { tbl, cell, table } = await deck(3, 3);
    fillTextIn(tbl());
    mergeCells(tbl(), [0, 0, 1, 1]);
    const origin = cell(0, 0);
    expect([origin.cell!.rowSpan, origin.cell!.gridSpan]).toEqual([2, 2]);
    expect(origin.rect.w).toBe(table().table!.cols[0] + table().table!.cols[1]);
    expect(origin.text!.paras.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["00", "01", "10", "11"]);
    expect(cell(1, 1).cell!.hidden).toBe(true);

    // A row inserted inside the merge extends it.
    insertRow(tbl(), 0, true);
    expect(cell(0, 0).cell!.rowSpan).toBe(3);
    expect(cell(1, 0).cell!.hidden).toBe(true);
    deleteRow(tbl(), 0);
    expect(cell(0, 0).cell!.rowSpan).toBe(2);
    expect(cell(0, 0).text!.paras).toHaveLength(4);

    splitCell(tbl(), 0, 0);
    expect([cell(0, 0).cell!.rowSpan, cell(0, 0).cell!.gridSpan, cell(1, 1).cell!.hidden]).toEqual([1, 1, false]);
  });

  it("deleting the first row of a vertical merge moves its content down", async () => {
    const { tbl, cell } = await deck(3, 2);
    fillTextIn(tbl());
    mergeCells(tbl(), [0, 0, 1, 0]);
    deleteRow(tbl(), 0);
    expect(cell(0, 0).cell!.rowSpan).toBe(1);
    expect(cell(0, 0).cell!.hidden).toBe(false);
    expect(cell(0, 0).text!.paras.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["00", "10"]);
  });

  it("sets outer borders around a range only", async () => {
    const { tbl, cell } = await deck(3, 3);
    setCellBorders(tbl(), [0, 0, 1, 1], "outer", { color: "#FF0000", width: 12700 });
    expect(cell(0, 0).cell!.borders.t?.color).toBe("#FF0000");
    expect(cell(0, 0).cell!.borders.l?.color).toBe("#FF0000");
    expect(cell(0, 0).cell!.borders.r?.color).not.toBe("#FF0000");
    expect(cell(1, 1).cell!.borders.b?.color).toBe("#FF0000");
  });

  it("toggles style options", async () => {
    const { tbl, cell } = await deck();
    setTableFlags(tbl(), { firstRow: false, firstCol: true });
    expect(cell(0, 1).text!.paras[0].defaults.b).toBe(false);
    expect(cell(1, 0).text!.paras[0].defaults.b).toBe(true);
  });

  it("resizes columns and rows, and scales the grid with the frame", async () => {
    const { tbl, table, frame } = await deck(2, 2);
    setColWidth(tbl(), 0, 914400);
    setRowHeight(tbl(), 1, 914400);
    expect(table().table!.cols).toEqual([914400, 2743200]);
    expect(table().rect).toMatchObject({ w: 914400 + 2743200, h: 457200 + 914400 });
    setRect(frame, { x: 0, y: 0, w: 2 * (914400 + 2743200), h: 457200 + 914400 });
    expect(table().table!.cols).toEqual([2 * 914400, 2 * 2743200]);
    expect(table().table!.rows).toEqual([457200, 914400]);
  });
});
