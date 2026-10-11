import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { sheetPictures } from "../src/domain/xlsx/drawing";
import { xlsx } from "./xlsx-fixtures";

const X = `xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const pic = (rId: string, name = "pic") =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="${name}" descr="alt text"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rId}"/></xdr:blipFill><xdr:spPr/></xdr:pic>`;
const at = (tag: string, col: number, colOff: number, row: number, rowOff: number) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;

const DRAWING =
  `<xdr:wsDr ${X}>` +
  `<xdr:oneCellAnchor>${at("from", 2, 9525 * 4, 1, 0)}<xdr:ext cx="${9525 * 50}" cy="${9525 * 30}"/>${pic("rId1", "one")}<xdr:clientData/></xdr:oneCellAnchor>` +
  `<xdr:twoCellAnchor>${at("from", 0, 0, 0, 0)}${at("to", 1, 9525 * 10, 2, 0)}${pic("rId2", "two")}<xdr:clientData/></xdr:twoCellAnchor>` +
  `<xdr:absoluteAnchor><xdr:pos x="${9525 * 5}" y="${9525 * 7}"/><xdr:ext cx="${9525 * 20}" cy="${9525 * 20}"/>${pic("rId1", "abs")}<xdr:clientData/></xdr:absoluteAnchor>` +
  `<xdr:twoCellAnchor>${at("from", 0, 0, 0, 0)}${at("to", 1, 0, 1, 0)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9" name="shape"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>` +
  `</xdr:wsDr>`;

describe("sheet pictures", () => {
  it("places pictures from one-cell, two-cell and absolute anchors in sheet pixels", async () => {
    const bytes = await xlsx({
      drawing: DRAWING,
      drawingRels: [["rId1", "../media/image1.png"], ["rId2", "../media/image2.png"]],
      media: { "xl/media/image1.png": new Uint8Array([1]), "xl/media/image2.png": new Uint8Array([2]) },
    });
    const wb = Workbook.open(await XlsxWorkbook.load(bytes));
    const pics = sheetPictures(wb, wb.sheets[0]);
    expect(pics.map((p) => p.name)).toEqual(["one", "two", "abs"]);
    // Column A is 64px, B 140px; rows are 20px with row 4 at 40px.
    expect(pics[0]).toMatchObject({ image: "xl/media/image1.png", x: 64 + 140 + 4, y: 20, w: 50, h: 30, alt: "alt text" });
    expect(pics[1]).toMatchObject({ image: "xl/media/image2.png", x: 0, y: 0, w: 64 + 10, h: 40 });
    expect(pics[2]).toMatchObject({ x: 5, y: 7, w: 20, h: 20 });
  });

  it("finds nothing on a sheet without a drawing", async () => {
    const wb = Workbook.open(await XlsxWorkbook.load(await xlsx()));
    expect(sheetPictures(wb, wb.sheets[1])).toEqual([]);
  });

  it("fits a picture placed exactly at a cell's corner and smaller than it to the cell, like an image in a cell", async () => {
    const drawing = `<xdr:wsDr ${X}><xdr:oneCellAnchor>${at("from", 1, 0, 5, 0)}<xdr:ext cx="${9525 * 10}" cy="${9525 * 10}"/>${pic("rId1")}<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const bytes = await xlsx({ drawing, drawingRels: [["rId1", "../media/image1.png"]], media: { "xl/media/image1.png": new Uint8Array([1]) } });
    const wb = Workbook.open(await XlsxWorkbook.load(bytes));
    // B6 is merged into B6:C6: 140 + 64 wide, and row 6 is hidden (0 tall) in the fixture, so unhide it first.
    const s = wb.sheets[0];
    s.rows.set(5, 40);
    const [p] = sheetPictures(wb, s);
    expect(p).toMatchObject({ w: 40, h: 40, y: s.rowY(5) });
    expect(p.x).toBeCloseTo(64 + (204 - 40) / 2);
  });
});

