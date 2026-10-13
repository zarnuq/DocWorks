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


describe("insertPicture", () => {
  it("adds a picture over the cells at a cell (a drawing part, an image part), which reads back and saves", async () => {
    const { Edits } = await import("../src/domain/xlsx/commands");
    const { insertPicture } = await import("../src/domain/xlsx/drawing");
    const JSZip = (await import("jszip")).default;
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const sheet = e.wb.sheets[0];
    const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3]);
    insertPicture(e, sheet, png, "png", { row: 2, col: 1 }, { w: 120, h: 80 }, "Logo");
    const pics = sheetPictures(e.wb, sheet);
    expect(pics).toHaveLength(1);
    expect(pics[0]).toMatchObject({ name: "Logo", x: sheet.colX(1), y: sheet.rowY(2), w: 120, h: 80 });
    expect(Array.from(await e.wb.pkg.media(pics[0].image))).toEqual(Array.from(png));
    const zip = await JSZip.loadAsync(await e.wb.pkg.save());
    expect(Object.keys(zip.files).some((f) => /^xl\/media\/image\d*\.png$/.test(f))).toBe(true);
    expect(await zip.file("[Content_Types].xml")!.async("text")).toMatch(/Extension="png"|image\/png/);
    // A second picture goes into the same drawing.
    insertPicture(e, sheet, png, "png", { row: 0, col: 0 }, { w: 10, h: 10 }, "Two");
    expect(sheetPictures(e.wb, sheet).map((p) => p.name)).toEqual(["Logo", "Two"]);
  });
});

describe("moving and deleting pictures", () => {
  it("moves / resizes a picture by its key and deletes it with its image relationship, undoably", async () => {
    const { Edits } = await import("../src/domain/xlsx/commands");
    const { insertPicture, movePicture, deletePicture } = await import("../src/domain/xlsx/drawing");
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const sheet = e.wb.sheets[0];
    insertPicture(e, sheet, new Uint8Array([1, 2, 3]), "png", { row: 1, col: 1 }, { w: 100, h: 50 }, "Pic");
    const [pic] = sheetPictures(e.wb, sheet);
    expect(pic.key).toBeTruthy();
    movePicture(e, e.wb.sheets[0], pic.key, { x: 300, y: 120, w: 200, h: 100 });
    expect(sheetPictures(e.wb, e.wb.sheets[0])[0]).toMatchObject({ x: 300, y: 120, w: 200, h: 100 });
    deletePicture(e, e.wb.sheets[0], pic.key);
    expect(sheetPictures(e.wb, e.wb.sheets[0])).toHaveLength(0);
    const drawing = e.wb.pkg.rels(e.wb.sheets[0].ref.path).find((r) => r.type === "drawing")!.target;
    expect(e.wb.pkg.rels(drawing).filter((r) => r.type === "image")).toHaveLength(0);
    e.undo();
    expect(sheetPictures(e.wb, e.wb.sheets[0])).toHaveLength(1);
  });

  it("moves only the dragged picture when pictures share a cNvPr id (Google exports them all as 0)", async () => {
    const { Edits } = await import("../src/domain/xlsx/commands");
    const { movePicture } = await import("../src/domain/xlsx/drawing");
    const bytes = await xlsx({
      drawing: DRAWING,
      drawingRels: [["rId1", "../media/image1.png"], ["rId2", "../media/image2.png"]],
      media: { "xl/media/image1.png": new Uint8Array([1]), "xl/media/image2.png": new Uint8Array([2]) },
    });
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    const before = sheetPictures(e.wb, e.wb.sheets[0]);
    expect(new Set(before.map((p) => p.key)).size).toBe(3);
    movePicture(e, e.wb.sheets[0], before[1].key, { x: 300, y: 120, w: 80, h: 40 });
    const after = sheetPictures(e.wb, e.wb.sheets[0]);
    expect(after.map((p) => p.name)).toEqual(["one", "two", "abs"]);
    expect(after[1]).toMatchObject({ x: 300, y: 120 });
    expect(after[0]).toMatchObject({ x: before[0].x, y: before[0].y });
    expect(after[2]).toMatchObject({ x: before[2].x, y: before[2].y });
  });
});
