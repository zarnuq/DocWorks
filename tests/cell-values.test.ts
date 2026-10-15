import { describe, expect, it } from "vitest";
import type { App } from "obsidian";
import { cellValue, parseCellCode, renderCellValue } from "../src/presentation/common/cell-values";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { TFile } from "./obsidian-stub";
import { sheetXml, xlsx } from "./xlsx-fixtures";

const book = () => xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>alpha</t></is></c></row>`), sheetXml(`<row r="9"><c r="H9" s="0"><v>17912</v></c></row>`)], names: ["One", "Budget"] });

describe("cell values in notes", () => {
  it("reads `cell: file.xlsx#Sheet!B4` (with or without [[ ]])", () => {
    expect(parseCellCode("cell: Budget.xlsx#Budget!H9")).toEqual({ path: "Budget.xlsx", subpath: "#Budget!H9" });
    expect(parseCellCode("cell:[[My Budget.xlsx#B1]]")).toEqual({ path: "My Budget.xlsx", subpath: "#B1" });
    expect(parseCellCode("cell: Budget.xlsx")).toBeNull();
    expect(parseCellCode("const x = 1")).toBeNull();
  });

  it("gives the cell's shown value, on the first sheet when none is named", async () => {
    const wb = Workbook.open(await XlsxWorkbook.load(await book()));
    expect(cellValue(wb, "#Budget!H9")).toBe("17912");
    expect(cellValue(wb, "#B1")).toBe("alpha");
    expect(cellValue(wb, "#One!Z99")).toBe("");
    expect(cellValue(wb, "#Nope!A1")).toBeNull();
    expect(cellValue(wb, "#One")).toBeNull();
  });

  it("puts the value in place of the code, and says when the file is missing", async () => {
    const bytes = await book();
    const file = new TFile("Money/Budget.xlsx");
    const app = {
      metadataCache: { getFirstLinkpathDest: (path: string) => (path === "Budget.xlsx" ? file : null) },
      vault: { readBinary: async () => bytes.slice().buffer },
    } as unknown as App;
    const el = createSpan();
    await renderCellValue(app, el, { path: "Budget.xlsx", subpath: "#Budget!H9" }, "note.md");
    expect(el.textContent).toBe("17912");
    await renderCellValue(app, el, { path: "Gone.xlsx", subpath: "#A1" }, "note.md");
    expect(el.textContent).toBe("Gone.xlsx not found");
    expect(el.hasClass("is-error")).toBe(true);
  });
});
