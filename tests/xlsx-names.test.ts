import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { addName, deleteName, listNames, renameName, setNameRange, validateName } from "../src/domain/xlsx/names";
import { renameNameInFormula } from "../src/domain/xlsx/formula";
import { xlsx, DATA_SHEET, sheetXml } from "./xlsx-fixtures";

const NAMES =
  `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$C$5</definedName>` +
  `<definedName name="Amounts">Data!$A$2:$A$4</definedName></definedNames>`;
const setup = async () =>
  new Edits(
    Workbook.open(
      await XlsxWorkbook.load(
        await xlsx({ sheets: [DATA_SHEET, sheetXml(`<row r="1"><c r="A1"><f>SUM(Amounts)</f><v>60</v></c><c r="B1"><f>amounts_total+1</f><v>0</v></c></row>`)], workbookExtra: NAMES })
      )
    )
  );

describe("named ranges", () => {
  it("lists user names, not hidden built-in ones", async () => {
    const e = await setup();
    expect(listNames(e.wb)).toEqual([{ name: "Amounts", range: "Data!$A$2:$A$4" }]);
  });

  it("checks names like Google Sheets", async () => {
    const e = await setup();
    expect(() => validateName(e.wb, "")).toThrow();
    expect(() => validateName(e.wb, "two words")).toThrow(/letters, numbers/);
    expect(() => validateName(e.wb, "1st")).toThrow(/start/);
    expect(() => validateName(e.wb, "B12")).toThrow(/cell/);
    expect(() => validateName(e.wb, "true")).toThrow();
    expect(() => validateName(e.wb, "amounts")).toThrow(/already/);
    expect(() => validateName(e.wb, "Amounts", "Amounts")).not.toThrow();
    expect(() => validateName(e.wb, "Tax_2026")).not.toThrow();
  });

  it("adds a name with absolute references; formulas using it calculate", async () => {
    const e = await setup();
    addName(e, "Total", "Data!A5");
    expect(listNames(e.wb).find((n) => n.name === "Total")!.range).toBe("Data!$A$5");
    e.input(e.wb.sheets[1], 2, 0, "=Total*2");
    expect(e.wb.sheets[1].cell("A3")!.value).toBe(120);
    e.undo();
    e.undo();
    expect(listNames(e.wb).map((n) => n.name)).toEqual(["Amounts"]);
  });

  it("changes a name's range and recalculates", async () => {
    const e = await setup();
    setNameRange(e, "Amounts", "Data!A2:A3");
    expect(e.wb.sheets[1].cell("A1")!.value).toBe(30);
  });

  it("renames a name in every formula, leaving names that only start the same", async () => {
    expect(renameNameInFormula("SUM(Amounts)+amounts_total+\"Amounts\"", "amounts", "Sales")).toBe("SUM(Sales)+amounts_total+\"Amounts\"");
    const e = await setup();
    renameName(e, "Amounts", "Sales");
    expect(listNames(e.wb).map((n) => n.name)).toEqual(["Sales"]);
    expect(e.wb.sheets[1].cell("A1")).toMatchObject({ formula: "SUM(Sales)", value: 60 });
  });

  it("deletes a name; formulas using it show #NAME?", async () => {
    const e = await setup();
    deleteName(e, "Amounts");
    expect(listNames(e.wb)).toEqual([]);
    expect(e.wb.sheets[1].cell("A1")!.value).toEqual({ error: "#NAME?" });
  });

  it("refuses a range that isn't a reference", async () => {
    const e = await setup();
    expect(() => addName(e, "Bad", "nope")).toThrow(/range/);
  });
});
