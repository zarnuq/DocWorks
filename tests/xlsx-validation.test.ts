import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { checkValue, listItems, readValidations, removeValidation, setValidation, validationAt } from "../src/domain/xlsx/validation";
import { sheetXml, S, xlsx } from "./xlsx-fixtures";

const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const DATA = `<row r="1">${str("A1", "Red")}${str("B1", "Yes")}</row><row r="2">${str("A2", "Green")}</row><row r="3">${str("A3", "Blue")}</row><row r="4">${str("A4", "Red")}</row>`;
const DV = (body: string) => `<dataValidations count="1">${body}</dataValidations>`;

async function setup(dv: string) {
  const bytes = await xlsx({ sheets: [sheetXml(DATA, { before: "", after: dv })], names: ["Data"] });
  const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
  const evaluate = (row: number, col: number, f: string) => e.calc.evaluate(e.wb.sheets[0], row, col, f);
  return { e, sheet: () => e.wb.sheets[0], evaluate };
}

describe("data validation", () => {
  it("reads Google's dropdowns: a quoted item list over several ranges", async () => {
    const { sheet } = await setup(DV(`<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B1:B3 D1"><formula1>"No,Yes,Wrong"</formula1></dataValidation>`));
    const v = validationAt(sheet(), 2, 1)!;
    expect(v.items).toEqual(["No", "Yes", "Wrong"]);
    expect(v.reject).toBe(true);
    expect(validationAt(sheet(), 0, 3)).toBe(v);
    expect(validationAt(sheet(), 3, 1)).toBeNull();
    expect(listItems(v, sheet())).toEqual(["No", "Yes", "Wrong"]);
  });

  it("lists items from a range once each, skipping blanks", async () => {
    const { sheet } = await setup(DV(`<dataValidation type="list" sqref="C1"><formula1>Data!$A$1:$A$6</formula1></dataValidation>`));
    const v = validationAt(sheet(), 0, 2)!;
    expect(v.items).toBeNull();
    expect(listItems(v, sheet())).toEqual(["Red", "Green", "Blue"]);
    expect(v.reject).toBe(false);
  });

  it("checks values: list items (any case), numbers, whole numbers, custom formulas and blanks", async () => {
    const { e, sheet, evaluate } = await setup(
      DV(
        `<dataValidation type="list" allowBlank="1" sqref="B1"><formula1>"No,Yes"</formula1></dataValidation>` +
          `<dataValidation type="decimal" operator="between" sqref="C1"><formula1>1</formula1><formula2>10</formula2></dataValidation>` +
          `<dataValidation type="whole" operator="greaterThan" sqref="C2"><formula1>0</formula1></dataValidation>` +
          `<dataValidation type="custom" sqref="D1:D4"><formula1>LEN(D1)&lt;4</formula1></dataValidation>`
      )
    );
    const at = (r: number, c: number) => validationAt(sheet(), r, c)!;
    expect(checkValue(at(0, 1), "yes", sheet(), 0, 1, evaluate)).toBe(true);
    expect(checkValue(at(0, 1), "Maybe", sheet(), 0, 1, evaluate)).toBe(false);
    expect(checkValue(at(0, 1), null, sheet(), 0, 1, evaluate)).toBe(true);
    expect(checkValue(at(0, 2), 5, sheet(), 0, 2, evaluate)).toBe(true);
    expect(checkValue(at(0, 2), 11, sheet(), 0, 2, evaluate)).toBe(false);
    expect(checkValue(at(0, 2), "x", sheet(), 0, 2, evaluate)).toBe(false);
    expect(checkValue(at(1, 2), 2.5, sheet(), 1, 2, evaluate)).toBe(false);
    expect(checkValue(at(1, 2), 3, sheet(), 1, 2, evaluate)).toBe(true);
    // Custom formulas are checked once the value is in the cell, relative to each cell.
    e.input(sheet(), 2, 3, "abc");
    expect(checkValue(at(2, 3), "abc", sheet(), 2, 3, evaluate)).toBe(true);
    e.input(sheet(), 2, 3, "abcdef");
    expect(checkValue(at(2, 3), "abcdef", sheet(), 2, 3, evaluate)).toBe(false);
  });

  it("sets a rule on a range, replacing the part of older rules it covers, as one undo step", async () => {
    const { e, sheet } = await setup(DV(`<dataValidation type="list" sqref="A1:A4"><formula1>"x,y"</formula1></dataValidation>`));
    setValidation(e, sheet(), [{ r1: 1, c1: 0, r2: 2, c2: 0 }], { type: "list", items: ["Red", "Green"], reject: true });
    expect(validationAt(sheet(), 0, 0)!.items).toEqual(["x", "y"]);
    expect(validationAt(sheet(), 1, 0)!.items).toEqual(["Red", "Green"]);
    expect(validationAt(sheet(), 3, 0)!.items).toEqual(["x", "y"]);
    expect(readValidations(sheet())).toHaveLength(2);
    const dvs = sheet().doc.getElementsByTagNameNS(S, "dataValidations")[0];
    expect(dvs.getAttribute("count")).toBe("2");
    e.undo();
    expect(readValidations(sheet())).toHaveLength(1);
    expect(validationAt(sheet(), 1, 0)!.items).toEqual(["x", "y"]);
  });

  it("removes rules from a range, and the container when nothing is left", async () => {
    const { e, sheet } = await setup(DV(`<dataValidation type="list" sqref="A1:A4"><formula1>"x,y"</formula1></dataValidation>`));
    removeValidation(e, sheet(), [{ r1: 0, c1: 0, r2: 3, c2: 0 }]);
    expect(readValidations(sheet())).toHaveLength(0);
    expect(sheet().doc.getElementsByTagNameNS(S, "dataValidations")).toHaveLength(0);
  });

  it("writes new rules where Excel expects them and with Google's quoting", async () => {
    const { e, sheet } = await setup("");
    setValidation(e, sheet(), [{ r1: 0, c1: 1, r2: 9, c2: 1 }], { type: "list", items: ['Say "hi"', "No"], reject: false });
    const el = sheet().doc.getElementsByTagNameNS(S, "dataValidation")[0];
    expect(el.getAttribute("sqref")).toBe("B1:B10");
    expect(el.getAttribute("errorStyle")).toBe("warning");
    expect(el.getElementsByTagNameNS(S, "formula1")[0].textContent).toBe('"Say ""hi"",No"');
    expect(validationAt(sheet(), 4, 1)!.items).toEqual(['Say "hi"', "No"]);
    setValidation(e, sheet(), [{ r1: 0, c1: 2, r2: 0, c2: 2 }], { type: "decimal", operator: "between", formulas: ["1", "5"], reject: true });
    expect(validationAt(sheet(), 0, 2)).toMatchObject({ type: "decimal", operator: "between", formulas: ["1", "5"], reject: true });
  });
});
