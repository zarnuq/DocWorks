import { describe, it, expect } from "vitest";
import { colIndex, colName, parseCellRef, parseRange, rangeName, quoteSheet } from "../src/domain/xlsx/references";
import { parseFormula, translateFormula, shiftFormula, renameSheetInFormula, tokenize, FormulaSyntaxError } from "../src/domain/xlsx/formula";

describe("A1 references", () => {
  it("converts column letters both ways", () => {
    expect([colName(0), colName(25), colName(26), colName(701), colName(702), colName(16383)]).toEqual(["A", "Z", "AA", "ZZ", "AAA", "XFD"]);
    expect([colIndex("A"), colIndex("z"), colIndex("AA"), colIndex("XFD")]).toEqual([0, 25, 26, 16383]);
  });

  it("parses absolute and mixed cell references", () => {
    expect(parseCellRef("B3")).toEqual({ row: 2, col: 1, absRow: false, absCol: false });
    expect(parseCellRef("$B$3")).toEqual({ row: 2, col: 1, absRow: true, absCol: true });
    expect(parseCellRef("B$3")).toEqual({ row: 2, col: 1, absRow: true, absCol: false });
    expect(parseCellRef("3B")).toBeNull();
    expect(parseCellRef("XFE1")).toBeNull();
    expect(parseCellRef("A1048577")).toBeNull();
  });

  it("parses ranges, whole columns and whole rows, in any corner order", () => {
    expect(parseRange("B2:A1")).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
    expect(parseRange("C:C")).toEqual({ r1: 0, c1: 2, r2: 1048575, c2: 2 });
    expect(parseRange("2:3")).toEqual({ r1: 1, c1: 0, r2: 2, c2: 16383 });
    expect(parseRange("D4")).toEqual({ r1: 3, c1: 3, r2: 3, c2: 3 });
    expect(rangeName({ r1: 0, c1: 0, r2: 1, c2: 1 })).toBe("A1:B2");
    expect(rangeName({ r1: 3, c1: 3, r2: 3, c2: 3 })).toBe("D4");
  });

  it("quotes sheet names only when needed", () => {
    expect(quoteSheet("Data")).toBe("Data");
    expect(quoteSheet("My Data")).toBe("'My Data'");
    expect(quoteSheet("Bob's")).toBe("'Bob''s'");
    expect(quoteSheet("2024")).toBe("'2024'");
    expect(quoteSheet("A1")).toBe("'A1'");
  });
});

describe("formula parsing", () => {
  const ast = (f: string) => JSON.stringify(parseFormula(f));

  it("follows spreadsheet operator precedence", () => {
    expect(ast("1+2*3")).toBe(ast("1+(2*3)"));
    expect(ast("-2^2")).toBe(ast("(-2)^2"));
    expect(ast("1&2=12")).toBe(ast("(1&2)=12"));
    expect(ast("2^3^2")).toBe(ast("(2^3)^2"));
    expect(ast("50%*2")).toBe(ast("(50%)*2"));
  });

  it("reads references, sheet-qualified ranges, functions, strings, booleans, errors and arrays", () => {
    expect(parseFormula("SUM('My Data'!$A$1:B2, Totals!C:C)")).toMatchObject({
      type: "call",
      name: "SUM",
      args: [
        { type: "ref", sheet: "My Data", range: { r1: 0, c1: 0, r2: 1, c2: 1 } },
        { type: "ref", sheet: "Totals", range: { c1: 2, c2: 2 } },
      ],
    });
    expect(parseFormula(`"a""b"`)).toEqual({ type: "str", value: 'a"b' });
    expect(parseFormula("true")).toEqual({ type: "bool", value: true });
    expect(parseFormula("#N/A")).toEqual({ type: "err", value: "#N/A" });
    expect(parseFormula("{1,2;3,4}")).toEqual({ type: "array", rows: [[1, 2], [3, 4]] });
    expect(parseFormula("_xlfn.CONCAT(A1)")).toMatchObject({ type: "call", name: "CONCAT" });
    expect(parseFormula("Rate*2")).toMatchObject({ type: "bin", left: { type: "name", name: "Rate" } });
  });

  it("rejects malformed formulas", () => {
    for (const f of ["1+", "SUM(1,", "(1", "1)", '"abc', "A1:"]) expect(() => parseFormula(f), f).toThrow(FormulaSyntaxError);
  });

  it("keeps the original text in tokens so rewriting is lossless", () => {
    const f = "SUM( A1 , 'x y'!B2 )  +  1.50";
    expect(tokenize(f).map((t) => t.text).join("")).toBe(f);
  });
});

describe("formula reference rewriting", () => {
  it("translates relative references when a formula is copied", () => {
    expect(translateFormula("A1+$A$1+A$1+$A1", 1, 2)).toBe("C2+$A$1+C$1+$A2");
    expect(translateFormula("SUM(A1:B2)", 3, 0)).toBe("SUM(A4:B5)");
    expect(translateFormula("A1", -1, 0)).toBe("#REF!");
    expect(translateFormula("SUM(C:C)", 5, 1)).toBe("SUM(D:D)");
  });

  it("leaves strings, functions and names alone", () => {
    expect(translateFormula(`IF(A1="A1",LOG10(A1),Rate)`, 1, 0)).toBe(`IF(A2="A1",LOG10(A2),Rate)`);
  });

  it("shifts references for inserted rows, including absolute ones, only on the edited sheet", () => {
    const opts = { sheet: "Data", axis: "row" as const, at: 2, count: 2, formulaSheet: "Data" };
    expect(shiftFormula("A1+A3+$A$5+Totals!A3", opts)).toBe("A1+A5+$A$7+Totals!A3");
    expect(shiftFormula("SUM(A1:A3)", opts)).toBe("SUM(A1:A5)");
    expect(shiftFormula("Data!B4", { ...opts, formulaSheet: "Totals" })).toBe("Data!B6");
  });

  it("shrinks ranges and gives #REF! for deleted references", () => {
    const opts = { sheet: "Data", axis: "col" as const, at: 1, count: -2, formulaSheet: "Data" };
    expect(shiftFormula("A1+B1+D1", opts)).toBe("A1+#REF!+B1");
    expect(shiftFormula("SUM(A1:E1)", opts)).toBe("SUM(A1:C1)");
    expect(shiftFormula("SUM(B1:C1)", opts)).toBe("SUM(#REF!)");
  });

  it("renames sheets in references", () => {
    expect(renameSheetInFormula("Data!A1+'Data'!B2+Other!C3", "Data", "My Data")).toBe("'My Data'!A1+'My Data'!B2+Other!C3");
  });
});
