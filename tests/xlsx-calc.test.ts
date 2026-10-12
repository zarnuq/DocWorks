import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Calc } from "../src/domain/xlsx/calculation";
import { parseCellRef } from "../src/domain/xlsx/references";
import { xlsx, sheetXml, S } from "./xlsx-fixtures";

type Cells = Record<string, number | string | boolean | null>;
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** A workbook from { Sheet: { A1: 1, A2: "=A1+1", A3: "text" } }; formulas have no cached value. */
async function book(sheets: Record<string, Cells>) {
  const xml = Object.values(sheets).map((cells) => {
    const rows = new Map<number, string[]>();
    for (const [ref, v] of Object.entries(cells)) {
      const r = parseCellRef(ref)!;
      let c: string;
      if (typeof v === "string" && v.startsWith("=")) c = `<c r="${ref}"><f>${esc(v.slice(1))}</f></c>`;
      else if (typeof v === "string") c = `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
      else if (typeof v === "boolean") c = `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
      else c = `<c r="${ref}"><v>${v}</v></c>`;
      rows.set(r.row, [...(rows.get(r.row) ?? []), c]);
    }
    return sheetXml([...rows.entries()].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r + 1}">${cs.join("")}</row>`).join(""));
  });
  const wb = Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: xml, names: Object.keys(sheets) })));
  return { wb, calc: new Calc(wb) };
}

/** Evaluate one formula on Sheet1 at Z100 against `cells`. */
async function ev(formula: string, cells: Cells = {}) {
  const { calc, wb } = await book({ Sheet1: cells });
  return calc.evaluate(wb.sheets[0], 99, 25, formula);
}

describe("formula evaluation", () => {
  it("does arithmetic with precedence, percent, power and concatenation", async () => {
    expect(await ev("1+2*3")).toBe(7);
    expect(await ev("-2^2")).toBe(4);
    expect(await ev("2^3^2")).toBe(64);
    expect(await ev("50%*4")).toBe(2);
    expect(await ev('"a"&1&TRUE')).toBe("a1TRUE");
    expect(await ev("0.1+0.2")).toBe(0.3);
    expect(await ev("10/4")).toBe(2.5);
  });

  it("returns errors: division by zero, bad text, unknown names", async () => {
    expect(await ev("1/0")).toEqual({ error: "#DIV/0!" });
    expect(await ev('1+"x"')).toEqual({ error: "#VALUE!" });
    expect(await ev("NoSuchName+1")).toEqual({ error: "#NAME?" });
    expect(await ev("#N/A+1")).toEqual({ error: "#N/A" });
  });

  it("coerces blanks, booleans and numeric text like a spreadsheet", async () => {
    expect(await ev("A1+1")).toBe(1);
    expect(await ev('A1&"x"')).toBe("x");
    expect(await ev("TRUE+1")).toBe(2);
    expect(await ev('"3"*2')).toBe(6);
    expect(await ev("A1=0")).toBe(true);
    expect(await ev('A1=""')).toBe(true);
  });

  it("compares numbers, text (case-insensitive) and booleans", async () => {
    expect(await ev('"abc"="ABC"')).toBe(true);
    expect(await ev('"b">"a"')).toBe(true);
    expect(await ev('1<"a"')).toBe(true);
    expect(await ev('"z"<TRUE')).toBe(true);
    expect(await ev("2<>2")).toBe(false);
  });

  it("sums ranges, skipping text and booleans in ranges but not in arguments", async () => {
    const cells = { A1: 1, A2: "x", A3: true, A4: 4, B1: 10 };
    expect(await ev("SUM(A1:A4)", cells)).toBe(5);
    expect(await ev("SUM(A1:A4,TRUE,\"2\")", cells)).toBe(8);
    expect(await ev("SUM(A:A)", cells)).toBe(5);
    expect(await ev("SUM(A1:B4)", cells)).toBe(15);
    expect(await ev('SUM("x")', cells)).toEqual({ error: "#VALUE!" });
    expect(await ev("AVERAGE(A1:A4)", cells)).toBe(2.5);
    expect(await ev("COUNT(A1:A4)", cells)).toBe(2);
    expect(await ev("COUNTA(A1:A4)", cells)).toBe(4);
    expect(await ev("COUNTBLANK(A1:A5)", cells)).toBe(1);
    expect(await ev("MIN(A1:A4)", cells)).toBe(1);
    expect(await ev("MAX(A1:B4)", cells)).toBe(10);
    expect(await ev("AVERAGE(C1:C3)", cells)).toEqual({ error: "#DIV/0!" });
    expect(await ev("MAX(C1:C3)", cells)).toBe(0);
  });

  it("propagates errors from ranges", async () => {
    expect(await ev("SUM(A1:A2)", { A1: 1, A2: "=1/0" })).toEqual({ error: "#DIV/0!" });
  });

  it("evaluates IF and IFERROR lazily, plus AND / OR / NOT", async () => {
    expect(await ev('IF(1>2,"y","n")')).toBe("n");
    expect(await ev("IF(TRUE,1)")).toBe(1);
    expect(await ev("IF(FALSE,1)")).toBe(false);
    expect(await ev("IF(TRUE,1,1/0)")).toBe(1);
    expect(await ev("IFERROR(1/0,\"oops\")")).toBe("oops");
    expect(await ev("IFNA(#N/A,2)")).toBe(2);
    expect(await ev("AND(TRUE,1,A1:A2)", { A1: true })).toBe(true);
    expect(await ev("OR(FALSE,0)")).toBe(false);
    expect(await ev("NOT(0)")).toBe(true);
    expect(await ev('IFS(1>2,"a",2>1,"b")')).toBe("b");
    expect(await ev('SWITCH(2,1,"one",2,"two","other")')).toBe("two");
  });

  it("rounds like a spreadsheet", async () => {
    expect(await ev("ROUND(2.5,0)")).toBe(3);
    expect(await ev("ROUND(-2.5,0)")).toBe(-3);
    expect(await ev("ROUND(1.005,2)")).toBe(1.01);
    expect(await ev("ROUND(1234,-2)")).toBe(1200);
    expect(await ev("ROUNDUP(1.21,1)")).toBe(1.3);
    expect(await ev("ROUNDDOWN(-1.29,1)")).toBe(-1.2);
    expect(await ev("INT(-1.5)")).toBe(-2);
    expect(await ev("MOD(-3,2)")).toBe(1);
    expect(await ev("MOD(1,0)")).toEqual({ error: "#DIV/0!" });
    expect(await ev("ABS(-2)+SQRT(16)+POWER(2,3)")).toBe(14);
    expect(await ev("SQRT(-1)")).toEqual({ error: "#NUM!" });
    expect(await ev("PRODUCT(2,3,4)")).toBe(24);
  });

  it("handles text functions", async () => {
    expect(await ev('CONCATENATE("a",1,"b")')).toBe("a1b");
    expect(await ev("CONCAT(A1:A2)", { A1: "x", A2: 2 })).toBe("x2");
    expect(await ev('TEXTJOIN(", ",TRUE,A1:A3)', { A1: "a", A3: "c" })).toBe("a, c");
    expect(await ev('LEN("hello")')).toBe(5);
    expect(await ev('LEFT("hello",2)&RIGHT("hello")&MID("hello",2,3)')).toBe("heoell");
    expect(await ev('UPPER("a")&LOWER("B")&PROPER("hello world")')).toBe("AbHello World");
    expect(await ev('TRIM("  a   b ")')).toBe("a b");
    expect(await ev('SUBSTITUTE("a-b-c","-","+")')).toBe("a+b+c");
    expect(await ev('SUBSTITUTE("a-b-c","-","+",2)')).toBe("a-b+c");
    expect(await ev('FIND("b","abc")')).toBe(2);
    expect(await ev('FIND("B","abc")')).toEqual({ error: "#VALUE!" });
    expect(await ev('SEARCH("B","abc")')).toBe(2);
    expect(await ev('TEXT(0.5,"0%")')).toBe("50%");
    expect(await ev('VALUE("1,234")')).toBe(1234);
    expect(await ev('REPT("ab",2)')).toBe("abab");
    expect(await ev('EXACT("a","A")')).toBe(false);
    expect(await ev("1&2")).toBe("12");
    expect(await ev("1/3&\"\"")).toBe("0.333333333333333");
  });

  it("looks values up", async () => {
    const t = { A1: "apple", B1: 1, A2: "banana", B2: 2, A3: "cherry", B3: 3 };
    expect(await ev('VLOOKUP("banana",A1:B3,2,FALSE)', t)).toBe(2);
    expect(await ev('VLOOKUP("BANANA",A1:B3,2,0)', t)).toBe(2);
    expect(await ev('VLOOKUP("kiwi",A1:B3,2,FALSE)', t)).toEqual({ error: "#N/A" });
    expect(await ev("VLOOKUP(2.5,B1:B3,1,TRUE)", t)).toBe(2);
    expect(await ev("VLOOKUP(2.5,B1:B3,1)", t)).toBe(2);
    expect(await ev('VLOOKUP("a*",A1:B3,2,FALSE)', t)).toBe(1);
    expect(await ev('HLOOKUP(2,B1:B3,1,FALSE)', { B1: 2 })).toBe(2);
    expect(await ev('INDEX(A1:B3,2,2)', t)).toBe(2);
    expect(await ev('INDEX(A1:B3,0,2)', t)).toBe(1);
    expect(await ev('MATCH("cherry",A1:A3,0)', t)).toBe(3);
    expect(await ev("MATCH(2.5,B1:B3)", t)).toBe(2);
    expect(await ev('INDEX(B1:B3,MATCH("cherry",A1:A3,0))', t)).toBe(3);
    expect(await ev('XLOOKUP("cherry",A1:A3,B1:B3)', t)).toBe(3);
    expect(await ev('XLOOKUP("kiwi",A1:A3,B1:B3,"none")', t)).toBe("none");
    expect(await ev('CHOOSE(2,"a","b","c")')).toBe("b");
    expect(await ev("ROW(B3)+COLUMN(B3)+ROWS(A1:B3)+COLUMNS(A1:B3)")).toBe(10);
    expect(await ev("ROW()")).toBe(100);
  });

  it("does conditional aggregation with criteria strings", async () => {
    const t = { A1: "x", B1: 1, A2: "y", B2: 2, A3: "x", B3: 3, A4: "xy", B4: 4 };
    expect(await ev('SUMIF(A1:A4,"x",B1:B4)', t)).toBe(4);
    expect(await ev('SUMIF(B1:B4,">2")', t)).toBe(7);
    expect(await ev('SUMIF(A1:A4,"x*",B1:B4)', t)).toBe(8);
    expect(await ev('SUMIF(A1:A4,"<>x",B1:B4)', t)).toBe(6);
    expect(await ev('COUNTIF(A1:A4,"x")', t)).toBe(2);
    expect(await ev('COUNTIF(B1:B4,">=2")', t)).toBe(3);
    expect(await ev('COUNTIFS(A1:A4,"x",B1:B4,">1")', t)).toBe(1);
    expect(await ev('SUMIFS(B1:B4,A1:A4,"x*",B1:B4,"<4")', t)).toBe(4);
    expect(await ev('AVERAGEIF(A1:A4,"x",B1:B4)', t)).toBe(2);
    expect(await ev("SUMPRODUCT(B1:B2,B3:B4)", t)).toBe(11);
    expect(await ev("MEDIAN(B1:B4)", t)).toBe(2.5);
  });

  it("works with dates", async () => {
    expect(await ev("DATE(2024,1,1)")).toBe(45292);
    expect(await ev("DATE(2024,13,1)")).toBe(45658);
    expect(await ev("YEAR(45292)*10000+MONTH(45292)*100+DAY(45292)")).toBe(20240101);
    expect(await ev("WEEKDAY(45292)")).toBe(2);
    expect(await ev("EDATE(45322,1)")).toBe(45351);
    expect(await ev("EOMONTH(45292,1)")).toBe(45351);
    expect(await ev("HOUR(0.75)+MINUTE(0.75)+SECOND(0.75)")).toBe(18);
    expect(await ev("TIME(12,0,0)")).toBe(0.5);
    expect(await ev('DATEDIF(45292,45658,"y")')).toBe(1);
    expect(await ev('DATEDIF(45292,45351,"m")')).toBe(1);
    expect(await ev("DAYS(45300,45292)")).toBe(8);
  });

  it("checks types", async () => {
    expect(await ev("ISBLANK(A1)")).toBe(true);
    expect(await ev("ISNUMBER(1)+ISTEXT(\"a\")+ISLOGICAL(TRUE)+ISERROR(1/0)+ISNA(NA())")).toBe(5);
    expect(await ev("ISERR(NA())")).toBe(false);
  });

  it("reads other sheets and defined names", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 1 }, "My Data": { B2: 5 } });
    const names = wb.pkg.workbook.createElementNS(S, "definedNames");
    names.innerHTML = `<definedName xmlns="${S}" name="Rate">'My Data'!$B$2</definedName><definedName xmlns="${S}" name="Two">2</definedName>`;
    wb.pkg.workbook.documentElement.insertBefore(names, wb.pkg.workbook.getElementsByTagNameNS(S, "calcPr")[0]);
    expect(calc.evaluate(wb.sheets[0], 0, 5, "'My Data'!B2+Sheet1!A1")).toBe(6);
    expect(calc.evaluate(wb.sheets[0], 0, 5, "Rate*Two")).toBe(10);
    expect(calc.evaluate(wb.sheets[0], 0, 5, "Nope!A1")).toEqual({ error: "#REF!" });
  });

  it("reports unsupported functions instead of guessing", async () => {
    const { wb, calc } = await book({ Sheet1: {} });
    expect(calc.evaluate(wb.sheets[0], 0, 0, "WEBSERVICE(\"x\")")).toEqual(Calc.UNSUPPORTED);
    expect(calc.evaluate(wb.sheets[0], 0, 0, "Table1[Col]")).toEqual({ error: "#NAME?" });
  });
});

describe("recalculation", () => {
  const edit = (wb: Workbook, calc: Calc, sheet: number, a1: string, value: number | string) => {
    const r = parseCellRef(a1)!;
    const s = wb.sheets[sheet];
    if (typeof value === "string" && value.startsWith("=")) s.setCell(r.row, r.col, { formula: value.slice(1), value: null });
    else s.setCell(r.row, r.col, { value });
    return calc.recalc([{ sheet: s, row: r.row, col: r.col }]);
  };

  it("fills in formulas without cached results on open, without writing them", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 2, A2: "=A1*3" } });
    calc.computeMissing();
    expect(wb.sheets[0].cell("A2")!.value).toBe(6);
    expect(wb.sheets[0].cell("A2")!.el.getElementsByTagNameNS(S, "v")).toHaveLength(0);
  });

  it("updates dependents in order, across sheets, and writes their caches", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 1, A2: "=A1+1", A3: "=A2*10" }, Sheet2: { A1: "=SUM(Sheet1!A1:A3)" } });
    calc.computeMissing();
    const changed = edit(wb, calc, 0, "A1", 5);
    expect(wb.sheets[0].cell("A2")!.value).toBe(6);
    expect(wb.sheets[0].cell("A3")!.value).toBe(60);
    expect(wb.sheets[1].cell("A1")!.value).toBe(71);
    expect(changed.map((c) => c.el.getAttribute("r")).sort()).toEqual(["A1", "A2", "A3"]);
    expect(wb.sheets[1].cell("A1")!.el.getElementsByTagNameNS(S, "v")[0].textContent).toBe("71");
  });

  it("calculates a newly entered formula", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 4 } });
    edit(wb, calc, 0, "B1", "=A1^2");
    expect(wb.sheets[0].cell("B1")!.value).toBe(16);
  });

  it("recalculates whole-column references without expanding them", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 1, B1: "=SUM(A:A)" } });
    calc.computeMissing();
    edit(wb, calc, 0, "A100000", 2);
    expect(wb.sheets[0].cell("B1")!.value).toBe(3);
  });

  it("reports circular references as errors instead of hanging", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: "=B1+1", B1: "=A1+1" } });
    calc.computeMissing();
    expect(wb.sheets[0].cell("A1")!.value).toEqual({ error: "#REF!" });
    expect(calc.isCircular(wb.sheets[0].cell("A1")!)).toBe(true);
    edit(wb, calc, 0, "B1", 1);
    expect(wb.sheets[0].cell("A1")!.value).toBe(2);
    expect(calc.isCircular(wb.sheets[0].cell("A1")!)).toBe(false);
  });

  it("keeps an unsupported formula's cache and marks it and its dependents stale when inputs change", async () => {
    const { wb } = await book({ Sheet1: { A1: 1 } });
    const s = wb.sheets[0];
    // As if the file had come with these formulas and their cached results.
    s.setCell(0, 1, { formula: "WEBSERVICE(A1)", value: "cached" });
    s.setCell(0, 2, { formula: 'B1&"!"', value: "cached!" });
    const calc = new Calc(wb);
    expect(calc.isStale(s.cell("B1")!)).toBe(false);
    edit(wb, calc, 0, "A1", 2);
    expect(s.cell("B1")!.value).toBe("cached");
    expect(calc.isStale(s.cell("B1")!)).toBe(true);
    expect(calc.isStale(s.cell("C1")!)).toBe(true);
    // Excel is asked to recalculate the file when it opens, since some results here couldn't be.
    expect(wb.pkg.workbook.getElementsByTagNameNS(S, "calcPr")[0].getAttribute("fullCalcOnLoad")).toBe("1");
  });

  it("drops the calculation chain when formulas change", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: 1 } });
    edit(wb, calc, 0, "B1", "=A1");
    expect(wb.pkg.rels(wb.pkg.workbookPath).some((r) => r.type === "calcChain")).toBe(false);
    expect(wb.pkg.xml("xl/calcChain.xml")).toBeNull();
  });

  it("recalculates TODAY and other volatile functions on every edit", async () => {
    const { wb, calc } = await book({ Sheet1: { A1: "=TODAY()", B1: 1 } });
    calc.computeMissing();
    wb.sheets[0].cell("A1")!.value = 0;
    edit(wb, calc, 0, "B1", 2);
    expect(wb.sheets[0].cell("A1")!.value).toBeGreaterThan(45000);
  });
});
