import { describe, expect, it } from "vitest";
import { DEFAULT_SECTPR, lineNumberLabels, lineNumbersOf, LineNumbers, withLineNumbers } from "../src/domain/docx/sections";
import { W } from "./helpers";

const ln = (o: Partial<LineNumbers> = {}): LineNumbers => ({ start: 1, countBy: 1, restart: "newPage", ...o });
const lines = (spec: [page: number, section: number][], set: (section: number) => LineNumbers | null) => spec.map(([page, section]) => ({ page, section, ln: set(section) }));

describe("w:lnNumType", () => {
  it("absent: no numbering", () => {
    expect(lineNumbersOf(DEFAULT_SECTPR)).toBeNull();
    expect(lineNumbersOf(null)).toBeNull();
  });

  it("Word's defaults: restart each page, every line, from 1", () => {
    expect(lineNumbersOf(`<w:sectPr xmlns:w="${W}"><w:lnNumType w:countBy="1"/></w:sectPr>`)).toEqual(ln());
  });

  it("w:start is one less than the first number shown", () => {
    expect(lineNumbersOf(`<w:sectPr xmlns:w="${W}"><w:lnNumType w:countBy="5" w:start="9" w:restart="continuous"/></w:sectPr>`)).toEqual(ln({ start: 10, countBy: 5, restart: "continuous" }));
  });

  it("round-trips and lands in schema order (after pgMar, before cols)", () => {
    const sectPr = `<w:sectPr xmlns:w="${W}"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440"/><w:cols w:num="2"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    const out = withLineNumbers(sectPr, ln({ restart: "newSection", start: 3 }));
    expect(out.match(/<w:(\w+)/g)!.map((t) => t.slice(3))).toEqual(["sectPr", "pgSz", "pgMar", "lnNumType", "cols", "docGrid"]);
    expect(lineNumbersOf(out)).toEqual(ln({ restart: "newSection", start: 3 }));
  });

  it("replacing keeps one element; null removes it, and leaves a section without one untouched", () => {
    const once = withLineNumbers(DEFAULT_SECTPR, ln());
    const twice = withLineNumbers(once, ln({ restart: "continuous" }));
    expect(twice.match(/lnNumType/g)).toHaveLength(1);
    expect(lineNumbersOf(withLineNumbers(twice, null))).toBeNull();
    expect(withLineNumbers(DEFAULT_SECTPR, null)).toBe(DEFAULT_SECTPR);
  });
});

describe("lineNumberLabels", () => {
  it("restart each page", () => {
    expect(lineNumberLabels(lines([[1, 0], [1, 0], [1, 0], [2, 0], [2, 0]], () => ln()))).toEqual([1, 2, 3, 1, 2]);
  });

  it("continuous runs across pages and sections", () => {
    expect(lineNumberLabels(lines([[1, 0], [2, 0], [2, 1], [3, 1]], () => ln({ restart: "continuous" })))).toEqual([1, 2, 3, 4]);
  });

  it("restart each section ignores pages", () => {
    expect(lineNumberLabels(lines([[1, 0], [2, 0], [2, 1], [3, 1]], () => ln({ restart: "newSection" })))).toEqual([1, 2, 1, 2]);
  });

  it("count by shows only multiples; start shifts the count", () => {
    expect(lineNumberLabels(lines([[1, 0], [1, 0], [1, 0], [1, 0], [1, 0]], () => ln({ countBy: 2 })))).toEqual([null, 2, null, 4, null]);
    expect(lineNumberLabels(lines([[1, 0], [1, 0]], () => ln({ start: 10 })))).toEqual([10, 11]);
  });

  it("a section without numbering is skipped and not counted; continuous carries over it", () => {
    const set = (s: number) => (s === 1 ? null : ln({ restart: "continuous" }));
    expect(lineNumberLabels(lines([[1, 0], [1, 0], [1, 1], [1, 2]], set))).toEqual([1, 2, null, 3]);
  });

  it("a continuous section after one that restarts each page goes on counting", () => {
    const set = (s: number) => ln({ restart: s === 0 ? "newPage" : "continuous" });
    expect(lineNumberLabels(lines([[1, 0], [2, 0], [2, 1]], set))).toEqual([1, 1, 2]);
  });
});
