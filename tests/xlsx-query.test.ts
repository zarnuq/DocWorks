import { describe, expect, it } from "vitest";
import { runQuery } from "../src/domain/xlsx/query";

const DATA = [
  ["Region", "Product", "Amount"],
  ["East", "A", 10],
  ["East", "B", 20],
  ["West", "A", 5],
  ["West", "B", 15],
  ["North", "B", 3],
];

describe("QUERY", () => {
  it("selects columns by letter, keeping the header row", () => {
    expect(runQuery(DATA, "select A, C", 1)).toEqual([["Region", "Amount"], ["East", 10], ["East", 20], ["West", 5], ["West", 15], ["North", 3]]);
    expect(runQuery(DATA, "", 1)).toEqual(DATA);
  });

  it("filters with where: comparisons, text tests and and / or / not", () => {
    expect(runQuery(DATA, "select A, C where C > 9", 1)).toEqual([["Region", "Amount"], ["East", 10], ["East", 20], ["West", 15]]);
    expect(runQuery(DATA, "select * where A = 'West' and not B = 'A'", 1)).toEqual([["Region", "Product", "Amount"], ["West", "B", 15]]);
    expect(runQuery(DATA, "select A where A contains 'or' or C < 6", 1)).toEqual([["Region"], ["West"], ["North"]]);
    expect(runQuery(DATA, "select A where A starts with 'Ea' and (C >= 20 or B != 'B')", 1)).toEqual([["Region"], ["East"], ["East"]]);
  });

  it("groups with aggregates, orders, limits and labels", () => {
    expect(runQuery(DATA, "select A, sum(C) group by A order by sum(C) desc", 1)).toEqual([["Region", "sum Amount"], ["East", 30], ["West", 20], ["North", 3]]);
    expect(runQuery(DATA, "select B, count(A), avg(C) group by B label count(A) 'Rows', avg(C) 'Mean'", 1)).toEqual([["Product", "Rows", "Mean"], ["A", 2, 7.5], ["B", 3, 38 / 3]]);
    expect(runQuery(DATA, "select A, C order by C limit 2 offset 1", 1)).toEqual([["Region", "Amount"], ["West", 5], ["East", 10]]);
    expect(runQuery(DATA, "select max(C), min(C)", 1)).toEqual([["max Amount", "min Amount"], [20, 3]]);
  });

  it("works without a header row, naming columns Col1... for arrays", () => {
    const rows = DATA.slice(1);
    expect(runQuery(rows, "select Col1 where Col3 > 12", 0, true)).toEqual([["East"], ["West"]]);
    expect(runQuery(rows, "select A where C > 12", 0)).toEqual([["East"], ["West"]]);
  });

  it("rejects what it can't read", () => {
    expect(() => runQuery(DATA, "select Z", 1)).toThrow();
    expect(() => runQuery(DATA, "pivot A", 1)).toThrow();
  });
});
