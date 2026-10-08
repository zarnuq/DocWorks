import { describe, it, expect } from "vitest";
import { History } from "../src/core/utils/history";

describe("History", () => {
  it("undoes and redoes in order", () => {
    const h = new History<number>();
    h.push(1);
    h.push(2);
    expect(h.undo(3)).toBe(2);
    expect(h.undo(2)).toBe(1);
    expect(h.undo(1)).toBeNull();
    expect(h.redo(1)).toBe(2);
    expect(h.redo(2)).toBe(3);
    expect(h.redo(3)).toBeNull();
  });

  it("clears redo on a new change", () => {
    const h = new History<string>();
    h.push("a");
    h.undo("b");
    h.push("a2");
    expect(h.redo("c")).toBeNull();
  });

  it("caps its size", () => {
    const h = new History<number>(2);
    [1, 2, 3].forEach((n) => h.push(n));
    expect(h.undo(4)).toBe(3);
    expect(h.undo(3)).toBe(2);
    expect(h.undo(2)).toBeNull();
  });
});
