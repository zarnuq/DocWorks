import { describe, expect, it } from "vitest";
import { diff, diffWords } from "../src/core/utils/diff";
import { compareText } from "../src/infrastructure/office-text";

const p = (text: string, label = "") => ({ label, text });

describe("diff", () => {
  it("keeps the longest common run and marks the rest removed or added", () => {
    expect(diff(["a", "b", "c", "d"], ["a", "x", "c", "d", "e"])).toEqual([
      { kind: "same", a: "a", b: "a" },
      { kind: "removed", a: "b" },
      { kind: "added", b: "x" },
      { kind: "same", a: "c", b: "c" },
      { kind: "same", a: "d", b: "d" },
      { kind: "added", b: "e" },
    ]);
  });

  it("marks words removed and added inside a changed paragraph", () => {
    expect(diffWords("The quick brown fox", "The slow brown fox jumps")).toEqual([
      { kind: "same", text: "The " },
      { kind: "removed", text: "quick" },
      { kind: "added", text: "slow" },
      { kind: "same", text: " brown fox" },
      { kind: "added", text: " jumps" },
    ]);
  });
});

describe("comparing a kept version with the file now", () => {
  it("documents: paragraphs added, removed and changed, in order", () => {
    const before = [p("Title"), p("Old intro"), p("Gone"), p("End")];
    const after = [p("Title"), p("New intro"), p("End"), p("Added at the end")];
    expect(compareText("docx", before, after)).toEqual([
      { kind: "changed", where: "", before: "Old intro", after: "New intro" },
      { kind: "removed", where: "", before: "Gone" },
      { kind: "added", where: "", after: "Added at the end" },
    ]);
  });

  it("slides: changes say which slide they're on", () => {
    const before = [p("Hello", "Slide 1"), p("Point one", "Slide 2")];
    const after = [p("Hello", "Slide 1"), p("Point 1", "Slide 2"), p("Thanks", "Slide 3")];
    expect(compareText("pptx", before, after)).toEqual([
      { kind: "changed", where: "Slide 2", before: "Point one", after: "Point 1" },
      { kind: "added", where: "Slide 3", after: "Thanks" },
    ]);
  });

  it("sheets: cells compared by address, in sheet then row then column order", () => {
    const before = [p("120", "Budget!B4"), p("x", "Budget!A10"), p("old", "Notes!A1")];
    const after = [p("150", "Budget!B4"), p("x", "Budget!A10"), p("new", "Budget!C2")];
    expect(compareText("xlsx", before, after)).toEqual([
      { kind: "added", where: "Budget!C2", after: "new" },
      { kind: "changed", where: "Budget!B4", before: "120", after: "150" },
      { kind: "removed", where: "Notes!A1", before: "old" },
    ]);
  });

  it("nothing changed: no changes", () => {
    expect(compareText("docx", [p("a")], [p("a")])).toEqual([]);
  });

  it("a very long document still compares (the middle is matched by text when too big to align)", () => {
    const before = Array.from({ length: 6000 }, (_, i) => p(`para ${i}`));
    const after = before.map((x, i) => (i === 3000 ? p("changed") : x));
    const changes = compareText("docx", before, after);
    expect(changes).toEqual([{ kind: "changed", where: "", before: "para 3000", after: "changed" }]);
  });
});
