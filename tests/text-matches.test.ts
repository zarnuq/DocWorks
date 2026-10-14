import { describe, expect, it } from "vitest";
import { expandReplacement, textMatches } from "../src/core/utils/text-matches";

const spans = (text: string, query: string, o: Parameters<typeof textMatches>[2]) => textMatches(text, query, o).map((m) => text.slice(m.start, m.end));

describe("textMatches options", () => {
  it("matches regular expressions, and finds nothing for an unfinished one", () => {
    expect(spans("cat cot cut", "c[ao]t", { regex: true })).toEqual(["cat", "cot"]);
    expect(spans("a.b axb", "a.b", {})).toEqual(["a.b"]);
    expect(textMatches("abc", "(", { regex: true })).toEqual([]);
    // A pattern that can match nothing doesn't loop or report empty matches.
    expect(textMatches("abc", "x*", { regex: true })).toEqual([]);
  });

  it("matches whole words only", () => {
    expect(spans("cat concat cat_s cat.", "cat", { wholeWord: true })).toEqual(["cat", "cat"]);
    expect(spans("Über über", "über", { wholeWord: true })).toEqual(["Über", "über"]);
  });

  it("ignores accents, keeping offsets in the original text", () => {
    expect(spans("Café cafe CAFÉ", "cafe", { ignoreAccents: true })).toEqual(["Café", "cafe", "CAFÉ"]);
    expect(spans("Café", "café", { ignoreAccents: true })).toEqual(["Café"]);
    // Decomposed text: e + combining acute is covered whole.
    expect(spans("café x", "cafe", { ignoreAccents: true })).toEqual(["café"]);
    expect(spans("Café", "cafe", {})).toEqual([]);
  });

  it("keeps the old matchCase / entire arguments", () => {
    expect(textMatches("Needle needle", "needle", true)).toEqual([{ start: 7, end: 13 }]);
    expect(textMatches("needle", "needle", false, true)).toHaveLength(1);
    expect(textMatches("needles", "needle", { entire: true })).toHaveLength(0);
  });
});

describe("expandReplacement", () => {
  it("fills $1, $& and $$ from a regular expression match", () => {
    const [m] = textMatches("2026-10-14", "(\\d+)-(\\d+)-(\\d+)", { regex: true });
    expect(expandReplacement("$3/$2/$1 ($&) $$5", m)).toBe("14/10/2026 (2026-10-14) $5");
  });

  it("leaves a plain search's replacement as typed", () => {
    const [m] = textMatches("price", "price", {});
    expect(expandReplacement("$1 off", m)).toBe("$1 off");
  });
});
