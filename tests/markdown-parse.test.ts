import { describe, expect, it } from "vitest";
import { looksLikeMarkdown, markdownToPlain, parseInline, parseMarkdown } from "../src/core/utils/markdown";

describe("parseMarkdown", () => {
  it("reads headings, paragraphs, quotes, rules and skips blank lines", () => {
    const blocks = parseMarkdown("# Title\n\nSome text\n> quoted\n---\n###### Six\n#tag not heading");
    expect(blocks.map((b) => b.kind)).toEqual(["heading", "blank", "para", "quote", "hr", "heading", "para"]);
    expect(blocks[0]).toMatchObject({ kind: "heading", level: 1, runs: [{ text: "Title" }] });
    expect(blocks[5]).toMatchObject({ level: 6 });
    expect(blocks[6]).toMatchObject({ runs: [{ text: "#tag not heading" }] });
  });

  it("reads nested bullet, numbered and task lists", () => {
    const blocks = parseMarkdown("- a\n  - b\n\t\t- c\n3. three\n- [ ] todo\n- [x] done\n* star\n+ plus");
    expect(blocks).toMatchObject([
      { kind: "item", list: "bullet", depth: 0, runs: [{ text: "a" }] },
      { kind: "item", list: "bullet", depth: 1, runs: [{ text: "b" }] },
      { kind: "item", list: "bullet", depth: 2, runs: [{ text: "c" }] },
      { kind: "item", list: "ordered", depth: 0, number: 3, runs: [{ text: "three" }] },
      { kind: "item", list: "task", checked: false, runs: [{ text: "todo" }] },
      { kind: "item", list: "task", checked: true, runs: [{ text: "done" }] },
      { kind: "item", list: "bullet", runs: [{ text: "star" }] },
      { kind: "item", list: "bullet", runs: [{ text: "plus" }] },
    ]);
  });

  it("keeps fenced code verbatim, blank lines included", () => {
    const blocks = parseMarkdown("```ts\nconst a = **1**;\n\nlet b;\n```\nafter");
    expect(blocks[0]).toEqual({ kind: "code", lang: "ts", text: "const a = **1**;\n\nlet b;" });
    expect(blocks[1]).toMatchObject({ kind: "para" });
  });

  it("reads pipe tables without the delimiter row", () => {
    const blocks = parseMarkdown("| A | **B** |\n|---|:-:|\n| 1 | 2 \\| 3 |");
    expect(blocks).toHaveLength(1);
    const t = blocks[0] as Extract<ReturnType<typeof parseMarkdown>[number], { kind: "table" }>;
    expect(t.rows.map((r) => r.map((c) => c.map((x) => x.text).join("")))).toEqual([["A", "B"], ["1", "2 | 3"]]);
    expect(t.rows[0][1][0].b).toBe(true);
  });
});

describe("parseInline", () => {
  it("reads emphasis, strike, highlight, code and links", () => {
    expect(parseInline("a **b** *c* _d_ ~~e~~ ==f== `g*h*` [i](http://x.y) ***j***")).toEqual([
      { text: "a " }, { text: "b", b: true }, { text: " " }, { text: "c", i: true }, { text: " " }, { text: "d", i: true }, { text: " " },
      { text: "e", strike: true }, { text: " " }, { text: "f", highlight: true }, { text: " " }, { text: "g*h*", code: true }, { text: " " },
      { text: "i", href: "http://x.y" }, { text: " " }, { text: "j", b: true, i: true },
    ]);
  });

  it("turns wikilinks and embeds into their text, keeps unmatched and intraword markers", () => {
    expect(parseInline("see [[Note#Head|alias]] and [[Other]] ![[Sub note]]")).toEqual([{ text: "see alias and Other Sub note" }]);
    expect(parseInline("snake_case_name 2*3 a ** b \\*lit\\*")).toEqual([{ text: "snake_case_name 2*3 a ** b *lit*" }]);
  });
});

describe("looksLikeMarkdown", () => {
  it("spots block and inline syntax, not ordinary text", () => {
    for (const md of ["# Head", "- item", "1. one", "> q", "```\nx\n```", "| a | b |\n|---|---|", "a **bold** word", "[l](http://a)", "[[Note]]", "- [ ] t"]) {
      expect(looksLikeMarkdown(md), md).toBe(true);
    }
    for (const text of ["Hello world", "Price: 3 * 4 = 12", "snake_case", "#hashtag", "a-b - c"]) {
      expect(looksLikeMarkdown(text), text).toBe(false);
    }
  });
});

describe("markdownToPlain", () => {
  it("drops heading, bullet, task, quote and inline syntax but keeps numbers and line structure", () => {
    const md = "## Plan\n\n- **first** item\n  - [x] done [[Note|link]]\n2. [two](http://a)\n> quote ~~old~~\n---\n```\nraw **code**\n```\n| a | b |\n|---|---|\n| 1 | 2 |";
    expect(markdownToPlain(md)).toBe("Plan\n\nfirst item\ndone link\n2. two\nquote old\n\nraw **code**\na\tb\n1\t2");
  });
});

describe("math and pictures", () => {
  it("reads inline math, not prices, and display math as its own block", () => {
    expect(parseInline("Mass $E = mc^2$ and $a_1$.")).toEqual([{ text: "Mass " }, { text: "E = mc^2", math: true }, { text: " and " }, { text: "a_1", math: true }, { text: "." }]);
    expect(parseInline("It costs $5 and $10, or $ 3 $.")).toEqual([{ text: "It costs $5 and $10, or $ 3 $." }]);
    expect(parseInline("inline $$x^2$$ too")).toEqual([{ text: "inline " }, { text: "x^2", math: true }, { text: " too" }]);
    expect(parseMarkdown("Before\n$$\n\\frac{a}{b}\n= c\n$$\nAfter\n$$x+y$$")).toEqual([
      { kind: "para", runs: [{ text: "Before" }] },
      { kind: "math", latex: "\\frac{a}{b}\n= c" },
      { kind: "para", runs: [{ text: "After" }] },
      { kind: "math", latex: "x+y" },
    ]);
  });

  it("reads vault pictures (embeds and relative links); web pictures and other embeds stay text", () => {
    expect(parseInline("See ![[Ridge view.png|300]] here")).toEqual([{ text: "See " }, { text: "Ridge view.png", image: "Ridge view.png" }, { text: " here" }]);
    expect(parseInline("![A map](attachments/map%20one.JPG)")).toEqual([{ text: "A map", image: "attachments/map one.JPG" }]);
    expect(parseInline("![logo](https://example.com/logo.png) ![[Note]] ![[deck.pptx#2]]")).toEqual([{ text: "logo Note deck.pptx#2" }]);
  });

  it("treats unclosed $$ as a paragraph, not a math block that swallows remaining lines", () => {
    const blocks = parseMarkdown("$$\nunclosed\nmore\ntext");
    expect(blocks.every((b) => b.kind !== "math")).toBe(true);
    expect(blocks.map((b) => b.kind)).toContain("para");
    // A properly closed one still works.
    const closed = parseMarkdown("$$\nok\n$$\nafter");
    expect(closed[0]).toMatchObject({ kind: "math", latex: "ok" });
    expect(closed[1]).toMatchObject({ kind: "para", runs: [{ text: "after" }] });
  });

  it("sees math as Markdown and keeps its source in plain text", () => {
    expect(looksLikeMarkdown("so $x^2$ grows")).toBe(true);
    expect(looksLikeMarkdown("$$\nx\n$$")).toBe(true);
    expect(looksLikeMarkdown("costs $5 and $10")).toBe(false);
    expect(looksLikeMarkdown("a fraction $\\frac{a}{b}$ here")).toBe(true);
    // Shell and template variables aren't math: dollars alone don't make text Markdown.
    for (const plain of ["export PATH=$HOME/bin:$PATH", "$HOME/$USER", "${A}${B}", "echo $MY_VAR and $OTHER_VAR"]) expect(looksLikeMarkdown(plain)).toBe(false);
    expect(markdownToPlain("a $x^2$ b\n$$\ny\n$$\n![[pic.png]]")).toBe("a x^2 b\ny\npic.png");
  });
});
