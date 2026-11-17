import { describe, expect, it } from "vitest";
import { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { markdownToDoc, pasteMarkdown, pasteSlice } from "../src/domain/docx/from-markdown";
import { markdownImages } from "../src/core/utils/markdown";

const ctx = { list: (kind: "bullet" | "ordered" | "task") => (kind === "bullet" ? "1" : kind === "task" ? "3" : "2"), image: () => null };
const nodes = (doc: PMNode) => {
  const out: PMNode[] = [];
  doc.forEach((n) => out.push(n));
  return out;
};

describe("Markdown pasted into a document", () => {
  it("makes headings, quotes, nested lists, tasks and rules", () => {
    const doc = markdownToDoc("# Title\nPlain\n> quoted\n- a\n  - b\n1. one\n- [x] done\n---", ctx);
    expect(nodes(doc).map((n) => [n.type.name, n.textContent, n.attrs.style, n.attrs.numId, n.attrs.ilvl])).toEqual([
      ["paragraph", "Title", "Heading1", null, null],
      ["paragraph", "Plain", null, null, null],
      ["paragraph", "quoted", "Quote", null, null],
      ["paragraph", "a", null, "1", 0],
      ["paragraph", "b", null, "1", 1],
      ["paragraph", "one", null, "2", 0],
      ["paragraph", "☑ done", null, "3", 0],
      ["raw_block", "", undefined, undefined, undefined],
    ]);
  });

  it("keeps inline formatting as marks", () => {
    const p = markdownToDoc("a **b** *c* ~~d~~ ==e== `f` [g](https://x.y)", ctx).firstChild!;
    const marks: Record<string, string[]> = {};
    p.forEach((t) => (marks[t.text!] = t.marks.map((m) => m.type.name)));
    expect(marks).toMatchObject({ b: ["bold"], c: ["italic"], d: ["strike"], e: ["highlight"], f: ["font"], g: ["link"] });
  });

  it("makes code boxes and bordered tables", () => {
    const [code, table] = nodes(markdownToDoc("```js\nlet a;\nlet b;\n```\n| H | I |\n|---|---|\n| 1 | 2 |", ctx));
    expect(code.type.name).toBe("sdt");
    expect(code.firstChild!.type.name).toBe("table");
    expect(code.textContent).toBe("let a;let b;");
    expect(table.type.name).toBe("table");
    expect(table.childCount).toBe(2);
    expect(table.firstChild!.attrs.header).toBe(true);
    expect(table.textContent).toBe("HI12");
  });

  it("opens the slice into the paragraph at the caret only around plain paragraphs", () => {
    expect([pasteSlice(markdownToDoc("a **b**", ctx)).openStart, pasteSlice(markdownToDoc("a **b**", ctx)).openEnd]).toEqual([1, 1]);
    const s = pasteSlice(markdownToDoc("# H\ntext", ctx));
    expect([s.openStart, s.openEnd]).toEqual([0, 1]);
  });
});

describe("math and pictures pasted into a document", () => {
  it("makes editable equations: inline in the text, display ones centred on their own line", () => {
    const [p, block] = nodes(markdownToDoc("Mass $E = mc^2$ here\n$$\n\\frac{a}{b}\n$$", ctx));
    const kinds: [string, string][] = [];
    p.forEach((n) => kinds.push([n.type.name, n.type.name === "opaque" ? String(n.attrs.text) : n.text!]));
    expect(kinds).toEqual([["text", "Mass "], ["opaque", "E = mc^2"], ["text", " here"]]);
    expect(String(p.child(1).attrs.xml)).toContain("oMath");
    expect([block.attrs.align, block.childCount, block.firstChild!.attrs.text]).toEqual(["center", 1, "\\frac{a}{b}"]);
    expect(String(block.firstChild!.attrs.xml)).toContain("oMathPara");
  });

  it("puts vault pictures in the text through the editor's loader, and names the ones it can't load", () => {
    const asked: string[] = [];
    const image = (img: HTMLImageElement) => {
      asked.push(img.getAttribute("src")!);
      return img.getAttribute("src") === "ridge.png" ? { rId: "rId9", src: "blob:x", width: 40, height: 30 } : null;
    };
    const [p] = nodes(markdownToDoc("A ![[ridge.png]] and ![gone](missing.png)", { ...ctx, image }));
    expect(asked).toEqual(["ridge.png", "missing.png"]);
    const kinds: string[] = [];
    p.forEach((n) => kinds.push(n.type.name === "image" ? `image:${n.attrs.rId}:${n.attrs.alt}` : n.text!));
    expect(kinds).toEqual(["A ", "image:rId9:ridge.png", " and gone"]);
    expect(markdownImages("A ![[ridge.png]]\n- ![x](a/b.jpg) ![[ridge.png]]\n| ![[t.gif]] |\n|---|")).toEqual(["ridge.png", "a/b.jpg", "t.gif"]);
  });
});

describe("loose numbered lists", () => {
  it("keeps items in one list even with blank lines between them", () => {
    const doc = markdownToDoc("1. first\n\n2. second\n\n3. third", ctx);
    const items = nodes(doc).filter((n) => n.attrs.numId);
    expect(items).toHaveLength(3);
    expect(items.every((n) => n.attrs.numId === items[0].attrs.numId)).toBe(true);
  });
});

describe("callouts pasted into a document", () => {
  it("become the callout boxes a note export makes: type and folding kept, title first, body as quotes", () => {
    const [box, after, plain] = nodes(markdownToDoc("> [!warning]- Mind the **gap**\n> first line\n> second\nafter\n> just a quote", ctx));
    expect(box.type.name).toBe("sdt");
    expect(String(box.attrs.xml)).toContain('w:val="doc-works:callout:warning:-"');
    const inside: [string, string | null][] = [];
    box.forEach((p) => inside.push([p.textContent, p.attrs.style]));
    expect(inside).toEqual([["Mind the gap", "Quote"], ["first line", "Quote"], ["second", "Quote"]]);
    expect([after.textContent, after.attrs.style]).toEqual(["after", null]);
    expect([plain.type.name, plain.textContent, plain.attrs.style]).toEqual(["paragraph", "just a quote", "Quote"]);
    // No title: the type's name, as Obsidian shows it.
    const [tip] = nodes(markdownToDoc("> [!tip]\n> body", ctx));
    expect(String(tip.attrs.xml)).toContain('w:val="doc-works:callout:tip:"');
    expect([tip.firstChild!.textContent, tip.childCount]).toEqual(["Tip", 2]);
  });
});

describe("caret after a Markdown paste", () => {
  // Word and Google Docs leave the caret at the end of what was pasted, so typing (or pasting again) continues there.
  const start = () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("See")),
      schema.nodes.paragraph.create(),
      schema.nodes.paragraph.create(null, schema.text("First bullet")),
    ]);
    const state = EditorState.create({ doc, schema });
    return state.apply(state.tr.setSelection(TextSelection.create(doc, 6))); // in the empty paragraph
  };
  const caret = (md: string) => {
    const state = start();
    const $from = state.apply(pasteMarkdown(state, markdownToDoc(md, ctx))).selection.$from;
    return [$from.parent.textContent, $from.parentOffset];
  };

  it("ends after a pasted list, not at the start of the next paragraph", () => {
    expect(caret("## Pasted heading\n\n- item one\n- item two")).toEqual(["item two", 8]);
  });

  it("ends after pasted plain text", () => {
    expect(caret("**plain** text")).toEqual(["plain text", 10]);
  });
});
