import { describe, expect, it } from "vitest";
import { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { markdownToDoc, pasteMarkdown, pasteSlice } from "../src/domain/docx/from-markdown";

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
