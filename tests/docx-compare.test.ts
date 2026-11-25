import { describe, expect, it } from "vitest";
import { compareBodies, comparisonName } from "../src/domain/docx/compare";
import { acceptRevisions, rejectRevisions, revisions } from "../src/domain/docx/revisions";
import { W } from "./helpers";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPTS = { author: "Ada", date: "2026-11-21T00:00:00Z" };

function body(inner: string): Element {
  return new DOMParser().parseFromString(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${inner}<w:sectPr/></w:body></w:document>`, "application/xml").getElementsByTagNameNS(W, "body")[0];
}
const p = (text: string, rPr = "", pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const table = (...rows: string[][]) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${rows.map((r) => `<w:tr>${r.map((c) => `<w:tc>${p(c)}</w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;

/** Paragraph texts of a body as it reads (deleted text left out). */
function texts(b: Element): string[] {
  return Array.from(b.getElementsByTagNameNS(W, "p")).map((para) => Array.from(para.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join(""));
}
const clone = (b: Element) => b.cloneNode(true) as Element;

/** Compare, then check that accepting gives `other`'s text and rejecting gives `base`'s. */
function check(baseXml: string, otherXml: string) {
  const base = body(baseXml), other = body(otherXml);
  const want = { base: texts(base), other: texts(other) };
  const result = compareBodies(base, other, OPTS);
  const accepted = clone(base), rejected = clone(base);
  acceptRevisions(accepted);
  rejectRevisions(rejected);
  expect(texts(accepted), "accept all").toEqual(want.other);
  expect(texts(rejected), "reject all").toEqual(want.base);
  return { base, result };
}

describe("compareBodies", () => {
  it("identical documents: no changes, nothing touched", () => {
    const xml = p("Hello world") + p("Second");
    const { base, result } = check(xml, xml);
    expect(result.changes).toBe(0);
    expect(revisions(base)).toEqual([]);
  });

  it("a changed word is a deletion and an insertion inside its paragraph", () => {
    const { base, result } = check(p("The quick brown fox"), p("The slow brown fox"));
    expect(result.changes).toBe(2);
    const del = base.getElementsByTagNameNS(W, "del")[0];
    const ins = base.getElementsByTagNameNS(W, "ins")[0];
    expect(del.getElementsByTagNameNS(W, "delText")[0].textContent).toBe("quick");
    expect(ins.textContent).toBe("slow");
    for (const el of [del, ins]) {
      expect(el.getAttributeNS(W, "author")).toBe("Ada");
      expect(el.getAttributeNS(W, "date")).toBe(OPTS.date);
    }
  });

  it("kept text keeps the base formatting; inserted text takes the other document's", () => {
    const { base } = check(p("Plain words", "<w:b/>"), p("Plain more words", "<w:i/>"));
    const runs = Array.from(base.getElementsByTagNameNS(W, "r"));
    const same = runs.find((r) => r.textContent === "Plain ")!;
    expect(same.getElementsByTagNameNS(W, "b")).toHaveLength(1);
    const added = base.getElementsByTagNameNS(W, "ins")[0].getElementsByTagNameNS(W, "r")[0];
    expect(added.getElementsByTagNameNS(W, "i")).toHaveLength(1);
  });

  it("keeps the base paragraph's properties", () => {
    const { base } = check(p("Title here", "", `<w:pStyle w:val="Heading1"/>`), p("Title there"));
    expect(base.getElementsByTagNameNS(W, "pStyle")[0].getAttributeNS(W, "val")).toBe("Heading1");
  });

  it("an added paragraph is inserted with an inserted paragraph mark", () => {
    const { base, result } = check(p("One") + p("Three"), p("One") + p("Two") + p("Three"));
    expect(result.changes).toBeGreaterThan(0);
    const added = Array.from(base.getElementsByTagNameNS(W, "p"))[1];
    expect(added.getElementsByTagNameNS(W, "pPr")[0].getElementsByTagNameNS(W, "ins")).toHaveLength(1);
  });

  it("a removed paragraph is deleted, mark and all", () => {
    const { base } = check(p("One") + p("Two") + p("Three"), p("One") + p("Three"));
    const gone = Array.from(base.getElementsByTagNameNS(W, "p"))[1];
    expect(gone.getElementsByTagNameNS(W, "delText")[0].textContent).toBe("Two");
    expect(gone.getElementsByTagNameNS(W, "pPr")[0].getElementsByTagNameNS(W, "del")).toHaveLength(1);
  });

  it("a rewritten paragraph that shares nothing is deleted and inserted whole", () => {
    check(p("Alpha beta gamma") + p("End"), p("Completely different sentence") + p("End"));
  });

  it("several edits at once", () => {
    check(
      p("Intro text stays.") + p("This line changes a bit.") + p("This one goes.") + p("Tail"),
      p("New first line.") + p("Intro text stays.") + p("This line changed a lot.") + p("Tail") + p("Appendix")
    );
  });

  it("tabs and line breaks survive in kept and inserted text", () => {
    const tab = `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>`;
    const { base } = check(tab, `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>d</w:t></w:r></w:p>`);
    expect(base.getElementsByTagNameNS(W, "tab").length).toBeGreaterThan(0);
    expect(base.getElementsByTagNameNS(W, "br").length).toBeGreaterThan(0);
  });

  it("a changed table is deleted row by row and the new one inserted", () => {
    const { base } = check(table(["a"], ["b"]) + p("x"), table(["a"], ["c"]) + p("x"));
    const marks = Array.from(base.getElementsByTagNameNS(W, "trPr")).map((tr) => tr.firstElementChild!.localName);
    expect(marks).toEqual(["del", "del", "ins", "ins"]);
  });

  it("an unchanged table is left as it was", () => {
    const { base, result } = check(table(["a"]) + p("x"), table(["a"]) + p("y"));
    expect(base.getElementsByTagNameNS(W, "trPr")).toHaveLength(0);
    expect(result.changes).toBe(2);
  });

  it("inserted content from the other file drops what points into it (pictures, links, notes, lists)", () => {
    const other =
      p("Same") +
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>` +
      `<w:hyperlink r:id="rId7"><w:r><w:t>linked</w:t></w:r></w:hyperlink>` +
      `<w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId8"/></w:drawing></w:r>` +
      `<w:r><w:footnoteReference w:id="3"/></w:r><w:bookmarkStart w:id="1" w:name="x"/><w:r><w:t xml:space="preserve"> text</w:t></w:r></w:p>`;
    const base = body(p("Same"));
    const result = compareBodies(base, body(other), OPTS);
    const added = Array.from(base.getElementsByTagNameNS(W, "p"))[1];
    expect(added.getElementsByTagNameNS(W, "numPr")).toHaveLength(0);
    expect(added.getElementsByTagNameNS(W, "hyperlink")).toHaveLength(0);
    expect(added.getElementsByTagNameNS(W, "drawing")).toHaveLength(0);
    expect(added.getElementsByTagNameNS(W, "footnoteReference")).toHaveLength(0);
    expect(added.getElementsByTagNameNS(W, "bookmarkStart")).toHaveLength(0);
    expect(texts(base)[1]).toBe("linked text");
    expect(result.dropped).toBeGreaterThan(0);
  });

  it("revision ids are unique and above the document's own ids", () => {
    const { base } = check(`<w:p><w:bookmarkStart w:id="50" w:name="b"/><w:r><w:t>one two</w:t></w:r><w:bookmarkEnd w:id="50"/></w:p>`, p("one three"));
    const ids = Array.from(base.querySelectorAll("*")).filter((e) => ["ins", "del"].includes(e.localName)).map((e) => Number(e.getAttributeNS(W, "id")));
    expect(new Set(ids).size).toBe(ids.length);
    expect(Math.min(...ids)).toBeGreaterThan(50);
  });

  it("paragraph properties stay in schema order when a mark revision is added", () => {
    const { base } = check(p("Gone", "", `<w:pStyle w:val="Quote"/><w:jc w:val="center"/>`) + p("Kept"), p("Kept"));
    const pPr = base.getElementsByTagNameNS(W, "pPr")[0];
    expect(Array.from(pPr.children).map((c) => c.localName)).toEqual(["pStyle", "jc", "rPr"]);
  });

  it("the final section break stays last", () => {
    const { base } = check(p("a"), p("a") + p("b"));
    expect(base.lastElementChild!.localName).toBe("sectPr");
  });
});

describe("comparisonName", () => {
  it("is Google's 'Comparison of A and B'", () => {
    expect(comparisonName("Draft.docx", "Final.docx")).toBe("Comparison of Draft and Final.docx");
  });
});

import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { fixture } from "./helpers";

describe("compareBodies on real files", () => {
  it("the sample against itself: no changes, body unchanged", async () => {
    const a = await DocxDocument.load(fixture("sample.docx")), b = await DocxDocument.load(fixture("sample.docx"));
    const before = new XMLSerializer().serializeToString(a.body);
    expect(compareBodies(a.body, b.body, OPTS).changes).toBe(0);
    expect(new XMLSerializer().serializeToString(a.body)).toBe(before);
  });

  it("the sample against a blank document: accept gives the blank's text, reject the sample's; it reopens", async () => {
    const a = await DocxDocument.load(fixture("sample.docx")), b = await DocxDocument.load(await blankDocx());
    const want = { base: texts(a.body), other: texts(b.body) };
    expect(compareBodies(a.body, b.body, OPTS).changes).toBeGreaterThan(0);
    const accepted = clone(a.body), rejected = clone(a.body);
    acceptRevisions(accepted);
    rejectRevisions(rejected);
    expect(texts(rejected)).toEqual(want.base);
    expect(texts(accepted).filter(Boolean)).toEqual(want.other.filter(Boolean));
    const reopened = await DocxDocument.load(await a.save());
    expect(revisions(reopened.body).length).toBeGreaterThan(0);
  });

  it("a blank document against the sample: the sample's text arrives as insertions", async () => {
    const a = await DocxDocument.load(await blankDocx()), b = await DocxDocument.load(fixture("sample.docx"));
    const result = compareBodies(a.body, b.body, OPTS);
    expect(result.changes).toBeGreaterThan(0);
    const accepted = clone(a.body);
    acceptRevisions(accepted);
    expect(texts(accepted).join(" ")).toContain(texts(b.body).find((t) => t.length > 5)!);
  });
});
