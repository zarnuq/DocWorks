import { describe, expect, it } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import { Node as PMNode } from "prosemirror-model";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { attrsOf, docxSchema as schema } from "../src/domain/docx/schema";
import { chipAttrs, DATE_FORMATS, formatDate, newChip, type ChipAttrs } from "../src/domain/docx/chips";
import { docxToMarkdown } from "../src/domain/docx/to-markdown";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { serializeXml, W } from "./helpers";

const ctx = { imageSrc: () => null, hyperlinkHref: () => null };
const wctx = { nextDocPrId: () => 1 };

function body(inner: string): Element {
  const xml = `<w:document xmlns:w="${W}"><w:body>${inner}<w:sectPr/></w:body></w:document>`;
  return new DOMParser().parseFromString(xml, "application/xml").getElementsByTagNameNS(W, "body")[0];
}

const DATE_SDT =
  `<w:sdt><w:sdtPr><w:id w:val="42"/><w:date w:fullDate="2026-03-04T00:00:00Z"><w:dateFormat w:val="MMM d, yyyy"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr>` +
  `<w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>Mar 4, 2026</w:t></w:r></w:sdtContent></w:sdt>`;
const DROP_SDT =
  `<w:sdt><w:sdtPr><w:alias w:val="Status"/><w:id w:val="7"/><w:dropDownList w:lastValue="b"><w:listItem w:displayText="Not started" w:value="a"/><w:listItem w:displayText="In progress" w:value="b"/><w:listItem w:displayText="Done" w:value="c"/></w:dropDownList></w:sdtPr>` +
  `<w:sdtContent><w:r><w:t>In progress</w:t></w:r></w:sdtContent></w:sdt>`;

function chips(doc: PMNode): { pos: number; node: PMNode }[] {
  const out: { pos: number; node: PMNode }[] = [];
  doc.descendants((node, pos) => void (node.type === schema.nodes.chip && out.push({ pos, node })));
  return out;
}

function roundTrip(inner: string, edit?: (doc: PMNode) => PMNode) {
  const b = body(`<w:p><w:r><w:t xml:space="preserve">Due </w:t></w:r>${inner}<w:r><w:t xml:space="preserve"> ok</w:t></w:r></w:p>`);
  const before = serializeXml(b);
  const { doc, origins } = readDocx(b, ctx);
  writeBody(b, edit ? edit(doc) : doc, origins, wctx);
  return { doc, before, after: serializeXml(b), body: b };
}

const sdtOf = (b: Element) => b.getElementsByTagNameNS(W, "sdt")[0];
const one = (el: Element, local: string) => el.getElementsByTagNameNS(W, local)[0] ?? null;

describe("formatDate", () => {
  it("writes Word's date format tokens", () => {
    expect(formatDate("2026-03-04", "MMM d, yyyy")).toBe("Mar 4, 2026");
    expect(formatDate("2026-03-04", "M/d/yyyy")).toBe("3/4/2026");
    expect(formatDate("2026-03-04", "dd/MM/yyyy")).toBe("04/03/2026");
    expect(formatDate("2026-03-04", "yyyy-MM-dd")).toBe("2026-03-04");
    expect(formatDate("2026-03-04", "dddd, MMMM d, yyyy")).toBe("Wednesday, March 4, 2026");
    expect(formatDate("2026-11-21", "ddd d MMM yy")).toBe("Sat 21 Nov 26");
  });

  it("keeps quoted literal text", () => {
    expect(formatDate("2026-03-04", "d 'of' MMMM")).toBe("4 of March");
  });

  it("every offered format renders", () => {
    for (const [, f] of DATE_FORMATS) expect(formatDate("2026-12-31", f)).not.toMatch(/[yMd]{2}/);
  });
});

describe("reading chips", () => {
  it("a date content control becomes a date chip with its value, format and text", () => {
    const { doc } = roundTrip(DATE_SDT);
    const [c] = chips(doc);
    expect(attrsOf<ChipAttrs>(c.node)).toMatchObject({ kind: "date", date: "2026-03-04", format: "MMM d, yyyy", text: "Mar 4, 2026" });
    expect(c.node.marks.some((m) => m.type === schema.marks.bold)).toBe(true);
    expect(doc.textContent).toBe("Due Mar 4, 2026 ok");
  });

  it("a dropdown content control becomes a dropdown chip with its items", () => {
    const { doc } = roundTrip(DROP_SDT);
    const a = attrsOf<ChipAttrs>(chips(doc)[0].node);
    expect(a.kind).toBe("dropdown");
    expect(a.text).toBe("In progress");
    expect(a.options).toEqual([{ text: "Not started", value: "a" }, { text: "In progress", value: "b" }, { text: "Done", value: "c" }]);
  });

  it("a combo box reads like a dropdown", () => {
    const { doc } = roundTrip(DROP_SDT.replace(/dropDownList/g, "comboBox"));
    expect(attrsOf<ChipAttrs>(chips(doc)[0].node).kind).toBe("dropdown");
  });

  it("other inline content controls keep their old reading (plain text inside)", () => {
    const { doc } = roundTrip(`<w:sdt><w:sdtPr><w:id w:val="3"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Plain</w:t></w:r></w:sdtContent></w:sdt>`);
    expect(chips(doc)).toHaveLength(0);
    expect(doc.textContent).toBe("Due Plain ok");
  });

  it("a date control holding more than runs (a field) isn't a chip", () => {
    const { doc } = roundTrip(DATE_SDT.replace("<w:sdtContent>", "<w:sdtContent><w:fldSimple w:instr=\" DATE \"><w:r><w:t>x</w:t></w:r></w:fldSimple>"));
    expect(chips(doc)).toHaveLength(0);
  });

  it("a date control without a date shows its placeholder text", () => {
    const { doc } = roundTrip(DATE_SDT.replace(' w:fullDate="2026-03-04T00:00:00Z"', "").replace("<w:id", "<w:showingPlcHdr/><w:id").replace("Mar 4, 2026", "Click to enter a date."));
    const a = attrsOf<ChipAttrs>(chips(doc)[0].node);
    expect(a.date).toBeNull();
    expect(a.text).toBe("Click to enter a date.");
  });
});

describe("writing chips", () => {
  it("unchanged chips are written back byte-identical", () => {
    for (const sdt of [DATE_SDT, DROP_SDT]) {
      const { before, after } = roundTrip(sdt);
      expect(after).toBe(before);
    }
  });

  it("an unchanged chip keeps its exact XML when other text in its paragraph is edited", () => {
    for (const sdt of [DATE_SDT, DROP_SDT]) {
      const { body: b, before } = roundTrip(sdt, (doc) => EditorState.create({ doc }).tr.insertText("!", doc.content.size - 1).doc);
      const original = new DOMParser().parseFromString(before, "application/xml").getElementsByTagNameNS(W, "sdt")[0];
      expect(serializeXml(sdtOf(b))).toBe(serializeXml(original));
      expect(b.textContent).toContain("ok!");
    }
  });

  it("a new date updates w:fullDate and the shown text, keeping id, format and run formatting", () => {
    const { body: b } = roundTrip(DATE_SDT, (doc) => {
      const [c] = chips(doc);
      const tr = EditorState.create({ doc }).tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, date: "2026-11-21", text: formatDate("2026-11-21", "MMM d, yyyy") });
      return tr.doc;
    });
    const sdt = sdtOf(b);
    expect(one(sdt, "date")!.getAttributeNS(W, "fullDate")).toBe("2026-11-21T00:00:00Z");
    expect(one(sdt, "id")!.getAttributeNS(W, "val")).toBe("42");
    expect(one(sdt, "dateFormat")!.getAttributeNS(W, "val")).toBe("MMM d, yyyy");
    const content = one(sdt, "sdtContent")!;
    expect(content.textContent).toBe("Nov 21, 2026");
    expect(one(content, "b")).not.toBeNull();
  });

  it("a new date format is written as the control's w:dateFormat", () => {
    const { body: b } = roundTrip(DATE_SDT, (doc) => {
      const [c] = chips(doc);
      return EditorState.create({ doc }).tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, format: "yyyy-MM-dd", text: "2026-03-04" }).doc;
    });
    expect(one(sdtOf(b), "dateFormat")!.getAttributeNS(W, "val")).toBe("yyyy-MM-dd");
    expect(one(sdtOf(b), "sdtContent")!.textContent).toBe("2026-03-04");
  });

  it("picking a date drops the placeholder flag", () => {
    const placeholder = DATE_SDT.replace(' w:fullDate="2026-03-04T00:00:00Z"', "").replace("<w:id", "<w:showingPlcHdr/><w:id");
    const { body: b } = roundTrip(placeholder, (doc) => {
      const [c] = chips(doc);
      return EditorState.create({ doc }).tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, date: "2026-01-02", text: "Jan 2, 2026" }).doc;
    });
    expect(one(sdtOf(b), "showingPlcHdr")).toBeNull();
  });

  it("choosing another dropdown item updates lastValue and the text; edited items are written in order", () => {
    const { body: b } = roundTrip(DROP_SDT, (doc) => {
      const [c] = chips(doc);
      const options = [{ text: "Done", value: "c" }, { text: "Blocked", value: "Blocked" }];
      return EditorState.create({ doc }).tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, options, text: "Blocked" }).doc;
    });
    const list = one(sdtOf(b), "dropDownList")!;
    expect(list.getAttributeNS(W, "lastValue")).toBe("Blocked");
    expect(Array.from(list.children).map((i) => i.getAttributeNS(W, "displayText"))).toEqual(["Done", "Blocked"]);
    expect(one(sdtOf(b), "alias")!.getAttributeNS(W, "val")).toBe("Status");
    expect(one(sdtOf(b), "sdtContent")!.textContent).toBe("Blocked");
  });

  it("a new date chip is a complete Word date picker control in schema order", () => {
    const b = body(`<w:p/>`);
    const { doc, origins } = readDocx(b, ctx);
    const chip = schema.nodes.chip.create(newChip("date", { date: "2026-11-21" }));
    writeBody(b, doc.copy(doc.content.replaceChild(0, schema.nodes.paragraph.create(null, chip))), origins, wctx);
    const sdt = sdtOf(b);
    const pr = one(sdt, "sdtPr")!;
    expect(Array.from(pr.children).map((c) => c.localName)).toEqual(["id", "date"]);
    expect(Number(one(pr, "id")!.getAttributeNS(W, "val"))).toBeGreaterThan(0);
    const date = one(pr, "date")!;
    expect(date.getAttributeNS(W, "fullDate")).toBe("2026-11-21T00:00:00Z");
    expect(Array.from(date.children).map((c) => c.localName)).toEqual(["dateFormat", "lid", "storeMappedDataAs", "calendar"]);
    expect(one(sdt, "sdtContent")!.textContent).toBe("Nov 21, 2026");
    expect(sdt.parentElement!.localName).toBe("p");
  });

  it("a new dropdown chip writes every item, the first chosen", () => {
    const b = body(`<w:p/>`);
    const { doc, origins } = readDocx(b, ctx);
    const chip = schema.nodes.chip.create(newChip("dropdown", { options: ["Low", "Medium", "High"] }));
    writeBody(b, doc.copy(doc.content.replaceChild(0, schema.nodes.paragraph.create(null, chip))), origins, wctx);
    const list = one(sdtOf(b), "dropDownList")!;
    expect(Array.from(list.children).map((i) => [i.getAttributeNS(W, "displayText"), i.getAttributeNS(W, "value")])).toEqual([["Low", "Low"], ["Medium", "Medium"], ["High", "High"]]);
    expect(one(sdtOf(b), "sdtContent")!.textContent).toBe("Low");
  });

  it("chip formatting from marks is written in the run", () => {
    const b = body(`<w:p/>`);
    const { doc, origins } = readDocx(b, ctx);
    const chip = schema.nodes.chip.create(newChip("date", { date: "2026-11-21" }), null, [schema.marks.italic.create()]);
    writeBody(b, doc.copy(doc.content.replaceChild(0, schema.nodes.paragraph.create(null, chip))), origins, wctx);
    expect(one(one(sdtOf(b), "sdtContent")!, "i")).not.toBeNull();
  });
});

describe("newChip and chipAttrs", () => {
  it("today's date chip takes Google's default format", () => {
    expect(newChip("date", { date: "2026-11-21" })).toMatchObject({ kind: "date", date: "2026-11-21", format: "MMM d, yyyy", text: "Nov 21, 2026", xml: "" });
  });

  it("drops blank and duplicate dropdown items", () => {
    expect(newChip("dropdown", { options: ["A", " ", "A", "B "] }).options).toEqual([{ text: "A", value: "A" }, { text: "B", value: "B" }]);
  });

  it("chipAttrs reads a w:sdt element", () => {
    const el = new DOMParser().parseFromString(`<x xmlns:w="${W}">${DROP_SDT}</x>`, "application/xml").documentElement.firstElementChild!;
    expect(chipAttrs(el)).toMatchObject({ kind: "dropdown", text: "In progress" });
  });
});

describe("chips elsewhere", () => {
  it("count as their text for find and word count (textBetween)", () => {
    const { doc } = roundTrip(DATE_SDT);
    expect(doc.textBetween(0, doc.content.size, "\n")).toBe("Due Mar 4, 2026 ok");
  });

  it("are their text in Markdown export", async () => {
    const { doc } = roundTrip(DROP_SDT);
    const docx = await DocxDocument.load(await blankDocx());
    expect((await docxToMarkdown(docx, "img", doc)).markdown).toContain("Due In progress ok");
  });

  it("can be selected and deleted like other atoms", () => {
    const { doc } = roundTrip(DATE_SDT);
    const [c] = chips(doc);
    const state = EditorState.create({ doc });
    const tr = state.tr.setSelection(TextSelection.create(doc, c.pos, c.pos + 1)).deleteSelection();
    expect(tr.doc.textContent).toBe("Due  ok");
  });
});

import { atTrigger, cmd } from "../src/presentation/docx/docx-editor";

function stateWith(text: string, caret = text.length + 1, marks: string[] = []) {
  const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, text ? schema.text(text, marks.map((m) => schema.marks[m].create())) : null));
  return EditorState.create({ doc, selection: TextSelection.create(doc, caret) });
}
const run = (state: EditorState, command: (s: EditorState, d?: (tr: import("prosemirror-state").Transaction) => void) => boolean) => {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
};

describe("atTrigger (Google's @ at the start of a word)", () => {
  it("is the @ position when @ was just typed at a word start", () => {
    expect(atTrigger(stateWith("@"))).toBe(1);
    expect(atTrigger(stateWith("Due @"))).toBe(5);
    expect(atTrigger(stateWith("Due\t@"))).toBe(5);
  });

  it("is null inside a word (email addresses) or without an @", () => {
    expect(atTrigger(stateWith("me@"))).toBeNull();
    expect(atTrigger(stateWith("Due "))).toBeNull();
    expect(atTrigger(stateWith(""))).toBeNull();
  });

  it("is null with a selection", () => {
    const s = stateWith("Due @");
    expect(atTrigger(s.apply(s.tr.setSelection(TextSelection.create(s.doc, 1, 5))))).toBeNull();
  });
});

describe("chip commands", () => {
  it("insertChip puts a chip at the caret, with the caret's formatting", () => {
    const s = run(stateWith("Due ", 5, ["bold"]), cmd.insertChip(newChip("date", { date: "2026-11-21" })));
    const [c] = chips(s.doc);
    expect(c.pos).toBe(5);
    expect(c.node.marks.map((m) => m.type.name)).toEqual(["bold"]);
    expect(s.doc.textContent).toBe("Due Nov 21, 2026");
    expect(s.selection.from).toBe(6);
  });

  it("insertChip replaces the typed @ when given its position", () => {
    const s = run(stateWith("Due @"), cmd.insertChip(newChip("dropdown", { options: ["Yes", "No"] }), 5));
    expect(s.doc.textContent).toBe("Due Yes");
    expect(chips(s.doc)).toHaveLength(1);
  });

  it("insertChip keeps the @ if it was edited away since the menu opened", () => {
    const s = run(stateWith("Due x"), cmd.insertChip(newChip("dropdown", { options: ["Yes"] }), 5));
    expect(s.doc.textContent).toBe("Due xYes");
  });

  it("insertChip leaves links out of the chip's marks", () => {
    const link = schema.marks.link.create({ href: "https://x" });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text("ab", [link])));
    const s = run(EditorState.create({ doc, selection: TextSelection.create(doc, 2) }), cmd.insertChip(newChip("date", { date: "2026-01-01" })));
    expect(chips(s.doc)[0].node.marks).toHaveLength(0);
  });

  it("setChip changes a chip's attrs in place", () => {
    const s0 = run(stateWith("Due "), cmd.insertChip(newChip("date", { date: "2026-11-21" })));
    const pos = chips(s0.doc)[0].pos;
    const s = run(s0, cmd.setChip(pos, { date: "2026-12-25", text: "Dec 25, 2026" }));
    expect(attrsOf<ChipAttrs>(chips(s.doc)[0].node)).toMatchObject({ date: "2026-12-25", text: "Dec 25, 2026" });
  });

  it("setChip refuses a position without a chip", () => {
    expect(cmd.setChip(1, { text: "x" })(stateWith("abc"))).toBe(false);
  });

  it("insertLinkedText inserts text linked to an address, replacing the @", () => {
    const s = run(stateWith("See @"), cmd.insertLinkedText("Plan.md", "obsidian://open?vault=v&file=Plan.md", 5));
    expect(s.doc.textContent).toBe("See Plan.md");
    const node = s.doc.child(0).child(1);
    expect(node.marks.find((m) => m.type === schema.marks.link)!.attrs.href).toBe("obsidian://open?vault=v&file=Plan.md");
    // Typing after it isn't linked.
    expect(s.storedMarks ?? []).toEqual([]);
  });
});
