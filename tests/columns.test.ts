import { describe, it, expect } from "vitest";
import { EditorState, TextSelection, Command } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, docxPlugins, EditorContext } from "../src/presentation/docx/docx-editor";
import { pagesKey, pagesPlugin } from "../src/presentation/docx/docx-pages";
import { W } from "./helpers";

const ctx: EditorContext = { ensureList: () => "1", listKind: () => null, styles: new DocxStyles(null, null) };
const writeCtx = { nextDocPrId: () => 1, ensureParagraphStyles: () => {} };
const p = (t: string, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r><w:t>${t}</w:t></w:r></w:p>`;
const sect = (inner = "") => `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>${inner}</w:sectPr>`;

function load(body: string) {
  const el = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${body}</w:body>`, "application/xml").documentElement;
  const { doc, origins } = readDocx(el, { imageSrc: () => null, hyperlinkHref: () => null });
  return { el, doc, origins };
}
function run(st: EditorState, command: Command): EditorState {
  let next = st;
  expect(command(st, (tr) => (next = st.apply(tr)))).toBe(true);
  return next;
}
/** Position inside the paragraph whose text is `text`. */
function inside(doc: ReturnType<typeof load>["doc"], text: string, offset = 1): number {
  let at = -1;
  doc.descendants((n, pos) => {
    if (at < 0 && n.type === schema.nodes.paragraph && n.textContent === text) at = pos + 1 + offset;
    return at < 0;
  });
  return at;
}
const shape = (doc: ReturnType<typeof load>["doc"]) => {
  const out: string[] = [];
  doc.forEach((n) => out.push(n.type === schema.nodes.columns ? `cols${n.attrs.count}[${n.content.content.map((c) => c.textContent).join(",")}]` : n.textContent));
  return out;
};
const colsOf = (sectPr: Element | undefined) => sectPr?.getElementsByTagNameNS(W, "cols")[0];
const typeOf = (sectPr: Element | undefined) => sectPr?.getElementsByTagNameNS(W, "type")[0]?.getAttributeNS(W, "val") ?? null;

describe("reading sections in columns", () => {
  it("groups the blocks of a section with two or more columns", () => {
    const { doc } = load(p("a") + p("b", sect(`<w:type w:val="continuous"/><w:cols w:num="2" w:space="360" w:sep="1"/>`)) + p("c") + sect());
    expect(shape(doc)).toEqual(["cols2[a,b]", "c"]);
    expect(doc.child(0).attrs).toMatchObject({ count: 2, space: 360, sep: true });
    expect(doc.child(0).child(1).attrs.sectPr).toContain("w:cols");
  });

  it("puts the final section in columns from the body's section properties", () => {
    const { doc } = load(p("a") + p("b") + sect(`<w:cols w:num="3"/>`));
    expect(shape(doc)).toEqual(["cols3[a,b]"]);
    expect(doc.child(0).attrs.space).toBe(720);
    expect(doc.attrs.sectPr).toContain('w:num="3"');
  });

  it("writes an unchanged document back unchanged", () => {
    const body = p("a") + p("b", sect(`<w:cols w:num="2"/>`)) + p("c") + sect();
    const { el, doc, origins } = load(body);
    const before = new XMLSerializer().serializeToString(el);
    writeBody(el, doc, origins, writeCtx);
    expect(new XMLSerializer().serializeToString(el)).toBe(before);
  });

  it("draws columns with CSS columns, the spacing and the line between", () => {
    const { doc } = load(p("a") + sect(`<w:cols w:num="2" w:space="720" w:sep="1"/>`));
    const view = new EditorView(document.body.appendChild(document.createElement("div")), { state: EditorState.create({ doc }) });
    const el = view.dom.querySelector<HTMLElement>(".docx-columns")!;
    expect(el.style.getPropertyValue("--docx-cols")).toBe("2");
    expect(el.style.getPropertyValue("--docx-col-gap")).toBe("0.5in");
    expect(el.classList.contains("docx-columns-sep")).toBe(true);
    view.destroy();
  });
});

describe("Format › Columns", () => {
  it("puts the whole section in columns when nothing is selected, saving them on the section", () => {
    const { el, doc, origins } = load(p("a") + p("b") + sect());
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "a")) });
    st = run(st, cmd.setColumns({ count: 2, space: 720, sep: true }));
    expect(shape(st.doc)).toEqual(["cols2[a,b]"]);
    expect(cmd.columnsAt(st)).toEqual({ count: 2, space: 720, sep: true });
    writeBody(el, st.doc, origins, writeCtx);
    const cols = colsOf(Array.from(el.children).pop());
    expect([cols?.getAttributeNS(W, "num"), cols?.getAttributeNS(W, "space"), cols?.getAttributeNS(W, "sep")]).toEqual(["2", "720", "1"]);
  });

  it("changes and removes columns again", () => {
    const { el, doc, origins } = load(p("a") + p("b") + sect(`<w:cols w:num="2" w:space="720"/>`));
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "b")) });
    st = run(st, cmd.setColumns({ count: 3, space: 360, sep: false }));
    expect(shape(st.doc)).toEqual(["cols3[a,b]"]);
    st = run(st, cmd.setColumns({ count: 1, space: 720, sep: false }));
    expect(shape(st.doc)).toEqual(["a", "b"]);
    expect(cmd.columnsAt(st).count).toBe(1);
    writeBody(el, st.doc, origins, writeCtx);
    const cols = colsOf(Array.from(el.children).pop());
    expect(cols === undefined || (cols.getAttributeNS(W, "num") ?? "1") === "1").toBe(true);
  });

  it("puts only the selected paragraphs in columns, between continuous section breaks (as Google Docs does)", () => {
    const { el, doc, origins } = load(p("a") + p("b") + p("c") + p("d") + sect(`<w:type w:val="nextPage"/>`));
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "b", 0), inside(doc, "c", 1)) });
    st = run(st, cmd.setColumns({ count: 2, space: 720, sep: false }));
    expect(shape(st.doc)).toEqual(["a", "cols2[b,c]", "d"]);
    writeBody(el, st.doc, origins, writeCtx);
    const sects = Array.from(el.getElementsByTagNameNS(W, "sectPr"));
    expect(sects).toHaveLength(3);
    // "a" ends the first part (as the section started), "c" the columns, then the rest continues on the same page.
    expect(sects[0].parentElement!.parentElement!.textContent).toBe("a");
    expect([typeOf(sects[0]), colsOf(sects[0])?.getAttributeNS(W, "num") ?? "1"]).toEqual(["nextPage", "1"]);
    expect(sects[1].parentElement!.parentElement!.textContent).toBe("c");
    expect([typeOf(sects[1]), colsOf(sects[1])?.getAttributeNS(W, "num")]).toEqual(["continuous", "2"]);
    expect(typeOf(sects[2])).toBe("continuous");
    // Each copy keeps the page setup.
    expect(sects[0].getElementsByTagNameNS(W, "pgMar")).toHaveLength(1);
  });
});

describe("Insert › Break", () => {
  it("inserts a column break", () => {
    const { doc } = load(p("ab") + sect(`<w:cols w:num="2"/>`));
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "ab")) });
    st = run(st, cmd.insertColumnBreak);
    const para = st.doc.child(0).child(0);
    expect(para.childCount).toBe(3);
    expect(para.child(1).attrs.xml).toContain('w:type="column"');
  });

  it("inserts a continuous section break, splitting the paragraph", () => {
    const { el, doc, origins } = load(p("ab") + sect());
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "ab")) });
    st = run(st, cmd.insertSectionBreak("continuous"));
    expect(shape(st.doc)).toEqual(["a", "b"]);
    writeBody(el, st.doc, origins, writeCtx);
    const sects = Array.from(el.getElementsByTagNameNS(W, "sectPr"));
    expect(sects.map((s) => s.parentElement!.localName)).toEqual(["pPr", "body"]);
    expect(typeOf(sects[1])).toBe("continuous");
  });

  it("starts a new page at a next-page section break (at the end of a paragraph it just ends the section there)", async () => {
    const { doc } = load(p("a") + p("b") + p("c") + sect());
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, inside(doc, "a", 1)) });
    st = run(st, cmd.insertSectionBreak("nextPage"));
    const host = document.body.appendChild(document.createElement("div"));
    const setup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
    const view = new EditorView(host, { state: EditorState.create({ doc: st.doc, plugins: [...docxPlugins(ctx), pagesPlugin({ setup: () => setup, header: () => null, footer: () => null })] }) });
    await new Promise((r) => setTimeout(r, 100));
    expect(pagesKey.getState(view.state)!.pages).toBe(2);
    expect(host.querySelector(".docx-page-gap")!.nextElementSibling!.textContent).toBe("b");
    view.destroy();
  });
});
