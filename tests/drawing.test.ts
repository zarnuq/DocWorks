import { describe, it, expect } from "vitest";
import { DOMSerializer } from "prosemirror-model";
import { EditorState, NodeSelection } from "prosemirror-state";
import { INLINE, placementCss, placementOf, setPlacement } from "../src/domain/docx/drawing";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { W } from "./helpers";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NSS = `xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;
const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><a:blip r:embed="rId9"/></a:graphicData></a:graphic>`;
const common = `<wp:extent cx="952500" cy="476250"/><wp:docPr id="3" name="Picture 3" descr="logo"/><wp:cNvGraphicFramePr/>`;
const anchor = (inner: string, attrs = `behindDoc="0"`) =>
  `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" ${attrs} locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${inner}</wp:anchor></w:drawing>`;
const pos = (h: string, v = `<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>`) => `<wp:positionH relativeFrom="column">${h}</wp:positionH>${v}`;
const drawingEl = (xml: string) => new DOMParser().parseFromString(`<w:r ${NSS}>${xml}</w:r>`, "application/xml").documentElement.firstElementChild!;

describe("where a picture sits (Word's / Google Docs' wrap options)", () => {
  it("reads in-line pictures as in line", () => {
    expect(placementOf(drawingEl(`<w:drawing><wp:inline>${common}${graphic}</wp:inline></w:drawing>`))).toEqual(INLINE);
  });

  it("reads a picture text wraps around, offset from the column and paragraph", () => {
    const p = placementOf(drawingEl(anchor(pos(`<wp:posOffset>768310</wp:posOffset>`) + common + `<wp:wrapTight wrapText="bothSides"/>` + graphic, `behindDoc="1"`)));
    // behindDoc only matters without wrapping; a tight wrap is a wrap.
    expect(p).toEqual({ wrap: "square", x: 81, y: 10, align: null });
  });

  it("reads break text, behind and in front of text, and aligned pictures", () => {
    expect(placementOf(drawingEl(anchor(pos(`<wp:align>center</wp:align>`) + common + `<wp:wrapTopAndBottom/>` + graphic))).wrap).toBe("topBottom");
    expect(placementOf(drawingEl(anchor(pos(`<wp:align>right</wp:align>`) + common + `<wp:wrapNone/>` + graphic, `behindDoc="1"`)))).toMatchObject({ wrap: "behind", align: "right" });
    expect(placementOf(drawingEl(anchor(pos(`<wp:posOffset>0</wp:posOffset>`) + common + `<wp:wrapNone/>` + graphic))).wrap).toBe("front");
  });

  it("lays pictures out with CSS: floats for wrapped text, blocks for break text, absolute for behind / in front", () => {
    expect(placementCss({ wrap: "square", x: 81, y: 10, align: null })).toMatchObject({ cls: "docx-wrap" });
    expect(placementCss({ wrap: "square", x: 81, y: 10, align: null }).style).toContain("float: left; margin-left: 81px; margin-top: 10px");
    expect(placementCss({ wrap: "square", x: 0, y: 0, align: "right" }).style).toContain("float: right");
    // Offset toward the right of the column: floats right, so text wraps on its left (Word wraps both sides).
    expect(placementCss({ wrap: "square", x: 336, y: 0, align: null }, 240).style).toContain("float: right; margin-right: 48px");
    expect(placementCss({ wrap: "topBottom", x: 0, y: 0, align: "center" }).style).toContain("display: block; margin-left: auto; margin-right: auto");
    expect(placementCss({ wrap: "behind", x: 20, y: 5, align: null })).toEqual({ cls: "docx-float docx-behind", style: "left: 20px; top: 5px" });
    expect(placementCss(INLINE)).toEqual({ cls: "", style: "" });
  });

  it("rewrites the drawing between in line and floating, keeping the picture", () => {
    const run = new DOMParser().parseFromString(`<w:r ${NSS}><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${common}${graphic}</wp:inline></w:drawing></w:r>`, "application/xml").documentElement;
    setPlacement(run, { wrap: "square", x: 50, y: 0, align: null });
    const a = run.getElementsByTagNameNS(WP, "anchor")[0];
    expect(Array.from(a.children).map((c) => c.localName)).toEqual(["simplePos", "positionH", "positionV", "extent", "wrapSquare", "docPr", "cNvGraphicFramePr", "graphic"]);
    expect(a.getElementsByTagNameNS(WP, "posOffset")[0].textContent).toBe(String(50 * 9525));
    expect(placementOf(run.firstElementChild!)).toEqual({ wrap: "square", x: 50, y: 0, align: null });
    setPlacement(run, { wrap: "behind", x: 50, y: 0, align: null });
    expect(placementOf(run.firstElementChild!).wrap).toBe("behind");
    setPlacement(run, INLINE);
    const inline = run.getElementsByTagNameNS(WP, "inline")[0];
    expect(Array.from(inline.children).map((c) => c.localName)).toEqual(["extent", "docPr", "cNvGraphicFramePr", "graphic"]);
  });

  it("reads, draws and writes back a floating picture's placement", () => {
    const xml = `<w:body ${NSS}><w:p><w:r>${anchor(pos(`<wp:posOffset>768310</wp:posOffset>`) + common + `<wp:wrapSquare wrapText="bothSides"/>` + graphic)}</w:r><w:r><w:t>text</w:t></w:r></w:p></w:body>`;
    const body = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    const { doc, origins } = readDocx(body, { imageSrc: () => "data:,", hyperlinkHref: () => null });
    const img = doc.child(0).child(0);
    expect(img.attrs).toMatchObject({ wrap: "square", x: 81 });
    const dom = DOMSerializer.fromSchema(schema).serializeNode(img, { document }) as HTMLElement;
    expect(dom.style.float).toBe("left");
    // Unchanged: written back as it was. Changed to in line: rewritten.
    const before = new XMLSerializer().serializeToString(body);
    writeBody(body, doc, origins, { nextDocPrId: () => 1 });
    expect(new XMLSerializer().serializeToString(body)).toBe(before);
    let st = EditorState.create({ doc });
    st = st.apply(st.tr.setSelection(NodeSelection.create(st.doc, 1)).setNodeMarkup(1, undefined, { ...img.attrs, wrap: "inline" }));
    writeBody(body, st.doc, origins, { nextDocPrId: () => 1 });
    expect(body.getElementsByTagNameNS(WP, "anchor")).toHaveLength(0);
    expect(body.getElementsByTagNameNS(WP, "inline")).toHaveLength(1);
  });
});

describe("text boxes", () => {
  it("draws a text box's text in a box of its size and place", () => {
    const box =
      `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
      pos(`<wp:posOffset>190500</wp:posOffset>`) +
      `<wp:extent cx="1905000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="4" name="Text Box 4"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>` +
      `<wps:spPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
      `<wps:txbx><w:txbxContent><w:p><w:r><w:t>First line</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const body = new DOMParser().parseFromString(`<w:body ${NSS}><w:p>${box}<w:r><w:t>after</w:t></w:r></w:p></w:body>`, "application/xml").documentElement;
    const { doc } = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null });
    const div = document.createElement("div");
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content, { document }));
    const el = div.querySelector<HTMLElement>(".docx-textbox")!;
    expect(Array.from(el.querySelectorAll(".docx-textbox-p")).map((p) => p.textContent)).toEqual(["First line", "Second"]);
    expect([el.style.width, el.style.height, el.style.float, el.style.marginLeft]).toEqual(["200px", "100px", "left", "20px"]);
    expect(el.style.backgroundColor).toBe("rgb(255, 242, 204)");
    expect(el.style.borderColor).toBe("rgb(0, 0, 0)");
  });
});

describe("picture options in the editor", () => {
  it("places the picture and changes its wrap from the picture's bar, as Google Docs' image options", async () => {
    const { EditorView } = await import("prosemirror-view");
    const { DocxImageView } = await import("../src/presentation/docx/docx-image-view");
    const img = schema.nodes.image.create({ src: "data:,", width: 100, height: 50, wrap: "square", x: 40, y: 0 });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [img, schema.text("text")]));
    const view = new EditorView(document.body.appendChild(document.createElement("div")), {
      state: EditorState.create({ doc }),
      nodeViews: { image: (node, v, getPos) => new DocxImageView(node, v, getPos, { replace: () => {}, editAlt: () => {} }) },
    });
    const dom = view.dom.querySelector<HTMLElement>(".docx-image")!;
    expect(dom.style.float).toBe("left");
    expect(dom.style.marginLeft).toBe("40px");
    const select = dom.querySelector<HTMLSelectElement>("select.docx-image-wrap")!;
    expect(select.value).toBe("square");
    select.value = "inline";
    select.dispatchEvent(new Event("change"));
    expect(view.state.doc.child(0).child(0).attrs.wrap).toBe("inline");
    expect(view.dom.querySelector<HTMLElement>(".docx-image")!.style.float).toBe("");
    view.destroy();
  });
});

describe("editing text boxes", () => {
  const box = (fallback = true) =>
    `<w:r ${NSS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
    pos(`<wp:posOffset>0</wp:posOffset>`) +
    `<wp:extent cx="1905000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="4" name="Text Box 4"/><a:graphic><a:graphicData uri="x"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Old text</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
    (fallback ? `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Old text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` : "") +
    `</mc:AlternateContent></w:r>`;

  it("reads a text box's paragraphs and writes edits back to it and its fallback picture", async () => {
    const { readTextBox, writeTextBox } = await import("../src/domain/docx/textbox");
    const xml = box();
    const read = readTextBox(xml, { imageSrc: () => null, hyperlinkHref: () => null })!;
    expect(read.doc.textContent).toBe("Old text");
    expect(read.doc.child(0).attrs.align).toBe("center");
    let st = EditorState.create({ doc: read.doc });
    st = st.apply(st.tr.insertText("New ", 1));
    const out = writeTextBox(xml, st.doc, read.origins, { nextDocPrId: () => 1 });
    const root = new DOMParser().parseFromString(out, "application/xml").documentElement;
    const contents = Array.from(root.getElementsByTagNameNS(W, "txbxContent"));
    expect(contents.map((c) => c.textContent)).toEqual(["New Old text", "New Old text"]);
    // Formatting the editor doesn't touch is kept.
    expect(contents[0].getElementsByTagNameNS(W, "jc")).toHaveLength(1);
    expect(readTextBox("<w:r xmlns:w=\"" + W + "\"><w:t>x</w:t></w:r>", { imageSrc: () => null, hyperlinkHref: () => null })).toBeNull();
  });

  it("edits a text box in place in the document", async () => {
    const { EditorView } = await import("prosemirror-view");
    const { TextBoxView } = await import("../src/presentation/docx/docx-textbox-view");
    const { DocxStyles } = await import("../src/domain/docx/docx-styles");
    const ctx = { ensureList: () => "1", listKind: () => null, styles: new DocxStyles(null, null) };
    const opaque = schema.nodes.opaque.create({ xml: box(false), text: "" });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [opaque, schema.text("after")]));
    const view = new EditorView(document.body.appendChild(document.createElement("div")), {
      state: EditorState.create({ doc }),
      nodeViews: { opaque: (node, v, getPos) => new TextBoxView(node, v, getPos, ctx, { imageSrc: () => null, hyperlinkHref: () => null }, { nextDocPrId: () => 1 }) },
    });
    const inner = view.dom.querySelector(".docx-textbox .ProseMirror")!;
    expect(inner.textContent).toBe("Old text");
    const sub = (view.dom.querySelector(".docx-textbox") as HTMLElement & { pmView?: unknown }) && TextBoxView.editorOf(view.dom.querySelector(".docx-textbox")!)!;
    sub.dispatch(sub.state.tr.insertText("Edited ", 1));
    expect(view.state.doc.child(0).child(0).attrs.xml).toContain("Edited Old text");
    // The inner editor survives the outer update.
    expect(view.dom.querySelector(".docx-textbox .ProseMirror")!.textContent).toBe("Edited Old text");
    view.destroy();
  });
});
