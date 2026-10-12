import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addSlideComment, editSlideComment, removeSlideThread, resolveSlideThread, slideComments } from "../src/domain/pptx/comments";

const P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";

async function deck() {
  const doc = await PptxDocument.load(await blankPptx());
  return { doc, slide: doc.slides()[0].path };
}

describe("slide comments (PowerPoint modern comments)", () => {
  it("adds a comment thread with replies, as PowerPoint stores them", async () => {
    const { doc, slide } = await deck();
    const id = addSlideComment(doc, slide, "Fix the title\nand the date", "Ada Lovelace");
    addSlideComment(doc, slide, "Done", "Grace Hopper", id);
    const [t] = slideComments(doc, slide);
    expect(t).toMatchObject({ id, resolved: false, legacy: false });
    expect(t.messages.map((m) => [m.author, m.text])).toEqual([["Ada Lovelace", "Fix the title\nand the date"], ["Grace Hopper", "Done"]]);
    // Parts: authors on the presentation, a comment part on the slide linked from the slide's extension list.
    const authorsRel = doc.rels("ppt/presentation.xml").find((r) => r.uri === "http://schemas.microsoft.com/office/2018/10/relationships/authors")!;
    const authors = doc.xml(authorsRel.target)!.getElementsByTagNameNS(P188, "author");
    expect(Array.from(authors).map((a) => [a.getAttribute("name"), a.getAttribute("initials"), a.getAttribute("providerId")])).toEqual([["Ada Lovelace", "AL", "None"], ["Grace Hopper", "GH", "None"]]);
    const rel = doc.rels(slide).find((r) => r.uri === "http://schemas.microsoft.com/office/2018/10/relationships/comments")!;
    const ext = Array.from(doc.xml(slide)!.getElementsByTagNameNS(P, "ext")).find((e) => e.getAttribute("uri") === "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}")!;
    expect(ext.getElementsByTagNameNS(P188, "commentRel")[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")).toBe(rel.id);
    const cm = doc.xml(rel.target)!.getElementsByTagNameNS(P188, "cm")[0];
    // Schema order: anchor, replies, then the text.
    expect(Array.from(cm.children).map((c) => c.localName)).toEqual(["sldMkLst", "replyLst", "txBody"]);
    const sldMk = cm.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2013/main/command", "sldMk")[0];
    expect(sldMk.getAttribute("sldId")).toBe(doc.slides()[0].id);
    expect(Number(sldMk.getAttribute("cId"))).toBeGreaterThan(0);
    const saved = await PptxDocument.load(await doc.save());
    expect(slideComments(saved, slide)[0].messages).toHaveLength(2);
  });

  it("edits, resolves, reopens and deletes threads; one link and one author per name", async () => {
    const { doc, slide } = await deck();
    const a = addSlideComment(doc, slide, "First", "Ada");
    const b = addSlideComment(doc, slide, "Second", "Ada");
    editSlideComment(doc, slide, a, "First, edited");
    resolveSlideThread(doc, slide, b, true);
    expect(slideComments(doc, slide).map((t) => [t.messages[0].text, t.resolved])).toEqual([["First, edited", false], ["Second", true]]);
    resolveSlideThread(doc, slide, b, false);
    expect(slideComments(doc, slide)[1].resolved).toBe(false);
    removeSlideThread(doc, slide, a);
    expect(slideComments(doc, slide).map((t) => t.id)).toEqual([b]);
    expect(doc.rels(slide).filter((r) => r.type === "comments")).toHaveLength(1);
    expect(Array.from(doc.xml(slide)!.getElementsByTagNameNS(P188, "commentRel"))).toHaveLength(1);
  });

  it("reads legacy comments with their replies", async () => {
    const { doc, slide } = await deck();
    const legacy = `<p:cmLst xmlns:p="${P}" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:cm authorId="1" dt="2020-01-01T10:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Old note</p:text></p:cm><p:cm authorId="2" dt="2020-01-02T10:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Old reply</p:text><p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="1" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm></p:cmLst>`;
    const authors = `<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="1" name="Old Author" initials="OA" lastIdx="1" clrIdx="0"/><p:cmAuthor id="2" name="Other" initials="O" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>`;
    (doc as any).putXml("ppt/comments/comment1.xml", new DOMParser().parseFromString(legacy, "application/xml"));
    (doc as any).putXml("ppt/commentAuthors.xml", new DOMParser().parseFromString(authors, "application/xml"));
    (doc as any).addRel(slide, "comments", "ppt/comments/comment1.xml");
    (doc as any).addRel("ppt/presentation.xml", "commentAuthors", "ppt/commentAuthors.xml");
    const [t] = slideComments(doc, slide);
    expect(t).toMatchObject({ legacy: true, resolved: false });
    expect(t.messages.map((m) => [m.author, m.text])).toEqual([["Old Author", "Old note"], ["Other", "Old reply"]]);
    removeSlideThread(doc, slide, t.id);
    expect(slideComments(doc, slide)).toHaveLength(0);
  });

  it("leaves comments behind when a slide is duplicated, as Google Slides does", async () => {
    const { doc, slide } = await deck();
    addSlideComment(doc, slide, "Only here", "Ada");
    doc.duplicateSlide(0);
    const copy = doc.slides().find((s) => s.path !== slide)!.path;
    expect(slideComments(doc, copy)).toEqual([]);
    expect(doc.xml(copy)!.getElementsByTagNameNS(P188, "commentRel")).toHaveLength(0);
    addSlideComment(doc, copy, "Copy's own", "Ada");
    expect(slideComments(doc, slide).map((t) => t.messages[0].text)).toEqual(["Only here"]);
  });

  it("stamps comments with local time (no zone), as PowerPoint writes and reads them", async () => {
    const { doc, slide } = await deck();
    addSlideComment(doc, slide, "Now", "Ada");
    const date = slideComments(doc, slide)[0].messages[0].date;
    expect(date).not.toMatch(/Z$/);
    expect(Math.abs(new Date(date).getTime() - Date.now())).toBeLessThan(60_000);
  });
});
