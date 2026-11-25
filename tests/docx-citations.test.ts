import { describe, expect, it } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";
import {
  bibliographyEntry, bibliographyTitle, citationText, newTag, parseAuthors, readSources, sortSources, writeSources,
  type Source, type CitationStyle,
} from "../src/domain/docx/citations";
import { chipAttrs, citationChip, type ChipAttrs } from "../src/domain/docx/chips";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { attrsOf, docxSchema as schema } from "../src/domain/docx/schema";
import { cmd } from "../src/presentation/docx/docx-editor";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { serializeXml, W } from "./helpers";

const B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

const book: Source = { tag: "Smi20", type: "book", authors: [{ first: "Ann", last: "Smith" }], title: "The Long Road", year: "2020", publisher: "Penguin", container: "", volume: "", issue: "", pages: "", url: "" };
const two: Source = { ...book, tag: "Lee19", authors: [{ first: "Kim", last: "Lee" }, { first: "Jo", last: "Park" }], title: "Two Minds", year: "2019" };
const three: Source = { ...book, tag: "Ado21", authors: [{ first: "Al", last: "Adams" }, { first: "Bo", last: "Brown" }, { first: "Cy", last: "Clark" }], title: "Three Ways", year: "2021" };
const journal: Source = { tag: "Doe18", type: "journal", authors: [{ first: "Jane", last: "Doe" }], title: "On Rivers", year: "2018", publisher: "", container: "Water Studies", volume: "12", issue: "3", pages: "45-67", url: "" };
const site: Source = { tag: "Wik22", type: "website", authors: [], title: "Obsidian Help", year: "2022", publisher: "", container: "Obsidian", volume: "", issue: "", pages: "", url: "https://help.obsidian.md" };
const undated: Source = { ...book, tag: "Nod", year: "" };

const plain = (segments: { text: string; italic?: boolean }[]) => segments.map((s) => (s.italic ? `*${s.text}*` : s.text)).join("");

describe("in-text citations", () => {
  const cases: [CitationStyle, Source, string][] = [
    ["apa", book, "(Smith, 2020)"],
    ["apa", two, "(Lee & Park, 2019)"],
    ["apa", three, "(Adams et al., 2021)"],
    ["apa", undated, "(Smith, n.d.)"],
    ["apa", site, "(Obsidian Help, 2022)"],
    ["mla", book, "(Smith)"],
    ["mla", two, "(Lee and Park)"],
    ["mla", three, "(Adams et al.)"],
    ["mla", site, "(\"Obsidian Help\")"],
    ["chicago", book, "(Smith 2020)"],
    ["chicago", two, "(Lee and Park 2019)"],
    ["chicago", three, "(Adams, Brown, and Clark 2021)"],
    ["chicago", { ...three, authors: [...three.authors, { first: "Di", last: "Dunn" }] }, "(Adams et al. 2021)"],
    ["chicago", undated, "(Smith n.d.)"],
  ];
  for (const [style, source, expected] of cases) {
    it(`${style}: ${expected}`, () => expect(citationText(source, style)).toBe(expected));
  }
});

describe("bibliography entries", () => {
  const cases: [CitationStyle, Source, string][] = [
    ["apa", book, "Smith, A. (2020). *The Long Road*. Penguin."],
    ["apa", two, "Lee, K., & Park, J. (2019). *Two Minds*. Penguin."],
    ["apa", three, "Adams, A., Brown, B., & Clark, C. (2021). *Three Ways*. Penguin."],
    ["apa", journal, "Doe, J. (2018). On Rivers. *Water Studies*, *12*(3), 45-67."],
    ["apa", site, "*Obsidian Help*. (2022). Obsidian. https://help.obsidian.md"],
    ["apa", undated, "Smith, A. (n.d.). *The Long Road*. Penguin."],
    ["mla", book, "Smith, Ann. *The Long Road*. Penguin, 2020."],
    ["mla", two, "Lee, Kim, and Jo Park. *Two Minds*. Penguin, 2019."],
    ["mla", three, "Adams, Al, et al. *Three Ways*. Penguin, 2021."],
    ["mla", journal, "Doe, Jane. \"On Rivers.\" *Water Studies*, vol. 12, no. 3, 2018, pp. 45-67."],
    ["mla", site, "\"Obsidian Help.\" *Obsidian*, 2022, https://help.obsidian.md."],
    ["chicago", book, "Smith, Ann. 2020. *The Long Road*. Penguin."],
    ["chicago", two, "Lee, Kim, and Jo Park. 2019. *Two Minds*. Penguin."],
    ["chicago", three, "Adams, Al, Bo Brown, and Cy Clark. 2021. *Three Ways*. Penguin."],
    ["chicago", journal, "Doe, Jane. 2018. \"On Rivers.\" *Water Studies* 12 (3): 45-67."],
    ["chicago", site, "\"Obsidian Help.\" 2022. Obsidian. https://help.obsidian.md."],
  ];
  for (const [style, source, expected] of cases) {
    it(`${style} ${source.type}: ${expected}`, () => expect(plain(bibliographyEntry(source, style))).toBe(expected));
  }

  it("titles of the list follow the style", () => {
    expect(bibliographyTitle("apa")).toBe("References");
    expect(bibliographyTitle("mla")).toBe("Works Cited");
    expect(bibliographyTitle("chicago")).toBe("Bibliography");
  });

  it("sorts by first author's last name, then by title for sources without authors", () => {
    expect(sortSources([book, site, journal, three, two]).map((s) => s.tag)).toEqual(["Ado21", "Doe18", "Lee19", "Wik22", "Smi20"]);
  });
});

describe("parseAuthors and newTag", () => {
  it("reads one author per line as 'First Last' or 'Last, First'", () => {
    expect(parseAuthors("Ann Smith\nLee, Kim\n\n  Mary Jane Watson ")).toEqual([
      { first: "Ann", last: "Smith" }, { first: "Kim", last: "Lee" }, { first: "Mary Jane", last: "Watson" },
    ]);
  });

  it("a single name is a last name", () => {
    expect(parseAuthors("Plato")).toEqual([{ first: "", last: "Plato" }]);
  });

  it("makes Word's tags (three letters of the name, two of the year), unique", () => {
    expect(newTag({ ...book, tag: "" }, [])).toBe("Smi20");
    expect(newTag({ ...book, tag: "" }, ["Smi20"])).toBe("Smi201");
    expect(newTag({ ...site, tag: "" }, [])).toBe("Obs22");
    expect(newTag({ ...undated, tag: "" }, [])).toBe("Smi");
  });
});

describe("sources XML (Word's bibliography part)", () => {
  it("round-trips every modelled field", () => {
    const xml = writeSources(null, "mla", [book, two, journal, site]);
    const read = readSources(new DOMParser().parseFromString(xml, "application/xml"));
    expect(read.style).toBe("mla");
    expect(read.sources).toEqual([book, two, journal, site]);
  });

  it("writes Word's element names and source types", () => {
    const doc = new DOMParser().parseFromString(writeSources(null, "apa", [journal, site]), "application/xml");
    expect(doc.documentElement.namespaceURI).toBe(B);
    expect(doc.documentElement.getAttribute("StyleName")).toBe("APA");
    const types = Array.from(doc.getElementsByTagNameNS(B, "SourceType")).map((t) => t.textContent);
    expect(types).toEqual(["JournalArticle", "InternetSite"]);
    expect(doc.getElementsByTagNameNS(B, "JournalName")[0].textContent).toBe("Water Studies");
    expect(doc.getElementsByTagNameNS(B, "InternetSiteTitle")[0].textContent).toBe("Obsidian");
    const person = doc.getElementsByTagNameNS(B, "Person")[0];
    expect(person.parentElement!.localName).toBe("NameList");
    expect(person.getElementsByTagNameNS(B, "Last")[0].textContent).toBe("Doe");
  });

  it("keeps fields it doesn't model when a source is edited", () => {
    const original = writeSources(null, "apa", [book]).replace("</b:Source>", "<b:City>London</b:City><b:Guid>{X}</b:Guid></b:Source>");
    const doc = new DOMParser().parseFromString(original, "application/xml");
    const out = new DOMParser().parseFromString(writeSources(doc, "apa", [{ ...book, title: "New Title" }]), "application/xml");
    expect(out.getElementsByTagNameNS(B, "City")[0].textContent).toBe("London");
    expect(out.getElementsByTagNameNS(B, "Guid")[0].textContent).toBe("{X}");
    expect(out.getElementsByTagNameNS(B, "Title")[0].textContent).toBe("New Title");
  });

  it("reads Word's other source types as the nearest of ours, and Corporate authors", () => {
    const xml = `<b:Sources xmlns:b="${B}" SelectedStyle="\\APASixthEdition.xsl" StyleName="APA"><b:Source><b:Tag>Rep</b:Tag><b:SourceType>Report</b:SourceType><b:Author><b:Author><b:Corporate>WHO</b:Corporate></b:Author></b:Author><b:Title>Health</b:Title><b:Year>2015</b:Year></b:Source></b:Sources>`;
    const read = readSources(new DOMParser().parseFromString(xml, "application/xml"));
    expect(read.sources[0]).toMatchObject({ tag: "Rep", type: "book", authors: [{ first: "", last: "WHO" }], title: "Health", year: "2015" });
  });

  it("an empty or missing part has no sources and APA style", () => {
    expect(readSources(null)).toEqual({ style: "apa", sources: [] });
  });
});

describe("the sources part in a document", () => {
  it("is created as customXml with its item properties, and found again after saving", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    expect(docx.sources()).toEqual({ style: "apa", sources: [] });
    docx.setSources("chicago", [book]);
    const reopened = await DocxDocument.load(await docx.save());
    expect(reopened.sources()).toEqual({ style: "chicago", sources: [book] });
    const paths = reopened.partPaths();
    const item = paths.find((p) => /^customXml\/item\d+\.xml$/.test(p))!;
    expect(item).toBeDefined();
    expect(paths).toContain(item.replace("item", "itemProps"));
  });

  it("is rewritten in place on later edits (one part only)", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    docx.setSources("apa", [book]);
    docx.setSources("apa", [book, journal]);
    const reopened = await DocxDocument.load(await docx.save());
    expect(reopened.sources().sources).toHaveLength(2);
    expect(reopened.partPaths().filter((p) => /^customXml\/item\d+\.xml$/.test(p))).toHaveLength(1);
  });
});

// --- citations in the text ---------------------------------------------------

const ctx = { imageSrc: () => null, hyperlinkHref: () => null };
function body(inner: string): Element {
  return new DOMParser().parseFromString(`<w:document xmlns:w="${W}"><w:body>${inner}<w:sectPr/></w:body></w:document>`, "application/xml").getElementsByTagNameNS(W, "body")[0];
}
const WORD_CITATION =
  `<w:sdt><w:sdtPr><w:id w:val="9"/><w:citation/></w:sdtPr><w:sdtContent>` +
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> CITATION Smi20 \\l 1033 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:rPr><w:noProof/></w:rPr><w:t>(Smith, 2020)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:sdt>`;
const chipsIn = (doc: PMNode) => {
  const out: { pos: number; node: PMNode }[] = [];
  doc.descendants((node, pos) => void (node.type === schema.nodes.chip && out.push({ pos, node })));
  return out;
};

describe("citation chips", () => {
  it("Word's CITATION content control reads as a citation chip with its tag", () => {
    const { doc } = readDocx(body(`<w:p><w:r><w:t xml:space="preserve">As shown </w:t></w:r>${WORD_CITATION}</w:p>`), ctx);
    const [c] = chipsIn(doc);
    expect(attrsOf<ChipAttrs>(c.node)).toMatchObject({ kind: "citation", tag: "Smi20", text: "(Smith, 2020)" });
    expect(c.node.marks.some((m) => m.type === schema.marks.noProof)).toBe(true);
    expect(doc.textContent).toBe("As shown (Smith, 2020)");
  });

  it("an unchanged one is written back byte-identical", () => {
    const b = body(`<w:p>${WORD_CITATION}</w:p>`);
    const before = serializeXml(b);
    const { doc, origins } = readDocx(b, ctx);
    writeBody(b, doc, origins, { nextDocPrId: () => 1 });
    expect(serializeXml(b)).toBe(before);
  });

  it("a new one is a Word citation control around a CITATION field with the cached text", () => {
    const b = body(`<w:p/>`);
    const { doc, origins } = readDocx(b, ctx);
    const chip = schema.nodes.chip.create(citationChip(book, "apa"));
    writeBody(b, doc.copy(doc.content.replaceChild(0, schema.nodes.paragraph.create(null, chip))), origins, { nextDocPrId: () => 1 });
    const sdt = b.getElementsByTagNameNS(W, "sdt")[0];
    expect(Array.from(sdt.getElementsByTagNameNS(W, "sdtPr")[0].children).map((c) => c.localName)).toEqual(["id", "citation"]);
    const types = Array.from(sdt.getElementsByTagNameNS(W, "fldChar")).map((f) => f.getAttributeNS(W, "fldCharType"));
    expect(types).toEqual(["begin", "separate", "end"]);
    expect(sdt.getElementsByTagNameNS(W, "instrText")[0].textContent).toBe(" CITATION Smi20 \\l 1033 ");
    expect(sdt.getElementsByTagNameNS(W, "t")[0].textContent).toBe("(Smith, 2020)");
    // And it reads back the same.
    expect(chipAttrs(sdt)).toMatchObject({ kind: "citation", tag: "Smi20", text: "(Smith, 2020)" });
  });

  it("a restyled one keeps its field, with the new cached text", () => {
    const b = body(`<w:p>${WORD_CITATION}</w:p>`);
    const { doc, origins } = readDocx(b, ctx);
    const [c] = chipsIn(doc);
    const edited = EditorState.create({ doc }).tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, text: "(Smith)" }).doc;
    writeBody(b, edited, origins, { nextDocPrId: () => 1 });
    const sdt = b.getElementsByTagNameNS(W, "sdt")[0];
    expect(sdt.getElementsByTagNameNS(W, "id")[0].getAttributeNS(W, "val")).toBe("9");
    expect(sdt.getElementsByTagNameNS(W, "instrText")[0].textContent).toContain("CITATION Smi20");
    expect(sdt.getElementsByTagNameNS(W, "t")[0].textContent).toBe("(Smith)");
  });
});

describe("citation commands", () => {
  const start = (text = "Fact. ") => {
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text(text)), schema.nodes.paragraph.create(null, schema.text("End."))]);
    return EditorState.create({ doc, selection: TextSelection.create(doc, text.length + 1) });
  };
  const apply = (state: EditorState, c: (s: EditorState, d?: (tr: import("prosemirror-state").Transaction) => void) => boolean) => {
    let next = state;
    expect(c(state, (tr) => (next = state.apply(tr)))).toBe(true);
    return next;
  };
  const bibliography = (doc: PMNode) => {
    let found: PMNode | null = null;
    doc.descendants((n) => {
      if (n.type === schema.nodes.sdt && n.attrs.kind === "bibliography") found = n;
      return !found;
    });
    return found as PMNode | null;
  };

  it("insertBibliography adds a titled, sorted, hanging-indented list after the caret's paragraph", () => {
    const s = apply(start(), cmd.insertBibliography([book, journal], "apa"));
    const bib = bibliography(s.doc)!;
    expect(s.doc.child(1)).toBe(bib);
    expect(bib.child(0).textContent).toBe("References");
    expect(bib.child(0).attrs.style).toBe("Heading1");
    expect(bib.child(1).textContent).toBe("Doe, J. (2018). On Rivers. Water Studies, 12(3), 45-67.");
    expect(bib.child(2).textContent).toBe("Smith, A. (2020). The Long Road. Penguin.");
    expect(bib.child(1).attrs).toMatchObject({ indent: 1, firstLine: -36 });
    const italic = bib.child(2).childCount && [] as string[];
    bib.child(2).forEach((n) => n.marks.some((m) => m.type === schema.marks.italic) && (italic as string[]).push(n.text!));
    expect(italic).toEqual(["The Long Road"]);
  });

  it("the list is a Word BIBLIOGRAPHY field inside a Bibliographies content control", () => {
    const s = apply(start(), cmd.insertBibliography([book], "mla"));
    const bib = bibliography(s.doc)!;
    expect(bib.attrs.xml).toContain('w:docPartGallery w:val="Bibliographies"');
    const xmls: string[] = [];
    bib.descendants((n) => void (n.type === schema.nodes.opaque && xmls.push(n.attrs.xml as string)));
    expect(xmls.join("")).toMatch(/fldCharType="begin".*BIBLIOGRAPHY.*fldCharType="separate".*fldCharType="end"/s);
  });

  it("refuses with no sources", () => {
    expect(cmd.insertBibliography([], "apa")(start())).toBe(false);
  });

  it("restyleCitations rewrites every citation chip and the bibliography for another style", () => {
    let s = apply(start(), cmd.insertChip(citationChip(book, "apa")));
    s = apply(s, cmd.insertBibliography([book], "apa"));
    s = apply(s, cmd.restyleCitations([book], "mla"));
    expect(chipsIn(s.doc).map((c) => c.node.attrs.text)).toEqual(["(Smith)"]);
    const bib = bibliography(s.doc)!;
    expect(bib.child(0).textContent).toBe("Works Cited");
    expect(bib.child(1).textContent).toBe("Smith, Ann. The Long Road. Penguin, 2020.");
  });

  it("restyleCitations leaves citations of unknown sources as they are", () => {
    const s = apply(start(), cmd.insertChip(citationChip(book, "apa")));
    expect(cmd.restyleCitations([journal], "mla")(s)).toBe(false);
    expect(chipsIn(s.doc)[0].node.attrs.text).toBe("(Smith, 2020)");
  });

  it("restyleCitations with nothing to change changes nothing", () => {
    expect(cmd.restyleCitations([book], "apa")(start())).toBe(false);
  });
});
