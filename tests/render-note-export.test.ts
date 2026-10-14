vi.mock("../src/presentation/common/png", () => ({ PNG_WIDTH: 1920, renderPng: vi.fn(async () => new Uint8Array([137, 80, 78, 71])) }));
import { afterEach, expect, it, vi } from "vitest";
import * as obsidian from "obsidian";
import { renderNoteForExport } from "../src/presentation/common/render-note-export";
import { fixture } from "./helpers";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => vi.restoreAllMocks());
const file = (path: string) => Object.assign(new obsidian.TFile(), { path });
function setup(notes: Record<string, { markdown: string; html: string }>, binaries: Record<string, Uint8Array> = {}) {
  const files = new Map([...Object.keys(notes), ...Object.keys(binaries)].map(path => [path, file(path)]));
  const read = vi.fn(async (f: obsidian.TFile) => notes[f.path].markdown);
  const app = {
    vault: { cachedRead: read, getName: () => "My vault", readBinary: async (f: obsidian.TFile) => binaries[f.path].slice().buffer },
    metadataCache: {
      getFirstLinkpathDest: vi.fn((path: string, source: string) => files.get(path) ?? files.get(source.slice(0, source.lastIndexOf("/") + 1) + path) ?? null),
      getFileCache: () => ({}),
    },
  } as unknown as obsidian.App;
  // A Word file's Markdown is shown as paragraphs, its pictures as <img>.
  const render = vi.spyOn(obsidian.MarkdownRenderer, "render").mockImplementation(async (_app, markdown, el, source) => {
    el.innerHTML = notes[source]?.html ?? markdown.split(/\n\n+/).map(b => /^!\[[^\]]*\]\(<?([^>)]+)>?\)$/.exec(b.trim()) ? `<p><img src="${/\(<?([^>)]+)>?\)/.exec(b)![1]}"></p>` : `<p>${b}</p>`).join("");
  });
  const finishMath = vi.spyOn(obsidian, "finishRenderMath").mockResolvedValue();
  return { app, files, read, render, finishMath, component: new obsidian.Component() };
}

it("renders nested note embeds explicitly without waiting for their preview", async () => {
  const t = setup({
    "main.md": { markdown: "Main", html: '<p>Main</p><div class="internal-embed" src="child.md"></div>' },
    "child.md": { markdown: "Child", html: '<h2>Child</h2><div class="internal-embed" src="grandchild.md"></div>' },
    "grandchild.md": { markdown: "Grandchild", html: "<p>Grandchild</p>" },
  });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.textContent).toBe("MainChildGrandchild");
  expect(out.el.querySelector("h2")?.textContent).toBe("Child");
  expect(t.render).toHaveBeenCalledTimes(3);
  expect(t.finishMath).toHaveBeenCalledTimes(3);
  expect(out.warnings).toEqual([]);
});

it("resolves links and pictures relative to the embedded note", async () => {
  const t = setup({
    "main.md": { markdown: "Main", html: '<div class="internal-embed" src="folder/child.md"></div>' },
    "folder/child.md": { markdown: "Child", html: '<p><a class="internal-link" data-href="target.md#Heading" href="target.md#Heading">Alias</a><img src="picture.png"></p>' },
    "folder/target.md": { markdown: "Target", html: "" },
  });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.querySelector("a")?.getAttribute("href")).toBe("obsidian://open?vault=My%20vault&file=folder%2Ftarget.md%23Heading");
  expect(out.el.querySelector("img")?.dataset.exportSource).toBe("folder/child.md");
});

it("limits a heading embed to its requested section using source offsets", async () => {
  const markdown = "---\ntitle: Child\n---\n# Keep\n\nIncluded\n\n# Skip\n\nExcluded";
  const t = setup({
    "main.md": { markdown: "Main", html: '<div class="internal-embed" src="child.md#Keep"></div>' },
    "child.md": { markdown, html: "<h1>Keep</h1><p>Included</p>" },
  });
  vi.spyOn(obsidian, "resolveSubpath").mockReturnValue({ start: { offset: markdown.indexOf("# Keep") }, end: { offset: markdown.indexOf("# Skip") } } as obsidian.HeadingSubpathResult);
  await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(t.render.mock.calls[1][1]).toBe("# Keep\n\nIncluded\n\n");
});

it("exports a referenced block and removes its block ID", async () => {
  const markdown = "Before\n\nBlock text ^chosen\n\nAfter";
  const t = setup({
    "main.md": { markdown: "Main", html: '<div class="internal-embed" src="child.md#^chosen"></div>' },
    "child.md": { markdown, html: "<p>Block text</p>" },
  });
  vi.spyOn(obsidian, "resolveSubpath").mockReturnValue({ start: { offset: markdown.indexOf("Block") }, end: { offset: markdown.indexOf("\n\nAfter") } } as obsidian.BlockSubpathResult);
  await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(t.render.mock.calls[1][1]).toBe("Block text");
});

it("keeps the embed name when its source note is missing", async () => {
  const t = setup({ "main.md": { markdown: "Main", html: '<div class="internal-embed markdown-embed" src="Missing.md" alt="Missing note"><p>Loading preview…</p></div>' } });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.textContent).toBe("Missing note");
  expect(out.warnings).toHaveLength(1);
});

it("strips properties and preserves links to headings in the same note", async () => {
  const t = setup({ "main.md": { markdown: "---\ntags: [test]\n---\n# Heading", html: '<h1>Heading</h1><p><a class="internal-link" data-href="#Heading">Jump</a></p>' } });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(t.render.mock.calls[0][1]).toBe("# Heading");
  expect(out.el.dataset.exportFrontmatter).toBe("---\ntags: [test]\n---\n");
  expect(out.el.querySelector("a")?.getAttribute("href")).toBe("obsidian://open?vault=My%20vault&file=main.md%23Heading");
});

it("preserves properties ending at EOF without rendering them as a divider", async () => {
  const t = setup({ "main.md": { markdown: "---\ntags: [test]\n---", html: "" } });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(t.render.mock.calls[0][1]).toBe("");
  expect(out.el.dataset.exportFrontmatter).toBe("---\ntags: [test]\n---");
});

it("keeps the parent properties without importing the embedded note's properties", async () => {
  const t = setup({
    "main.md": { markdown: "---\ntags: [parent]\n---\n![[child.md]]", html: '<div class="internal-embed" src="child.md"></div>' },
    "child.md": { markdown: "---\ntags: [child]\n---\nChild", html: '<p>Child</p>' },
  });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.dataset.exportFrontmatter).toBe("---\ntags: [parent]\n---\n");
  expect(out.el.querySelector('[data-export-frontmatter]')).toBeNull();
  expect(t.render.mock.calls[1][1]).toBe("Child");
});

it("keeps a link when an embed cycles, and allows repeated independent embeds", async () => {
  const t = setup({
    "main.md": { markdown: "Main", html: '<div class="internal-embed" src="child.md"></div><div class="internal-embed" src="child.md"></div>' },
    "child.md": { markdown: "Child", html: '<p>Child</p><div class="internal-embed" src="main.md"></div>' },
  });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.textContent).toBe("Childmain.mdChildmain.md");
  expect(out.el.querySelectorAll("a")).toHaveLength(2);
  expect(out.warnings).toHaveLength(1);
  expect(t.render).toHaveBeenCalledTimes(3);
});

it("splits a paragraph around an embedded note to retain the note's block structure", async () => {
  const t = setup({
    "main.md": { markdown: "Main", html: '<p>Before <span class="internal-embed" src="child.md"></span> after</p>' },
    "child.md": { markdown: "Child", html: "<h2>Heading</h2><p>Body</p>" },
  });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(Array.from(out.el.children).map(e => e.textContent)).toEqual(["Before ", "HeadingBody", " after"]);
});

it("turns a sheet range embed into a table of its shown values", async () => {
  const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Cost</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Tea</t></is></c><c r="B2"><v>3</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Not asked for</t></is></c></row>`)], names: ["One"] });
  const t = setup({ "main.md": { markdown: "", html: '<p>Before</p><div class="internal-embed" src="book.xlsx#One!A1:B2"></div>' } }, { "book.xlsx": bytes });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  const table = out.el.querySelector("table")!;
  expect(Array.from(table.querySelectorAll("thead th")).map(c => c.textContent)).toEqual(["Item", "Cost"]);
  expect(Array.from(table.querySelectorAll("tbody td")).map(c => c.textContent)).toEqual(["Tea", "3"]);
  expect(out.el.querySelector(".internal-embed")).toBeNull();
});

it("turns a slide embed into a picture of the slide", async () => {
  const t = setup({ "main.md": { markdown: "", html: '<div class="internal-embed" src="deck.pptx#2"></div>' } }, { "deck.pptx": fixture("sample.pptx") });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  const img = out.el.querySelector("img")!;
  expect(img.getAttribute("alt")).toBe("deck, slide 2");
  const pic = out.pictures.get(img)!;
  expect([pic.ext, pic.bytes[0], pic.width > pic.height]).toEqual(["png", 137, true]);
});

it("brings an embedded Word document's text and pictures into the export", async () => {
  globalThis.createImageBitmap = (async () => ({ width: 40, height: 20, close: () => undefined })) as unknown as typeof createImageBitmap;
  const t = setup({ "main.md": { markdown: "", html: '<div class="internal-embed" src="report.docx"></div>' } }, { "report.docx": fixture("sample.docx") });
  const out = await renderNoteForExport(t.app, t.files.get("main.md")!, t.component);
  expect(out.el.textContent).toContain("Doc Heading");
  const img = out.el.querySelector("img")!;
  expect(out.pictures.get(img)).toMatchObject({ ext: "png", width: 40, height: 20 });
});
