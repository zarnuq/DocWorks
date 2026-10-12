import { afterEach, expect, it, vi } from "vitest";
import * as obsidian from "obsidian";
import { renderNoteForExport } from "../src/presentation/common/render-note-export";

afterEach(() => vi.restoreAllMocks());
const file = (path: string) => Object.assign(new obsidian.TFile(), { path });
function setup(notes: Record<string, { markdown: string; html: string }>) {
  const files = new Map(Object.keys(notes).map(path => [path, file(path)]));
  const read = vi.fn(async (f: obsidian.TFile) => notes[f.path].markdown);
  const app = {
    vault: { cachedRead: read, getName: () => "My vault" },
    metadataCache: {
      getFirstLinkpathDest: vi.fn((path: string, source: string) => files.get(path) ?? files.get(source.slice(0, source.lastIndexOf("/") + 1) + path) ?? null),
      getFileCache: () => ({}),
    },
  } as unknown as obsidian.App;
  const render = vi.spyOn(obsidian.MarkdownRenderer, "render").mockImplementation(async (_app, _markdown, el, source) => { el.innerHTML = notes[source].html; });
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
