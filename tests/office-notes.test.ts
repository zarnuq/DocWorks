import { afterEach, describe, expect, it, vi } from "vitest";
import type { App as RealApp, TFile as RealFile } from "obsidian";
import { App, TFile } from "./obsidian-stub";
import { openCompanionNote, createSelectionNote, noteText } from "../src/presentation/common/office-notes";

vi.mock("obsidian", async (original) => ({ ...await original<object>(), parseYaml: (s: string) => {
  const out: Record<string, unknown> = {};
  for (const line of s.split("\n")) {
    const m = /^([^:]+):\s*(.*)$/.exec(line);
    if (m) { try { out[m[1]] = JSON.parse(m[2]); } catch { out[m[1]] = m[2]; } }
  }
  return out;
} }));

function vault() {
  const app = new App();
  const files = new Map<string, { file: TFile; text: string }>();
  const opened: TFile[] = [];
  Object.assign(app.workspace, { getLeaf: () => ({ openFile: async (file: TFile) => { opened.push(file); } }) });
  app.vault.getFiles = () => [...files.values()].map(v => v.file);
  app.vault.getAbstractFileByPath = p => files.get(p)?.file ?? null;
  app.vault.cachedRead = async f => files.get(f.path)!.text;
  app.vault.create = vi.fn(async (path, text) => {
    if (files.has(path)) throw Error("Already exists");
    const file = new TFile(path); files.set(path, { file, text }); return file;
  });
  app.metadataCache.getFirstLinkpathDest = (link, from) => files.get(link)?.file ?? files.get(from.slice(0, from.lastIndexOf("/") + 1) + link)?.file ?? null;
  const source = new TFile("Projects/Report.docx");
  files.set(source.path, { file: source, text: "" });
  return { app: app as unknown as RealApp, source: source as unknown as RealFile, files, opened };
}
afterEach(() => vi.restoreAllMocks());

describe("Office companion notes", () => {
  it("creates beside the source with editable properties and opens a new tab", async () => {
    const { app, source, files, opened } = vault();
    const note = await openCompanionNote(app, source);
    expect(note?.path).toBe("Projects/Report.docx notes.md");
    expect(files.get(note!.path)!.text).toContain('source: "[[Report.docx]]"');
    expect(files.get(note!.path)!.text).toContain('tags: []\nproject: ""\nstatus: ""');
    expect(opened).toEqual([note]);
  });
  it("reopens without overwriting commentary, even before metadata is cached", async () => {
    const { app, source, files } = vault();
    const first = await openCompanionNote(app, source);
    files.get(first!.path)!.text += "My commentary\n";
    expect(await openCompanionNote(app, source)).toBe(first);
    expect(app.vault.create).toHaveBeenCalledTimes(1);
    expect(files.get(first!.path)!.text).toContain("My commentary");
  });
  it("finds a moved companion through its source property", async () => {
    const { app, source, files } = vault();
    const moved = new TFile("Other/My research.md");
    files.set(moved.path, { file: moved, text: "" });
    app.metadataCache.getFileCache = f => f.path === moved.path ? { frontmatter: { "doc-works": "companion", source: "[[Projects/Report.docx]]" } } : null;
    expect(await openCompanionNote(app, source)).toBe(moved);
    expect(app.vault.create).not.toHaveBeenCalled();
  });
  it("does not mistake an unrelated same-named note for a companion", async () => {
    const { app, source, files } = vault();
    files.set("Projects/Report.docx notes.md", { file: new TFile("Projects/Report.docx notes.md"), text: "Private notes" });
    expect((await openCompanionNote(app, source))?.path).toBe("Projects/Report.docx notes 2.md");
    expect(files.get("Projects/Report.docx notes.md")!.text).toBe("Private notes");
  });
  it("coalesces simultaneous clicks into one creation", async () => {
    const { app, source } = vault();
    const [a, b] = await Promise.all([openCompanionNote(app, source), openCompanionNote(app, source)]);
    expect(a).toBe(b);
    expect(app.vault.create).toHaveBeenCalledTimes(1);
  });
  it("recognizes Markdown source links when wikilinks are disabled", async () => {
    const { app, source, files } = vault();
    const note = new TFile("Projects/Research.md");
    files.set(note.path, { file: note, text: "" });
    app.metadataCache.getFileCache = () => ({ frontmatter: { "doc-works": "companion", source: "[Report](Report.docx)" } });
    expect(await openCompanionNote(app, source)).toBe(note);
  });
  it("keeps companions for different Office formats separate", async () => {
    const { app, source, files } = vault();
    const slides = new TFile("Projects/Report.pptx");
    files.set(slides.path, { file: slides, text: "" });
    const a = await openCompanionNote(app, source);
    const b = await openCompanionNote(app, slides as unknown as RealFile);
    expect(a?.path).toBe("Projects/Report.docx notes.md");
    expect(b?.path).toBe("Projects/Report.pptx notes.md");
  });
});

describe("selection notes", () => {
  it("creates independent notes with the source range, without overwriting", async () => {
    const { app, source, files } = vault();
    const a = await createSelectionNote(app, source, "#Results", "Selected text");
    const b = await createSelectionNote(app, source, "#Results", "Other text");
    expect(a!.path).not.toBe(b!.path);
    expect(files.get(a!.path)!.text).toContain("[[Report.docx#Results]]");
    expect(files.get(a!.path)!.text).toContain("Selected text");
  });
  it("refuses empty selections and catches vault errors", async () => {
    const { app, source } = vault();
    expect(await createSelectionNote(app, source, "", "  ")).toBeNull();
    vi.spyOn(app.vault, "create").mockRejectedValue(Error("Disk full"));
    expect(await openCompanionNote(app, source)).toBeNull();
  });
  it("preserves literal text without turning it into Markdown or embeds", () => {
    expect(noteText("# Title\n\n![[other]] *literal*")).toBe("\\# Title\n\n\\!\\[\\[other\\]\\] \\*literal\\*");
    expect(noteText("1. First\n- Second\n---")).toBe("1\\. First\n\n\\- Second\n\n\\-\\-\\-");
  });
});
