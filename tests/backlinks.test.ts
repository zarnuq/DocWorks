import { afterEach, describe, expect, it } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { fixture } from "./helpers";
import { sheetXml, xlsx } from "./xlsx-fixtures";
import { backlinksTo, groupBacklinks, placeLabel, renderBacklinks } from "../src/presentation/common/backlinks";

afterEach(() => document.body.empty());

/** Obsidian's link index over `notes` (path -> text); links resolve to `targets` by path. */
function linkIndex(notes: Record<string, string>, targets: TFile[]) {
  const all = new Map<string, TFile>([...Object.keys(notes).map((p) => new TFile(p)), ...targets].map((f) => [f.path, f]));
  const linkRe = /!?\[\[([^\]|]+)(?:\|[^\]]*)?\]\]/g;
  const parsed = (path: string) => {
    const links: { link: string; original: string; position: { start: { line: number } } }[] = [];
    notes[path].split("\n").forEach((text, line) => {
      for (const m of text.matchAll(linkRe)) links.push({ link: m[1], original: m[0], position: { start: { line } } });
    });
    return links;
  };
  const dest = (link: string) => all.get(link.split("#")[0]) ?? null;
  const resolvedLinks: Record<string, Record<string, number>> = {};
  for (const p of Object.keys(notes)) {
    resolvedLinks[p] = {};
    for (const l of parsed(p)) {
      const d = dest(l.link);
      if (d) resolvedLinks[p][d.path] = (resolvedLinks[p][d.path] ?? 0) + 1;
    }
  }
  return {
    vault: { getAbstractFileByPath: (p: string) => all.get(p) ?? null, cachedRead: async (f: TFile) => notes[f.path] },
    metadataCache: {
      resolvedLinks,
      getFirstLinkpathDest: (link: string) => dest(link),
      getFileCache: (f: TFile) => {
        const links = parsed(f.path);
        return { links: links.filter((l) => !l.original.startsWith("!")), embeds: links.filter((l) => l.original.startsWith("!")) };
      },
      on: () => ({}),
    },
  };
}

function vault(notes: Record<string, string>) {
  const targets = ["deck.pptx", "Budget.xlsx", "report.docx"].map((p) => new TFile(p));
  return { app: linkIndex(notes, targets) as never, file: (p: string) => targets.find((t) => t.path === p) as unknown as RealTFile };
}

describe("backlinks into an Office file", () => {
  it("finds every note linking to the file, with the place, line and the line's text", async () => {
    const { app, file } = vault({
      "Field Notes.md": "intro\nsee [[deck.pptx#3]] for the map\n![[deck.pptx#3]]",
      "Index.md": "- [[deck.pptx]]\n- [[report.docx#Method]]",
    });
    const links = await backlinksTo(app, file("deck.pptx"));
    expect(links.map((l) => [l.source.path, l.subpath, l.line, l.text])).toEqual([
      ["Field Notes.md", "#3", 1, "see [[deck.pptx#3]] for the map"],
      ["Field Notes.md", "#3", 2, "![[deck.pptx#3]]"],
      ["Index.md", "", 0, "- [[deck.pptx]]"],
    ]);
  });

  it("groups by place in document order, the whole file first", async () => {
    const { app, file } = vault({
      "a.md": "[[deck.pptx#10]] [[deck.pptx#2]]",
      "b.md": "[[deck.pptx]] [[deck.pptx#2]]",
    });
    const groups = groupBacklinks("pptx", await backlinksTo(app, file("deck.pptx")));
    expect(groups.map((g) => [g.label, g.links.map((l) => l.source.path)])).toEqual([
      ["Whole file", ["b.md"]],
      ["Slide 2", ["a.md", "b.md"]],
      ["Slide 10", ["a.md"]],
    ]);
  });

  it("names places as each editor shows them", () => {
    expect(placeLabel("pptx", "#4")).toBe("Slide 4");
    expect(placeLabel("xlsx", "#Budget!B4:C8")).toBe("Budget!B4:C8");
    expect(placeLabel("docx", "#Results#Totals")).toBe("Results › Totals");
    expect(placeLabel("docx", "")).toBe("Whole file");
  });

  it("draws the panel: a place jumps there, a note opens at its line", async () => {
    const { app, file } = vault({ "a.md": "x\n[[deck.pptx#2]]", "b.md": "[[deck.pptx]]" });
    const groups = groupBacklinks("pptx", await backlinksTo(app, file("deck.pptx")));
    const went: string[] = [], opened: [string, number][] = [];
    const panel = document.body.createDiv();
    renderBacklinks(panel, groups, { goTo: (s) => went.push(s), open: (l) => opened.push([l.source.path, l.line]) });
    const places = Array.from(panel.querySelectorAll<HTMLButtonElement>(".dw-backlinks-place"));
    expect(places.map((p) => p.textContent)).toEqual(["Whole file", "Slide 2"]);
    places[1].click();
    expect(went).toEqual(["#2"]);
    panel.querySelectorAll<HTMLElement>(".dw-backlinks-note")[1].click();
    expect(opened).toEqual([["a.md", 1]]);
  });

  it("says so when nothing links here", () => {
    const panel = document.body.createDiv();
    renderBacklinks(panel, [], { goTo: () => {}, open: () => {} });
    expect(panel.textContent).toContain("No notes link to this file yet.");
  });
});

describe("Linked from, in the editors", () => {
  async function mount<V extends DocxView | PptxView | XlsxView>(View: new (leaf: RealLeaf, settings: typeof DEFAULT_SETTINGS) => V, bytes: Uint8Array, path: string, notes: Record<string, string>) {
    const app = new App();
    const file = new TFile(path);
    const index = linkIndex(notes, [file]);
    Object.assign(app.vault, index.vault, { readBinary: async () => bytes.slice().buffer });
    Object.assign(app, { metadataCache: index.metadataCache });
    const view = new View(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
    document.body.appendChild(view.containerEl);
    view.file = file as unknown as RealTFile;
    await view.onLoadFile(file as unknown as RealTFile);
    await new Promise((r) => setTimeout(r, 0));
    return view;
  }
  const openPanel = (view: { contentEl: HTMLElement }) => {
    view.contentEl.querySelector<HTMLButtonElement>('.dw-menubar-end [aria-label="Linked from"]')!.click();
    return new Promise((r) => setTimeout(r, 0));
  };

  it("slides: badges on linked thumbnails, and the panel jumps to a slide", async () => {
    const view = await mount(PptxView, fixture("sample.pptx"), "deck.pptx", { "a.md": "[[deck.pptx#2]]", "b.md": "![[deck.pptx#2]]\n[[deck.pptx]]" });
    const badges = Array.from(view.contentEl.querySelectorAll<HTMLElement>(".pptx-thumb .dw-linked-badge"));
    expect(badges.map((b) => [b.closest(".pptx-thumb")!.getAttribute("aria-label"), b.textContent])).toEqual([["Slide 2", "2"]]);
    await openPanel(view);
    const places = Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-backlinks-place"));
    expect(places.map((p) => p.textContent)).toEqual(["Whole file", "Slide 2"]);
    places[1].click();
    expect((view as unknown as { index: number }).index).toBe(1);
  });

  it("sheets: linked ranges on the shown sheet are outlined", async () => {
    const view = await mount(XlsxView, await xlsx({ sheets: [sheetXml(""), sheetXml("")], names: ["One", "Two"] }), "Budget.xlsx", { "a.md": "[[Budget.xlsx#One!B2:C3]] [[Budget.xlsx#Two!A1]]" });
    const boxes = view.contentEl.querySelectorAll(".xlsx-linked");
    expect(boxes).toHaveLength(1);
  });

  it("documents: linked headings are marked in the outline", async () => {
    const view = await mount(DocxView, fixture("sample.docx"), "report.docx", { "a.md": "[[report.docx#Doc Heading]]" });
    runOutline(view);
    const entry = Array.from(view.contentEl.querySelectorAll<HTMLElement>(".docx-outline-entry")).find((e) => e.querySelector(".dw-linked-badge"));
    expect(entry?.title).toBe("Doc Heading");
  });

  it("works on read-only files too", async () => {
    const app = new App();
    const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, readOnly: true });
    app.vault.readBinary = async () => fixture("sample.pptx").slice().buffer;
    document.body.appendChild(view.containerEl);
    const file = new TFile("deck.pptx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    expect(view.contentEl.querySelector('.dw-menubar-end [aria-label="Linked from"]')).not.toBeNull();
  });
});

function runOutline(view: DocxView) {
  (view as unknown as { toggleOutline(): void }).toggleOutline();
}
