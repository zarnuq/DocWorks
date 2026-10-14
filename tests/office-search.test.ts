import { afterEach, describe, expect, it, vi } from "vitest";
import type { WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { OfficeSearchView } from "../src/presentation/common/office-search";
import { fixture } from "./helpers";

afterEach(() => document.body.empty());

function setup(paths: string[]) {
  const app = new App();
  const files = paths.map((p) => new TFile(p));
  app.vault.getFiles = () => files as never;
  const reads = vi.fn(async (f: TFile) => fixture(f.path.endsWith(".pptx") ? "sample.pptx" : "sample.docx").slice().buffer);
  app.vault.readBinary = reads as never;
  const opened: string[] = [];
  const leaf = { openFile: async (f: TFile) => void opened.push(f.path), view: { openFind: vi.fn() } };
  (app.workspace as unknown as { getLeaf: () => typeof leaf }).getLeaf = () => leaf;
  const view = new OfficeSearchView(new WorkspaceLeaf(app) as unknown as RealLeaf, () => ["docx", "pptx"]);
  document.body.appendChild(view.containerEl);
  return { view, reads, opened, leaf };
}

describe("search in Office files", () => {
  it("lists matches by file with their place, skipping other file types", async () => {
    const { view } = setup(["a/deck.pptx", "notes.md", "report.docx", "book.xlsx"]);
    await view.onOpen();
    await view.runSearch("bullets");
    const files = Array.from(view.contentEl.querySelectorAll(".dw-search-file"));
    expect(files.map((f) => f.querySelector(".dw-search-name")!.textContent)).toEqual(["deck.pptx"]);
    const hit = files[0].querySelector(".dw-search-hit")!;
    expect(hit.querySelector(".dw-search-where")!.textContent).toBe("Slide 2");
    expect(hit.querySelector("mark")!.textContent).toBe("Bullets");
    expect(view.contentEl.querySelector(".dw-search-status")!.textContent).toBe("1 result in 1 file");
  });

  it("reads each file once while it's unchanged", async () => {
    const { view, reads } = setup(["report.docx"]);
    await view.onOpen();
    await view.runSearch("heading");
    await view.runSearch("plain");
    expect(reads).toHaveBeenCalledTimes(1);
    expect(view.contentEl.querySelector(".dw-search-status")!.textContent).toBe("1 result in 1 file");
  });

  it("opens the file with its find bar on the search", async () => {
    const { view, opened, leaf } = setup(["report.docx"]);
    await view.onOpen();
    await view.runSearch("heading");
    view.contentEl.querySelector<HTMLElement>(".dw-search-hit")!.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(opened).toEqual(["report.docx"]);
    expect(leaf.view.openFind).toHaveBeenCalled();
  });

  it("leaves Obsidian's View.open alone (overriding it stops the panel from ever drawing)", () => {
    expect(Object.getOwnPropertyNames(OfficeSearchView.prototype)).not.toContain("open");
  });

  it("says when nothing matches", async () => {
    const { view } = setup(["report.docx"]);
    await view.onOpen();
    await view.runSearch("zzz");
    expect(view.contentEl.querySelector(".dw-search-status")!.textContent).toBe("No results");
  });
});
