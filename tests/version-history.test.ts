import { describe, expect, it, vi } from "vitest";
import { MemoryVersionStore, VersionHistory } from "../src/presentation/common/version-history";

const MIN = 60_000, DAY = 86_400_000, MB = 1024 * 1024;

function setup(opts: { days?: number; enabled?: boolean } = {}) {
  let now = Date.UTC(2026, 9, 14, 12);
  const store = new MemoryVersionStore();
  const history = new VersionHistory(store, () => ({ enabled: opts.enabled ?? true, days: opts.days ?? 14 }), () => now);
  const disk = new Map<string, Uint8Array>();
  const keep = (path: string, size = 10) => history.beforeSave(path, async () => (disk.get(path) ?? new Uint8Array(size)));
  return { history, store, keep, disk, tick: (ms: number) => (now += ms) };
}

describe("version history", () => {
  it("keeps the file as it was before the first save of a session, then at most every 10 minutes", async () => {
    const { history, keep, tick } = setup();
    await keep("a.docx");
    tick(2 * MIN);
    await keep("a.docx");
    expect(await history.list("a.docx")).toHaveLength(1);
    tick(9 * MIN);
    await keep("a.docx");
    expect(await history.list("a.docx")).toHaveLength(2);
  });

  it("lists a file's versions newest first, and gives their bytes back", async () => {
    const { history, keep, tick, disk } = setup();
    disk.set("a.docx", new Uint8Array([1]));
    await keep("a.docx");
    tick(11 * MIN);
    disk.set("a.docx", new Uint8Array([2]));
    await keep("a.docx");
    const list = await history.list("a.docx");
    expect(list.map((v) => v.size)).toEqual([1, 1]);
    expect(list[0].time).toBeGreaterThan(list[1].time);
    expect(Array.from((await history.bytes(list[0].id))!)).toEqual([2]);
  });

  it("keeps 10 versions per file, and 1 of a file over 25 MB", async () => {
    const { history, keep, tick } = setup();
    for (let i = 0; i < 12; i++) {
      await keep("a.docx");
      tick(11 * MIN);
    }
    expect(await history.list("a.docx")).toHaveLength(10);
    for (let i = 0; i < 3; i++) {
      await keep("big.xlsx", 26 * MB);
      tick(11 * MIN);
    }
    expect(await history.list("big.xlsx")).toHaveLength(1);
  });

  it("forgets versions older than the days kept", async () => {
    const { history, keep, tick } = setup({ days: 7 });
    await keep("a.docx");
    tick(8 * DAY);
    await keep("b.docx");
    expect(await history.list("a.docx")).toHaveLength(0);
    expect(await history.list("b.docx")).toHaveLength(1);
  });

  it("keeps all versions together under 200 MB, dropping the oldest", async () => {
    const { history, keep, tick } = setup();
    for (let i = 0; i < 10; i++) {
      await keep(`f${i}.pptx`, 24 * MB);
      tick(MIN);
    }
    const left = (await Promise.all(Array.from({ length: 10 }, (_, i) => history.list(`f${i}.pptx`)))).map((l) => l.length);
    expect(left.reduce((a, b) => a + b)).toBe(8);
    expect(left.slice(0, 2)).toEqual([0, 0]);
  });

  it("keeps nothing when turned off", async () => {
    const { history, keep } = setup({ enabled: false });
    await keep("a.docx");
    expect(await history.list("a.docx")).toHaveLength(0);
  });

  it("follows renames and forgets deleted files", async () => {
    const { history, keep } = setup();
    await keep("a.docx");
    await history.renamed("a.docx", "b.docx");
    expect(await history.list("b.docx")).toHaveLength(1);
    await history.deleted("b.docx");
    expect(await history.list("b.docx")).toHaveLength(0);
  });

  it("a forced keep (before a restore) ignores the 10 minutes", async () => {
    const { history, keep } = setup();
    await keep("a.docx");
    await history.beforeSave("a.docx", async () => new Uint8Array(3), true);
    expect(await history.list("a.docx")).toHaveLength(2);
  });
});

describe("editors keep a version before saving", () => {
  it("a Word document's first save keeps the file as it was on disk", async () => {
    const { App, TFile, WorkspaceLeaf } = await import("./obsidian-stub");
    const { DocxView } = await import("../src/presentation/docx/docx-view");
    const { DEFAULT_SETTINGS } = await import("../src/presentation/common/settings");
    const { setVersionHistory } = await import("../src/presentation/common/version-history");
    const { fixture } = await import("./helpers");
    const history = new VersionHistory(new MemoryVersionStore(), () => ({ enabled: true, days: 14 }));
    setVersionHistory(history);
    try {
      const app = new App();
      const original = fixture("sample.docx");
      app.vault.readBinary = async () => original.slice().buffer;
      const view = new DocxView(new WorkspaceLeaf(app) as never, { ...DEFAULT_SETTINGS });
      document.body.appendChild(view.containerEl);
      const file = new TFile("report.docx");
      view.file = file as never;
      await view.onLoadFile(file as never);
      const editor = (view as unknown as { editor: import("prosemirror-view").EditorView }).editor;
      editor.dispatch(editor.state.tr.insertText("x", 2));
      await view.requestSave();
      const list = await history.list("report.docx");
      expect(list).toHaveLength(1);
      expect((await history.bytes(list[0].id))!.length).toBe(original.length);
    } finally {
      setVersionHistory(null);
      document.body.empty();
    }
  });
});

describe("File › Version history", () => {
  it("lists versions newest first and restores one, keeping the current file first", async () => {
    const { App, TFile } = await import("./obsidian-stub");
    const { setVersionHistory } = await import("../src/presentation/common/version-history");
    const { versionHistoryDialog } = await import("../src/presentation/common/version-dialog");
    let now = Date.UTC(2026, 9, 14, 12);
    const history = new VersionHistory(new MemoryVersionStore(), () => ({ enabled: true, days: 14 }), () => now);
    setVersionHistory(history);
    try {
      const app = new App();
      const file = new TFile("a.docx");
      let disk = new Uint8Array([3]);
      app.vault.readBinary = async () => disk.slice().buffer;
      app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void (disk = new Uint8Array(data));
      await history.beforeSave("a.docx", async () => new Uint8Array([1]));
      now += 11 * 60_000;
      await history.beforeSave("a.docx", async () => new Uint8Array([2]));
      const calls: string[] = [];
      const view = { file, requestSave: async () => void calls.push("save"), onUnloadFile: async () => void calls.push("unload"), onLoadFile: async () => void calls.push("load") };
      await versionHistoryDialog(app as never, view as never, false);
      const rows = Array.from(document.querySelectorAll(".dw-version"));
      expect(rows).toHaveLength(2);
      // The older one ([1]) is second.
      rows[1].querySelector<HTMLButtonElement>("button.mod-cta")!.click();
      await new Promise((r) => setTimeout(r, 0));
      await new Promise((r) => setTimeout(r, 0));
      expect(Array.from(disk)).toEqual([1]);
      expect(calls).toEqual(["save", "unload", "load"]);
      // What was on disk before the restore ([3]) is now the newest version.
      const list = await history.list("a.docx");
      expect(Array.from((await history.bytes(list[0].id))!)).toEqual([3]);
    } finally {
      setVersionHistory(null);
      document.body.empty();
    }
  });

  it("offers no Restore on a read-only file", async () => {
    const { App, TFile } = await import("./obsidian-stub");
    const { setVersionHistory } = await import("../src/presentation/common/version-history");
    const { versionHistoryDialog } = await import("../src/presentation/common/version-dialog");
    const history = new VersionHistory(new MemoryVersionStore(), () => ({ enabled: true, days: 14 }));
    setVersionHistory(history);
    try {
      await history.beforeSave("a.docx", async () => new Uint8Array([1]));
      await versionHistoryDialog(new App() as never, { file: new TFile("a.docx") } as never, true);
      expect(document.querySelectorAll(".dw-version")).toHaveLength(1);
      expect(document.querySelector(".dw-version button.mod-cta")).toBeNull();
    } finally {
      setVersionHistory(null);
      document.body.empty();
    }
  });
});

describe("Version history › Compare", () => {
  async function docx(...paras: string[]): Promise<Uint8Array> {
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    zip.file("word/document.xml", `<w:document xmlns:w="${W}"><w:body>${paras.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`);
    return zip.generateAsync({ type: "uint8array" });
  }

  it("shows what changed since a version, words struck out and added", async () => {
    const { App, TFile } = await import("./obsidian-stub");
    const { setVersionHistory } = await import("../src/presentation/common/version-history");
    const { versionHistoryDialog } = await import("../src/presentation/common/version-dialog");
    const history = new VersionHistory(new MemoryVersionStore(), () => ({ enabled: true, days: 14 }));
    setVersionHistory(history);
    try {
      const old = await docx("Title", "The quick fox", "Gone");
      const now = await docx("Title", "The slow fox", "New end");
      await history.beforeSave("a.docx", async () => old);
      const app = new App();
      app.vault.readBinary = async () => now.slice().buffer;
      const saved: string[] = [];
      const view = { file: new TFile("a.docx"), requestSave: async () => void saved.push("save") };
      await versionHistoryDialog(app as never, view as never, true);
      Array.from(document.querySelectorAll<HTMLButtonElement>(".dw-version button")).find((b) => b.textContent === "Compare")!.click();
      await vi.waitFor(() => expect(document.querySelectorAll(".dw-change")).toHaveLength(2));
      expect(saved).toEqual(["save"]);
      const changes = Array.from(document.querySelectorAll(".dw-change"));
      expect(changes[0].querySelector("del")!.textContent).toBe("quick");
      expect(changes[0].querySelector("ins")!.textContent).toBe("slow");
      expect(changes[1].querySelector("del")!.textContent).toBe("Gone");
      expect(changes[1].querySelector("ins")!.textContent).toBe("New end");
    } finally {
      setVersionHistory(null);
      document.body.empty();
    }
  });

  it("says when the text is the same", async () => {
    const { App, TFile } = await import("./obsidian-stub");
    const { setVersionHistory } = await import("../src/presentation/common/version-history");
    const { versionHistoryDialog } = await import("../src/presentation/common/version-dialog");
    const history = new VersionHistory(new MemoryVersionStore(), () => ({ enabled: true, days: 14 }));
    setVersionHistory(history);
    try {
      const same = await docx("Title");
      await history.beforeSave("a.docx", async () => same);
      const app = new App();
      app.vault.readBinary = async () => same.slice().buffer;
      await versionHistoryDialog(app as never, { file: new TFile("a.docx"), requestSave: async () => {} } as never, true);
      Array.from(document.querySelectorAll<HTMLButtonElement>(".dw-version button")).find((b) => b.textContent === "Compare")!.click();
      await vi.waitFor(() => expect(document.querySelector(".dw-version-preview")!.textContent).toContain("No text changes"));
    } finally {
      setVersionHistory(null);
      document.body.empty();
    }
  });
});
