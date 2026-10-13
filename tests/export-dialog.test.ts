import { afterEach, describe, expect, it } from "vitest";
import { App, TFile } from "./obsidian-stub";
import { exportDialog, ExportFormat, writeExport } from "../src/presentation/common/export-dialog";

afterEach(() => document.body.empty());

const modal = () => document.querySelector<HTMLElement>(".test-modal")!;
const button = (text: string | RegExp) => Array.from(modal().querySelectorAll("button")).find((b) => (typeof text === "string" ? b.textContent === text : text.test(b.textContent ?? "")))!;
const field = (label: string) => modal().querySelector<HTMLInputElement | HTMLSelectElement>(`[aria-label="${label}"]`)!;
const set = (el: HTMLInputElement | HTMLSelectElement, value: string) => {
  el.value = value;
  el.dispatchEvent(new Event("input"));
  el.dispatchEvent(new Event("change"));
};

function app(existing: string[] = []) {
  const a = new App();
  a.vault.getAbstractFileByPath = (p: string) => (existing.includes(p) ? new TFile(p) : null);
  return a;
}

const formats = (log: string[]): ExportFormat[] => [
  { id: "pdf", label: "PDF document (.pdf)", ext: "pdf", outside: true, options: (el) => void el.createDiv({ text: "pdf options" }) },
  { id: "md", label: "Markdown (.md)", ext: "md", options: (el) => void (log.push("md options"), el.createDiv({ text: "md options" })) },
];

describe("export dialog", () => {
  it("defaults to the first format, the source's name and its folder, and resolves the vault path", async () => {
    const done = exportDialog(app() as never, new TFile("Reports/Q3 Plan.docx") as never, formats([]));
    expect(modal().textContent).toContain("pdf options");
    expect((field("File name") as HTMLInputElement).value).toBe("Q3 Plan");
    expect(modal().textContent).toContain("Reports/Q3 Plan.pdf");
    button("Export").click();
    expect(await done).toEqual({ format: "pdf", target: { vault: "Reports/Q3 Plan.pdf" }, open: true });
  });

  it("rebuilds the options and extension when the format changes", async () => {
    const log: string[] = [];
    const done = exportDialog(app() as never, new TFile("Q3.docx") as never, formats(log));
    set(field("Format"), "md");
    expect(log).toEqual(["md options"]);
    expect(modal().textContent).not.toContain("pdf options");
    expect(modal().textContent).toContain("Q3.md");
    set(field("File name"), "Notes");
    button("Export").click();
    expect(await done).toMatchObject({ format: "md", target: { vault: "Notes.md" } });
  });

  it("warns about an existing file and offers Replace", async () => {
    const done = exportDialog(app(["Q3.pdf"]) as never, new TFile("Q3.docx") as never, formats([]));
    expect(modal().textContent).toContain("already exists");
    expect(button("Replace")).toBeTruthy();
    set(field("File name"), "Q3 final");
    expect(modal().textContent).not.toContain("already exists");
    button("Export").click();
    await done;
  });

  it("starts on a given format and remembers Open after export", async () => {
    const a = app();
    let done = exportDialog(a as never, new TFile("Q3.docx") as never, formats([]), "md");
    expect((field("Format") as HTMLSelectElement).value).toBe("md");
    const open = modal().querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    expect(open.checked).toBe(true);
    open.checked = false;
    open.dispatchEvent(new Event("change"));
    button("Export").click();
    expect((await done)!.open).toBe(false);
    done = exportDialog(a as never, new TFile("Q3.docx") as never, formats([]));
    expect(modal().querySelector<HTMLInputElement>('input[type="checkbox"]')!.checked).toBe(false);
    button("Cancel").click();
    expect(await done).toBeNull();
  });

  it("hides the format choice when there is only one", async () => {
    const done = exportDialog(app() as never, new TFile("n.md") as never, [{ id: "docx", label: "Word (.docx)", ext: "docx" }]);
    expect(field("Format")).toBeNull();
    button("Cancel").click();
    expect(await done).toBeNull();
  });
});

describe("writeExport", () => {
  it("creates a new vault file, or replaces an existing one", async () => {
    const created: string[] = [], modified: string[] = [];
    const a = app(["out/old.pdf"]) as any;
    a.vault.createBinary = async (p: string) => (created.push(p), new TFile(p));
    a.vault.modifyBinary = async (f: TFile) => void modified.push(f.path);
    a.vault.create = async (p: string) => (created.push(p), new TFile(p));
    a.vault.modify = async (f: TFile) => void modified.push(f.path);
    await writeExport(a, { vault: "out/new.pdf" }, new Uint8Array([1]), false);
    await writeExport(a, { vault: "out/old.pdf" }, new Uint8Array([1]), false);
    await writeExport(a, { vault: "out/sheet.csv" }, "a,b\r\n", false);
    expect([created, modified]).toEqual([["out/new.pdf", "out/sheet.csv"], ["out/old.pdf"]]);
  });
});

describe("format file names", () => {
  it("uses a format's own suggested name", async () => {
    const done = exportDialog(app() as never, new TFile("Deck.pptx") as never, [
      { id: "pdf", label: "PDF", ext: "pdf" },
      { id: "png", label: "PNG", ext: "png", name: "Deck - slide 3" },
    ], "png");
    expect((field("File name") as HTMLInputElement).value).toBe("Deck - slide 3");
    set(field("Format"), "pdf");
    expect((field("File name") as HTMLInputElement).value).toBe("Deck");
    button("Cancel").click();
    await done;
  });
});
