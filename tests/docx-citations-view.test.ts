/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { TextSelection } from "prosemirror-state";
import { App, Menu, Notice, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readDocx } from "../src/domain/docx/docx-reader";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { fixture } from "./helpers";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

async function openDoc(readOnly = false) {
  const app = new App();
  const bytes = fixture("sample.docx");
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void saved.push(data);
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officeDocx: false, readOnly });
  document.body.appendChild(view.containerEl);
  const file = new TFile("cite.docx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const ed = () => v.editor;
  const caretAtEnd = () => ed().dispatch(ed().state.tr.setSelection(TextSelection.create(ed().state.doc, ed().state.doc.child(0).nodeSize - 1)));
  const menu = (name: string, item: string) => {
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!.click();
    const found = Menu.last!.items.find((i) => i.title === item);
    if (!found) throw new Error(`No ${item}: ${Menu.last!.items.map((i) => i.title).join(", ")}`);
    found.run!(new MouseEvent("click"));
  };
  const modals = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal"));
  const top = () => modals().at(-1)!;
  const button = (text: string, root = top()) => {
    const b = Array.from(root.querySelectorAll("button")).find((x) => x.textContent === text);
    if (!b) throw new Error(`No button ${text}: ${Array.from(root.querySelectorAll("button")).map((x) => x.textContent).join(", ")}`);
    return b;
  };
  const flush = () => new Promise((r) => setTimeout(r, 0));
  /** Fill the source form (the top modal) by setting names. */
  const fill = (values: Record<string, string>) => {
    for (const item of Array.from(top().querySelectorAll<HTMLElement>(".setting-item"))) {
      const name = item.querySelector(".setting-item-name")?.textContent ?? "";
      if (!(name in values)) continue;
      const input = item.querySelector<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>("input, select, textarea")!;
      input.value = values[name];
      input.dispatchEvent(new Event("change"));
    }
  };
  const addSource = async (values: Record<string, string>) => {
    button("Add citation source").click();
    fill(values);
    button("Add source").click();
    await flush();
  };
  const citations = () => {
    const out: any[] = [];
    ed().state.doc.descendants((n: any) => void (n.type === schema.nodes.chip && n.attrs.kind === "citation" && out.push(n)));
    return out;
  };
  const save = async () => {
    await v.autosave.flush();
    return DocxDocument.load(new Uint8Array(saved.at(-1)!));
  };
  return { view, v, ed, caretAtEnd, menu, modals, top, button, flush, fill, addSource, citations, save };
}

describe("Tools › Citations", () => {
  it("is in the Tools menu, not for read-only files", async () => {
    const { view } = await openDoc(true);
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Tools")!.click();
    expect(Menu.last!.items.map((i) => i.title)).not.toContain("Citations");
  });

  it("starts with no sources, APA, and Insert bibliography disabled", async () => {
    const { menu, top, button } = await openDoc();
    menu("Tools", "Citations");
    expect(top().textContent).toContain("No sources yet");
    expect(top().querySelector<HTMLSelectElement>("select")!.value).toBe("apa");
    expect(button("Insert bibliography").disabled).toBe(true);
  });

  it("adds a source, cites it at the caret, and saves both in Word's format", async () => {
    const { menu, caretAtEnd, addSource, top, button, citations, save } = await openDoc();
    caretAtEnd();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Ann Smith", Title: "The Long Road", Year: "2020", Publisher: "Penguin" });
    expect(top().textContent).toContain("The Long Road");
    expect(top().textContent).toContain("Book · Ann Smith · 2020");
    button("Cite").click();
    expect(citations().map((c) => c.attrs.text)).toEqual(["(Smith, 2020)"]);
    const reopened = await save();
    expect(reopened.sources().sources[0]).toMatchObject({ tag: "Smi20", title: "The Long Road", publisher: "Penguin", authors: [{ first: "Ann", last: "Smith" }] });
    const { doc } = readDocx(reopened.body, { imageSrc: () => null, hyperlinkHref: () => null });
    const chips: any[] = [];
    doc.descendants((n) => void (n.type === schema.nodes.chip && chips.push(n.attrs)));
    expect(chips).toEqual([expect.objectContaining({ kind: "citation", tag: "Smi20", text: "(Smith, 2020)" })]);
  });

  it("a source without a title isn't added", async () => {
    const { menu, button, top, modals } = await openDoc();
    menu("Tools", "Citations");
    button("Add citation source").click();
    const before = Notice.messages.length;
    button("Add source").click();
    expect(Notice.messages.slice(before)).toEqual(["A source needs a title."]);
    expect(modals()).toHaveLength(2);
    button("Cancel").click();
    expect(top().textContent).toContain("No sources yet");
  });

  it("the form shows only the fields of the source type", async () => {
    const { menu, button, top, fill } = await openDoc();
    menu("Tools", "Citations");
    button("Add citation source").click();
    const shown = () => Array.from(top().querySelectorAll<HTMLElement>(".setting-item")).filter((el) => el.style.display !== "none").map((el) => el.querySelector(".setting-item-name")?.textContent).filter(Boolean);
    expect(shown()).toEqual(["Source type", "Contributors", "Title", "Year", "Publisher", "URL"]);
    fill({ "Source type": "journal" });
    expect(shown()).toEqual(["Source type", "Contributors", "Title", "Year", "Journal", "Volume", "Issue", "Pages", "URL"]);
    fill({ "Source type": "website" });
    expect(shown()).toEqual(["Source type", "Contributors", "Title", "Year", "Website name", "URL"]);
  });

  it("changing the format restyles every citation and the bibliography, and is saved", async () => {
    const { menu, caretAtEnd, addSource, button, top, citations, ed, save, flush } = await openDoc();
    caretAtEnd();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Ann Smith", Title: "The Long Road", Year: "2020", Publisher: "Penguin" });
    button("Cite").click();
    menu("Tools", "Citations");
    button("Insert bibliography").click();
    menu("Tools", "Citations");
    const format = top().querySelector<HTMLSelectElement>("select")!;
    format.value = "mla";
    format.dispatchEvent(new Event("change"));
    await flush();
    expect(citations().map((c) => c.attrs.text)).toEqual(["(Smith)"]);
    expect(ed().state.doc.textContent).toContain("Works Cited");
    expect(ed().state.doc.textContent).toContain("Smith, Ann. The Long Road. Penguin, 2020.");
    expect((await save()).sources().style).toBe("mla");
  });

  it("editing a source updates its citations; deleting it leaves them as text", async () => {
    const { menu, caretAtEnd, addSource, button, top, citations, fill, flush } = await openDoc();
    caretAtEnd();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Ann Smith", Title: "The Long Road", Year: "2020" });
    button("Cite").click();
    menu("Tools", "Citations");
    button("Edit").click();
    fill({ Year: "2021" });
    button("Save").click();
    await flush();
    expect(citations()[0].attrs.text).toBe("(Smith, 2021)");
    expect(top().textContent).toContain("2021");
    button("Delete").click();
    expect(top().textContent).toContain("No sources yet");
    expect(citations()[0].attrs.text).toBe("(Smith, 2021)");
  });

  it("Insert bibliography adds the list after the caret's paragraph", async () => {
    const { menu, caretAtEnd, addSource, button, ed, modals } = await openDoc();
    caretAtEnd();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Zed Young", Title: "Last", Year: "2001" });
    await addSource({ Contributors: "Al Adams", Title: "First", Year: "1999" });
    button("Insert bibliography").click();
    expect(modals()).toHaveLength(0);
    const bib = ed().state.doc.child(1);
    expect(bib.attrs.kind).toBe("bibliography");
    expect(bib.child(0).textContent).toBe("References");
    expect([bib.child(1).textContent, bib.child(2).textContent]).toEqual(["Adams, A. (1999). First.", "Young, Z. (2001). Last."]);
  });

  it("two sources of one author and year get different tags", async () => {
    const { menu, addSource, save } = await openDoc();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Ann Smith", Title: "One", Year: "2020" });
    await addSource({ Contributors: "Ann Smith", Title: "Two", Year: "2020" });
    expect((await save()).sources().sources.map((s) => s.tag)).toEqual(["Smi20", "Smi201"]);
  });

  it("clicking a citation doesn't open a chip menu", async () => {
    const { menu, caretAtEnd, addSource, button, v } = await openDoc();
    caretAtEnd();
    menu("Tools", "Citations");
    await addSource({ Contributors: "Ann Smith", Title: "T", Year: "2020" });
    button("Cite").click();
    const chip = v.editor.state.doc.child(0).lastChild;
    const handled = v.editor.someProp("handleClickOn", (f: any) => f(v.editor, 0, chip, 0, new MouseEvent("click"), true));
    expect(handled).toBeFalsy();
  });
});
