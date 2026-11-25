/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { lineNumbersOf } from "../src/domain/docx/sections";
import { fixture } from "./helpers";

afterEach(() => document.body.empty());

async function openDoc(readOnly = false) {
  const app = new App();
  const bytes = fixture("sample.docx");
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void saved.push(data);
  const view = new DocxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officeDocx: false, readOnly });
  document.body.appendChild(view.containerEl);
  const file = new TFile("lines.docx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const tools = () => {
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Tools")!.click();
    return Menu.last!;
  };
  const modal = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1)!;
  const dialog = (show: boolean, restart?: string) => {
    tools().items.find((i) => i.title === "Line numbers")!.run!();
    const toggle = modal().querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    toggle.checked = show;
    toggle.dispatchEvent(new Event("change"));
    if (restart) {
      const select = modal().querySelector("select")!;
      select.value = restart;
      select.dispatchEvent(new Event("change"));
    }
    Array.from(modal().querySelectorAll("button")).find((b) => b.textContent === "Apply")!.click();
  };
  const savedSectPr = async () => {
    await v.autosave.flush();
    const body = (await DocxDocument.load(new Uint8Array(saved.at(-1)!))).body;
    const last = Array.from(body.children).filter((c) => c.localName === "sectPr").pop()!;
    return new XMLSerializer().serializeToString(last);
  };
  return { view, v, tools, modal, dialog, savedSectPr };
}

describe("Tools › Line numbers", () => {
  it("offers Google's choices, off for a document without numbering", async () => {
    const { tools, modal } = await openDoc();
    tools().items.find((i) => i.title === "Line numbers")!.run!();
    expect(modal().querySelector<HTMLInputElement>('input[type="checkbox"]')!.checked).toBe(false);
    expect(Array.from(modal().querySelectorAll("option")).map((o) => o.textContent)).toEqual(["Restart each page", "Restart each section", "Continuous"]);
  });

  it("showing them writes w:lnNumType into the section, and the dialog opens on what's set", async () => {
    const { dialog, savedSectPr, tools, modal } = await openDoc();
    dialog(true, "continuous");
    expect(lineNumbersOf(await savedSectPr())).toEqual({ start: 1, countBy: 1, restart: "continuous" });
    tools().items.find((i) => i.title === "Line numbers")!.run!();
    expect(modal().querySelector<HTMLInputElement>('input[type="checkbox"]')!.checked).toBe(true);
    expect(modal().querySelector("select")!.value).toBe("continuous");
  });

  it("turning them off removes it; undo brings it back", async () => {
    const { dialog, savedSectPr, v } = await openDoc();
    const { closeHistory, undo } = await import("prosemirror-history");
    dialog(true);
    v.editor.dispatch(closeHistory(v.editor.state.tr));
    dialog(false);
    expect(await savedSectPr()).not.toContain("lnNumType");
    undo(v.editor.state, v.editor.dispatch);
    expect(lineNumbersOf(await savedSectPr())).toEqual({ start: 1, countBy: 1, restart: "newPage" });
  });

  it("isn't offered for a read-only file", async () => {
    const { tools } = await openDoc(true);
    expect(tools().items.map((i) => i.title)).not.toContain("Line numbers");
  });
});
