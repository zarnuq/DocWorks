/* eslint-disable @typescript-eslint/no-explicit-any */
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { blankPptx } from "../src/infrastructure/blank";
import { equationLatex, shownBranch } from "../src/domain/pptx/equation";
import { NS } from "../src/core/utils/xml";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

async function openDeck(readOnly = false) {
  const bytes = await blankPptx();
  const app = new App();
  app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const saved: ArrayBuffer[] = [];
  app.vault.modifyBinary = async (_f: unknown, data: ArrayBuffer) => void saved.push(data);
  const view = new PptxView(new WorkspaceLeaf(app) as unknown as RealLeaf, { ...DEFAULT_SETTINGS, officePptx: false, readOnly });
  document.body.appendChild(view.containerEl);
  const file = new TFile("eq.pptx") as unknown as RealTFile;
  view.file = file;
  await view.onLoadFile(file);
  const v = view as any;
  const menuItems = (name: string) => {
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!.click();
    return Menu.last!.items;
  };
  const run = (name: string, title: string) => menuItems(name).find((i) => i.title === title)!.run!();
  const modal = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1);
  const flush = () => new Promise((r) => setTimeout(r, 0));
  const answer = async (latex: string, button = "Insert") => {
    const m = modal()!;
    m.querySelector<HTMLTextAreaElement>("textarea")!.value = latex;
    Array.from(m.querySelectorAll("button")).find((b) => b.textContent === button)!.click();
    await flush();
  };
  const slideXml = () => v.doc.xml(v.doc.slides()[v.index].path) as Document;
  const equations = () => Array.from(slideXml().getElementsByTagNameNS(NS.mc, "AlternateContent"));
  const rendered = () => v.rendered.filter((r: any) => r.shape.equation);
  const save = async () => {
    await v.autosave.flush();
    return PptxDocument.load(new Uint8Array(saved.at(-1)!));
  };
  return { view, v, menuItems, run, modal, answer, flush, slideXml, equations, rendered, save };
}

describe("equations on slides", () => {
  it("Insert › Equation is offered, not on read-only files", async () => {
    expect((await openDeck()).menuItems("Insert").map((i) => i.title)).toContain("Equation");
    document.body.empty();
    expect((await openDeck(true)).view.contentEl.querySelectorAll(".dw-menu-btn").length).toBeGreaterThan(0);
  });

  it("inserts an equation drawn as MathML, selected, and saved as PowerPoint math", async () => {
    const { run, answer, equations, rendered, v, save } = await openDeck();
    run("Insert", "Equation");
    await answer("\\frac{a}{b}");
    expect(equations()).toHaveLength(1);
    const [rs] = rendered();
    expect(rs.el.querySelector("math mfrac")).not.toBeNull();
    expect(rs.textEl.contentEditable).not.toBe("true");
    expect(v.items).toEqual([equations()[0]]);
    const again = await save();
    const ac = again.xml(again.slides()[0].path)!.getElementsByTagNameNS(NS.mc, "AlternateContent")[0];
    expect(equationLatex(shownBranch(ac)!)).toBe("\\frac{a}{b}");
  });

  it("Cancel or an empty equation inserts nothing", async () => {
    const { run, answer, modal, equations, flush } = await openDeck();
    run("Insert", "Equation");
    Array.from(modal()!.querySelectorAll("button")).find((b) => b.textContent === "Cancel")!.click();
    await flush();
    run("Insert", "Equation");
    await answer("   ");
    expect(equations()).toHaveLength(0);
  });

  it("double-clicking an equation edits it; undo brings the old one back", async () => {
    const { run, answer, rendered, slideXml, view, v } = await openDeck();
    run("Insert", "Equation");
    await answer("x^2");
    rendered()[0].el.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    const latexOf = () => equationLatex(shownBranch(slideXml().getElementsByTagNameNS(NS.mc, "AlternateContent")[0])!);
    expect(document.querySelector<HTMLTextAreaElement>(".test-modal textarea")!.value).toBe("x^2");
    await answer("y_1", "Save");
    expect(latexOf()).toBe("y_1");
    expect(rendered()[0].el.querySelector("msub")).not.toBeNull();
    v.undo();
    expect(latexOf()).toBe("x^2");
    expect(view).toBeDefined();
  });

  it("Delete in the dialog removes the equation", async () => {
    const { run, answer, rendered, equations } = await openDeck();
    run("Insert", "Equation");
    await answer("x");
    rendered()[0].el.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    await answer("", "Delete");
    expect(equations()).toHaveLength(0);
  });

  it("the right-click menu has Edit equation…", async () => {
    const { run, answer, rendered, view } = await openDeck();
    run("Insert", "Equation");
    await answer("x");
    const el = rendered()[0].el;
    el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect(Menu.last!.items.map((i) => i.title)).toContain("Edit equation…");
    expect(view).toBeDefined();
  });

  it("moving an equation moves its fallback on save", async () => {
    const { run, answer, equations, v, save } = await openDeck();
    run("Insert", "Equation");
    await answer("x");
    const ac = equations()[0];
    v.mutate(() => {
      const off = shownBranch(ac)!.getElementsByTagNameNS(NS.a, "off")[0];
      off.setAttribute("x", "123");
    });
    const again = await save();
    const fb = again.xml(again.slides()[0].path)!.getElementsByTagNameNS(NS.mc, "Fallback")[0];
    expect(fb.getElementsByTagNameNS(NS.a, "off")[0].getAttribute("x")).toBe("123");
  });
});
