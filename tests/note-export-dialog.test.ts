import { afterEach, expect, it, vi } from "vitest";
import { App, Menu, TFile } from "obsidian";
import { addNoteExportToMenu, promptNoteExport } from "../src/presentation/common/export-note";

afterEach(() => document.body.replaceChildren());

const row = (name: string) => Array.from(document.querySelectorAll<HTMLElement>(".setting-item")).find(el => el.querySelector(".setting-item-name")?.textContent === name)!;
const click = (text: string) => Array.from(document.querySelectorAll<HTMLButtonElement>(".test-modal button")).find(b => b.textContent === text)!.click();

it("adds DOCX after the built-in items, once, when the menu is shown", () => {
  const menu = new Menu();
  const add = vi.spyOn(menu, "addItem");
  addNoteExportToMenu(menu, new App(), Object.assign(new TFile(), { path: "Note.md" }));
  expect(add).not.toHaveBeenCalled();
  menu.addItem(i => i.setTitle("Export to PDF..."));
  menu.showAtPosition({ x: 0, y: 0 });
  expect(add).toHaveBeenCalledTimes(2);
  const item = { setTitle: vi.fn().mockReturnThis(), setIcon: vi.fn().mockReturnThis(), setSection: vi.fn().mockReturnThis(), onClick: vi.fn().mockReturnThis() };
  add.mock.calls[1][0](item as never);
  expect(item.setTitle).toHaveBeenCalledWith("Export to .docx");
  expect(item.setSection).toHaveBeenCalledWith("action");
  menu.showAtMouseEvent(new MouseEvent("contextmenu"));
  expect(add).toHaveBeenCalledTimes(2);
});

it("returns no export settings when cancelled", async () => {
  const result = promptNoteExport(new App(), "Note");
  click("Cancel");
  expect(await result).toBeNull();
});

it("uses the selected settings only after pressing Export", async () => {
  const result = promptNoteExport(new App(), "My note");
  for (const name of ["Include file name as title", "Landscape"]) {
    const toggle = row(name).querySelector<HTMLInputElement>("input")!;
    toggle.checked = true;
    toggle.dispatchEvent(new Event("change"));
  }
  for (const [name, value] of [["Page size", "a4"], ["Margin", "narrow"], ["Downscale percent", "75"]]) {
    const input = row(name).querySelector<HTMLInputElement | HTMLSelectElement>("input,select")!;
    input.value = value;
    input.dispatchEvent(new Event("change"));
  }
  click("Export to .docx");
  expect(await result).toEqual({ pageSize: "a4", landscape: true, margin: "narrow", scale: 75, title: "My note" });
  expect(document.querySelector(".test-modal")).toBeNull();
});
