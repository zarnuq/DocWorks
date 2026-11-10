import { describe, expect, it } from "vitest";
import type { WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, WorkspaceLeaf } from "./obsidian-stub";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { DEFAULT_SETTINGS } from "../src/presentation/common/settings";

/** A key registered twice on one scope runs twice (or never, the second time): each view registers a key once. */
describe("view scopes", () => {
  for (const View of [DocxView, PptxView, XlsxView]) {
    it(`${View.name} registers each key once`, () => {
      const view = new View(new WorkspaceLeaf(new App()) as unknown as RealLeaf, { ...DEFAULT_SETTINGS });
      const keys = (view as unknown as { scope: { keys: { mods: string[]; key: string }[] } }).scope.keys.map((k) => `${[...k.mods].sort().join("+")}+${k.key.toLowerCase()}`);
      expect(keys.filter((k, i) => keys.indexOf(k) !== i)).toEqual([]);
    });
  }
});
