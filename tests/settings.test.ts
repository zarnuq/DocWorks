import { afterEach, describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, paperFor } from "../src/presentation/common/settings";

describe("default paper theme setting", () => {
  afterEach(() => document.body.classList.remove("theme-dark"));

  it("keeps each file type's own default unless set", () => {
    expect(DEFAULT_SETTINGS.paper).toBe("default");
    expect(paperFor(DEFAULT_SETTINGS, "light")).toBe("light");
    expect(paperFor(DEFAULT_SETTINGS, "dark")).toBe("dark");
  });

  it("can force light or dark, or follow Obsidian's theme", () => {
    expect(paperFor({ ...DEFAULT_SETTINGS, paper: "dark" }, "light")).toBe("dark");
    expect(paperFor({ ...DEFAULT_SETTINGS, paper: "light" }, "dark")).toBe("light");
    expect(paperFor({ ...DEFAULT_SETTINGS, paper: "obsidian" }, "dark")).toBe("light");
    document.body.classList.add("theme-dark");
    expect(paperFor({ ...DEFAULT_SETTINGS, paper: "obsidian" }, "light")).toBe("dark");
  });

  it("opens every file type by default", () => {
    expect([DEFAULT_SETTINGS.openDocx, DEFAULT_SETTINGS.openPptx, DEFAULT_SETTINGS.openXlsx]).toEqual([true, true, true]);
  });
});
