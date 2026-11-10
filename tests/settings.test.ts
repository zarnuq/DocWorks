import { afterEach, describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, migrateSettings, paperFor } from "../src/presentation/common/settings";

describe("paper theme per file type", () => {
  afterEach(() => document.body.classList.remove("theme-dark"));

  it("keeps each file type's own default unless set", () => {
    expect([DEFAULT_SETTINGS.paperDocx, DEFAULT_SETTINGS.paperPptx, DEFAULT_SETTINGS.paperXlsx]).toEqual(["default", "default", "default"]);
    expect(paperFor(DEFAULT_SETTINGS, "docx", "light")).toBe("light");
    expect(paperFor(DEFAULT_SETTINGS, "xlsx", "dark")).toBe("dark");
  });

  it("can force light or dark, or follow Obsidian's theme, for one file type only", () => {
    expect(paperFor({ ...DEFAULT_SETTINGS, paperDocx: "dark" }, "docx", "light")).toBe("dark");
    expect(paperFor({ ...DEFAULT_SETTINGS, paperDocx: "dark" }, "pptx", "light")).toBe("light");
    expect(paperFor({ ...DEFAULT_SETTINGS, paperXlsx: "light" }, "xlsx", "dark")).toBe("light");
    expect(paperFor({ ...DEFAULT_SETTINGS, paperPptx: "obsidian" }, "pptx", "dark")).toBe("light");
    document.body.classList.add("theme-dark");
    expect(paperFor({ ...DEFAULT_SETTINGS, paperDocx: "obsidian" }, "docx", "light")).toBe("dark");
  });

  it("opens every file type by default, with the Google layout", () => {
    expect([DEFAULT_SETTINGS.openDocx, DEFAULT_SETTINGS.openPptx, DEFAULT_SETTINGS.openXlsx]).toEqual([true, true, true]);
    expect(DEFAULT_SETTINGS.officeDocx).toBe(false);
    expect(DEFAULT_SETTINGS.officePptx).toBe(false);
    expect(DEFAULT_SETTINGS.officeXlsx).toBe(false);
  });
});

describe("settings saved by earlier versions", () => {
  it("copies the one paper setting to every file type and drops it", () => {
    const s = migrateSettings({ paper: "dark", author: "Ada" });
    expect([s.paperDocx, s.paperPptx, s.paperXlsx]).toEqual(["dark", "dark", "dark"]);
    expect("paper" in s).toBe(false);
    expect(s.author).toBe("Ada");
    expect(s.officeDocx).toBe(false);
    expect([s.officePptx, s.officeXlsx]).toEqual([false, false]);
    expect(migrateSettings({ officePptx: true, officeXlsx: "yes" })).toMatchObject({ officePptx: true, officeXlsx: false });
  });

  it("keeps the per-type defaults for 'default' and for no saved data", () => {
    for (const raw of [{ paper: "default" }, null]) {
      const s = migrateSettings(raw);
      expect([s.paperDocx, s.paperPptx, s.paperXlsx]).toEqual(["default", "default", "default"]);
    }
  });

  it("prefers the per-type values when both are saved", () => {
    const s = migrateSettings({ paperDocx: "light", paper: "dark" });
    expect([s.paperDocx, s.paperPptx, s.paperXlsx]).toEqual(["light", "dark", "dark"]);
  });

  it("ignores values it doesn't know", () => {
    const s = migrateSettings({ paper: "sepia", paperXlsx: 3 });
    expect([s.paperDocx, s.paperPptx, s.paperXlsx]).toEqual(["default", "default", "default"]);
  });
});
