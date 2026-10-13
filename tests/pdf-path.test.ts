import { describe, expect, it } from "vitest";
import { Platform } from "obsidian";
import { canExportPdf, renderPdf } from "../src/presentation/common/pdf";

describe("PDF export on mobile", () => {
  it("is offered only in the desktop app, and refuses without Electron", async () => {
    expect(canExportPdf()).toBe(true);
    Platform.isDesktopApp = false;
    try {
      expect(canExportPdf()).toBe(false);
      expect(await renderPdf([], { width: 1, height: 1 })).toBeNull();
    } finally {
      Platform.isDesktopApp = true;
    }
  });
});
