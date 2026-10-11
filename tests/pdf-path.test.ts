import { describe, expect, it } from "vitest";
import { vaultPath } from "../src/presentation/common/pdf";

describe("PDF save location", () => {
  it("maps a path inside the vault to a vault path", () => {
    expect(vaultPath("/home/a/notes", "/home/a/notes/Folder/Report.pdf")).toBe("Folder/Report.pdf");
    expect(vaultPath("C:\\Vault", "C:\\Vault\\Sub\\Deck.pdf")).toBe("Sub/Deck.pdf");
  });

  it("returns null outside the vault, including folders that only share a prefix", () => {
    expect(vaultPath("/home/a/notes", "/home/a/Documents/Report.pdf")).toBeNull();
    expect(vaultPath("/home/a/notes", "/home/a/notes-old/Report.pdf")).toBeNull();
    expect(vaultPath(null, "/home/a/notes/Report.pdf")).toBeNull();
  });
});
