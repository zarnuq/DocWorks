import { afterEach, describe, expect, it } from "vitest";
import { App } from "./obsidian-stub";
import { specialCharacters } from "../src/presentation/common/special-chars";

afterEach(() => document.body.empty());

describe("special characters dialog", () => {
  it("lists characters by category, filters by name, and hands back the picked one (staying open)", () => {
    const picked: string[] = [];
    specialCharacters(new App() as any, (ch) => picked.push(ch));
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    const chars = () => Array.from(modal.querySelectorAll<HTMLButtonElement>(".dw-char")).map((b) => b.textContent);
    expect(chars().length).toBeGreaterThan(50);
    const search = modal.querySelector<HTMLInputElement>('input[type="search"]')!;
    search.value = "arrow right";
    search.dispatchEvent(new Event("input"));
    expect(chars()).toContain("→");
    expect(chars()).not.toContain("α");
    search.value = "alpha";
    search.dispatchEvent(new Event("input"));
    expect(chars()).toEqual(["α", "Α"]);
    modal.querySelector<HTMLButtonElement>(".dw-char")!.click();
    expect(picked).toEqual(["α"]);
    expect(document.querySelector(".test-modal")).not.toBeNull();
  });

  it("picks a category", () => {
    specialCharacters(new App() as any, () => {});
    const modal = document.querySelector<HTMLElement>(".test-modal")!;
    const category = modal.querySelector<HTMLSelectElement>('select[aria-label="Category"]')!;
    category.value = "Currency";
    category.dispatchEvent(new Event("change"));
    const chars = Array.from(modal.querySelectorAll(".dw-char")).map((b) => b.textContent);
    expect(chars).toContain("€");
    expect(chars).not.toContain("→");
  });
});
