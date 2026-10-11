import { describe, expect, it } from "vitest";
import { App } from "obsidian";
import { promptText } from "../src/presentation/common/toolbar";

describe("text prompt", () => {
  it("submits on Enter without letting the key reach whatever gets focus next", async () => {
    const answer = promptText(new App(), "Your name", "Ada");
    const input = document.querySelector<HTMLInputElement>(".test-modal input")!;
    const enter = new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true });
    input.dispatchEvent(enter);
    expect(await answer).toBe("Ada");
    expect(enter.defaultPrevented).toBe(true);
  });
});
