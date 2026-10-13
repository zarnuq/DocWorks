import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Notice } from "./obsidian-stub";
import { SaveStatus } from "../src/presentation/common/save-status";

describe("SaveStatus", () => {
  let status: SaveStatus;
  const text = () => status.el.querySelector(".dw-status-text")!.textContent;
  const retryButton = () => status.el.querySelector<HTMLButtonElement>("button");

  beforeEach(() => {
    vi.useFakeTimers();
    Notice.messages.length = 0;
    status = new SaveStatus(document.body.createDiv());
  });
  afterEach(() => {
    status.destroy();
    vi.useRealTimers();
    document.body.empty();
  });

  it("shows Saved, then Saving… while saving, as a polite live region", () => {
    expect(text()).toBe("Saved");
    expect(status.el.getAttribute("role")).toBe("status");
    status.saving();
    expect(text()).toBe("Saving…");
    status.saved();
    expect(text()).toBe("Saved");
    expect(retryButton()).toBeNull();
  });

  it("shows the error with a Retry button, and a Notice for the first failure only", () => {
    const retry = vi.fn();
    status.failed(new Error("disk full"), retry, "a.docx");
    expect(text()).toBe("Couldn't save");
    expect(status.el.getAttribute("title")).toBe("disk full");
    expect(status.el.classList.contains("is-error")).toBe(true);
    retryButton()!.click();
    expect(retry).toHaveBeenCalledTimes(1);
    status.failed(new Error("disk full"), retry, "a.docx");
    expect(Notice.messages).toEqual(["Couldn't save a.docx: disk full. Retrying automatically."]);
  });

  it("retries after 5 s, 15 s, then every 60 s, and stops once saved", () => {
    const retry = vi.fn(() => status.failed(new Error("x"), retry, "a"));
    status.failed(new Error("x"), retry, "a");
    vi.advanceTimersByTime(4999);
    expect(retry).toHaveBeenCalledTimes(0);
    vi.advanceTimersByTime(1);
    expect(retry).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(15000);
    expect(retry).toHaveBeenCalledTimes(2);
    vi.advanceTimersByTime(60000);
    expect(retry).toHaveBeenCalledTimes(3);
    vi.advanceTimersByTime(60000);
    expect(retry).toHaveBeenCalledTimes(4);
    retry.mockImplementation(() => status.saved());
    vi.advanceTimersByTime(60000);
    expect(text()).toBe("Saved");
    vi.advanceTimersByTime(600000);
    expect(retry).toHaveBeenCalledTimes(5);
  });

  it("cancels a pending retry when a save starts some other way, and stops on destroy", () => {
    const retry = vi.fn();
    status.failed(new Error("x"), retry, "a");
    status.saving();
    vi.advanceTimersByTime(10000);
    expect(retry).not.toHaveBeenCalled();
    status.failed(new Error("x"), retry, "a");
    status.destroy();
    vi.advanceTimersByTime(100000);
    expect(retry).not.toHaveBeenCalled();
  });
});
