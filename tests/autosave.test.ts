import { afterEach, describe, expect, it, vi } from "vitest";
import { Autosave } from "../src/presentation/common/autosave";
import { SaveStatus } from "../src/presentation/common/save-status";

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

function setup(opts: { readOnly?: boolean; fail?: () => boolean; canSave?: () => boolean } = {}) {
  const writes: number[] = [];
  const status = new SaveStatus(createDiv());
  const autosave = new Autosave({
    delay: 800,
    readOnly: () => opts.readOnly ?? false,
    canSave: opts.canSave ?? (() => true),
    name: () => "file.docx",
    write: async () => {
      writes.push(Date.now());
      if (opts.fail?.()) throw new Error("disk full");
    },
  });
  autosave.status = status;
  const text = () => status.el.querySelector(".dw-status-text")!.textContent;
  return { autosave, writes, status, text };
}

describe("Autosave", () => {
  it("saves once, a moment after the last change", async () => {
    vi.useFakeTimers();
    const { autosave, writes, text } = setup();
    autosave.markDirty();
    autosave.markDirty();
    expect(text()).toBe("Saving…");
    expect(autosave.dirty).toBe(true);
    await vi.advanceTimersByTimeAsync(799);
    expect(writes.length).toBe(0);
    await vi.advanceTimersByTimeAsync(1);
    await autosave.idle();
    expect(writes.length).toBe(1);
    expect(autosave.dirty).toBe(false);
    expect(text()).toBe("Saved");
  });

  it("never saves a read-only file", async () => {
    vi.useFakeTimers();
    const { autosave, writes } = setup({ readOnly: true });
    autosave.markDirty();
    await vi.advanceTimersByTimeAsync(5000);
    await autosave.flush();
    expect(writes.length).toBe(0);
    expect(autosave.dirty).toBe(false);
  });

  it("keeps the edits and offers Retry when a save fails", async () => {
    vi.useFakeTimers();
    let fail = true;
    const { autosave, writes, text, status } = setup({ fail: () => fail });
    autosave.markDirty();
    await autosave.flush();
    expect(autosave.dirty).toBe(true);
    expect(text()).toBe("Couldn't save");
    fail = false;
    status.el.querySelector<HTMLButtonElement>(".dw-status-retry")!.click();
    await autosave.idle();
    expect(writes.length).toBe(2);
    expect(text()).toBe("Saved");
  });

  it("flush saves now and close waits for the last save", async () => {
    const { autosave, writes } = setup();
    await autosave.flush();
    expect(writes.length).toBe(0);
    autosave.markDirty();
    await autosave.close();
    expect(writes.length).toBe(1);
  });

  it("leaves nothing to save while no file is open", async () => {
    let open = false;
    const { autosave, writes } = setup({ canSave: () => open });
    autosave.markDirty();
    await autosave.flush();
    expect(writes.length).toBe(0);
    expect(autosave.dirty).toBe(true);
    open = true;
    await autosave.flush();
    expect(writes.length).toBe(1);
  });

  it("runs one save at a time; a change made while saving is saved after", async () => {
    const order: string[] = [];
    let second: Promise<void> | null = null;
    const autosave: Autosave = new Autosave({
      delay: 800,
      readOnly: () => false,
      canSave: () => true,
      name: () => "f",
      write: async () => {
        const which = second ? "b" : "a";
        order.push(`${which} start`);
        if (!second) {
          autosave.markDirty();
          second = autosave.flush();
        }
        await new Promise((r) => setTimeout(r, which === "a" ? 20 : 0));
        order.push(`${which} end`);
      },
    });
    autosave.markDirty();
    await autosave.flush();
    await second;
    expect(order).toEqual(["a start", "a end", "b start", "b end"]);
    expect(autosave.dirty).toBe(false);
  });
});

describe("Autosave's own writes", () => {
  it("count as its own while a slow save is still writing (the vault reports the change before the write returns)", async () => {
    let now = 1_000_000;
    vi.spyOn(Date, "now").mockImplementation(() => now);
    let during: boolean | null = null;
    const autosave = new Autosave({
      delay: 0,
      readOnly: () => false,
      canSave: () => true,
      name: () => "big.docx",
      write: async () => {
        // A big file: building and writing it takes 5 s; the vault's modify event comes at the end of the write.
        now += 5000;
        during = autosave.wroteRecently();
      },
    });
    autosave.markDirty();
    await autosave.flush();
    expect(during).toBe(true);
    now += 1000;
    expect(autosave.wroteRecently()).toBe(true);
    now += 2000;
    expect(autosave.wroteRecently()).toBe(false);
  });
});
