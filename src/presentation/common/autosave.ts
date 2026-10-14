/**
 * Autosave, shared by the three editors: a change schedules one save a moment later, saves run one at a time
 * (a package must not change while the previous save is zipping it), and a failed save keeps the edits and hands
 * Retry to the toolbar's SaveStatus. Read-only files are never saved.
 */
import type { SaveStatus } from "./save-status";

export interface AutosaveOptions {
  /** Milliseconds after the last change. */
  delay: number;
  readOnly: () => boolean;
  /** Whether a file is open to save into. */
  canSave: () => boolean;
  /** The file's name, for the error notice. */
  name: () => string;
  /** Write the file. */
  write: () => Promise<void>;
  /** Run first in each save (an editor flushing pending text, which may mark the file dirty). */
  before?: () => void;
  /** Whether edits are still waiting after a save (shown as "Saving…"). */
  pending?: () => boolean;
}

export class Autosave {
  /** Edits not written yet. */
  dirty = false;
  status: SaveStatus | null = null;
  private timer: number | null = null;
  private queue: Promise<void> = Promise.resolve();

  constructor(private opts: AutosaveOptions) {}

  markDirty(): void {
    if (this.opts.readOnly()) return;
    this.dirty = true;
    this.status?.saving();
    this.cancelTimer();
    this.timer = window.setTimeout(() => void this.save(), this.opts.delay);
  }

  /** Save now when there are edits (Obsidian quitting, the window losing focus). */
  flush(): Promise<void> {
    this.cancelTimer();
    return this.save();
  }

  /** The file is closing: a last save, then no more retries. */
  async close(): Promise<void> {
    await this.flush();
    this.status?.destroy();
  }

  /** Resolves once the saves queued so far are done. */
  idle(): Promise<void> {
    return this.queue;
  }

  private save(): Promise<void> {
    this.queue = this.queue.then(async () => {
      this.opts.before?.();
      if (!this.dirty || this.opts.readOnly() || !this.opts.canSave()) return;
      this.cancelTimer();
      const name = this.opts.name();
      this.dirty = false;
      try {
        await this.opts.write();
        if (this.dirty || this.opts.pending?.()) this.status?.saving();
        else this.status?.saved();
      } catch (err) {
        this.dirty = true;
        console.error(`Failed to save ${name}:`, err);
        this.status?.failed(err, () => void this.save(), name);
      }
    });
    return this.queue;
  }

  private cancelTimer(): void {
    if (this.timer !== null) window.clearTimeout(this.timer);
    this.timer = null;
  }
}
