/**
 * Autosave, shared by the three editors: a change schedules one save a moment later, saves run one at a time
 * (a package must not change while the previous save is zipping it), and a failed save keeps the edits and hands
 * Retry to the toolbar's SaveStatus. Read-only files are never saved.
 */
import { type FileView, Notice } from "obsidian";
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
  /** When the last write ended (or now, while one runs). */
  private lastWrite = 0;

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
        this.lastWrite = Date.now();
        try {
          await this.opts.write();
        } finally {
          this.lastWrite = Date.now();
        }
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

  /** Whether a vault `modify` now is likely this editor's own save. */
  wroteRecently(): boolean {
    return Date.now() - this.lastWrite < 2000;
  }

  private cancelTimer(): void {
    if (this.timer !== null) window.clearTimeout(this.timer);
    this.timer = null;
  }
}

/** Reload the view's file when something else changes it (sync, Office), unless that would throw away unsaved edits. */
export function reloadOnOutsideChange(view: FileView, autosave: Autosave, reload: () => Promise<void>): void {
  view.registerEvent(
    view.app.vault.on("modify", (f) => {
      if (f !== view.file || autosave.wroteRecently()) return;
      if (autosave.dirty) new Notice(`${f.name} changed on disk. Your unsaved edits here will replace that version.`);
      else void reload();
    })
  );
}
