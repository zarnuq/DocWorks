/*
 * The save indicator at the end of every toolbar: "Saved", "Saving…", or "Couldn't save" with a Retry button. A
 * failed save is retried automatically after 5 s, 15 s and then every minute, as Google Docs does, until one works.
 */
import { Notice } from "obsidian";

const RETRY_DELAYS = [5000, 15000, 60000];

export class SaveStatus {
  readonly el: HTMLElement;
  private text: HTMLElement;
  private retryBtn: HTMLButtonElement | null = null;
  private timer: number | null = null;
  /** Failed saves in a row (0 once a save works). */
  private failures = 0;

  constructor(parent: HTMLElement) {
    this.el = parent.createSpan({ cls: "dw-status", attr: { role: "status" } });
    this.text = this.el.createSpan({ cls: "dw-status-text", text: "Saved" });
  }

  saving(): void {
    this.cancelRetry();
    this.show("Saving…");
  }

  saved(): void {
    this.cancelRetry();
    this.failures = 0;
    this.show("Saved");
  }

  /** A save failed: show it, offer Retry and schedule the next automatic attempt. */
  failed(err: unknown, retry: () => void, name: string): void {
    this.cancelRetry();
    const message = err instanceof Error ? err.message : String(err);
    if (this.failures++ === 0) new Notice(`Couldn't save ${name}: ${message}. Retrying automatically.`);
    this.show("Couldn't save", message);
    this.retryBtn = this.el.createEl("button", { cls: "dw-status-retry", text: "Retry" });
    this.retryBtn.addEventListener("click", () => {
      this.cancelRetry();
      retry();
    });
    const delay = RETRY_DELAYS[Math.min(this.failures - 1, RETRY_DELAYS.length - 1)];
    this.timer = window.setTimeout(() => {
      this.timer = null;
      retry();
    }, delay);
  }

  destroy(): void {
    this.cancelRetry();
  }

  private show(text: string, error?: string): void {
    this.text.setText(text);
    this.el.toggleClass("is-error", !!error);
    if (error) this.el.setAttribute("title", error);
    else this.el.removeAttribute("title");
    this.retryBtn?.remove();
    this.retryBtn = null;
  }

  private cancelRetry(): void {
    if (this.timer !== null) window.clearTimeout(this.timer);
    this.timer = null;
  }
}
