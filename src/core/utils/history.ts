/** Undo/redo stacks of state snapshots. Callers pass the current state when moving. */
export class History<T> {
  private past: T[] = [];
  private future: T[] = [];

  constructor(private limit = 100) {}

  /** Record the state *before* a change. */
  push(state: T): void {
    this.past.push(state);
    if (this.past.length > this.limit) this.past.shift();
    this.future = [];
  }

  undo(current: T): T | null {
    const prev = this.past.pop();
    if (prev === undefined) return null;
    this.future.push(current);
    return prev;
  }

  redo(current: T): T | null {
    const next = this.future.pop();
    if (next === undefined) return null;
    this.past.push(current);
    return next;
  }
}
