/*
 * Version history: before Doc Works saves over a file, the file as it was is kept, once per editing session and then
 * at most every 10 minutes. Versions live in the app's own storage (IndexedDB, as Obsidian's File recovery), never in
 * the vault, and are pruned: 10 per file (1 for files over 25 MB), none older than the days set, 200 MB in all.
 */
import type { App, TFile } from "obsidian";
export interface VersionMeta {
  id: number;
  path: string;
  time: number;
  size: number;
}

/** Where versions are kept: metadata apart from bytes, so lists don't load files. */
export interface VersionStore {
  all(): Promise<VersionMeta[]>;
  add(meta: Omit<VersionMeta, "id">, bytes: Uint8Array): Promise<void>;
  bytes(id: number): Promise<Uint8Array | null>;
  remove(ids: number[]): Promise<void>;
  setPath(ids: number[], path: string): Promise<void>;
}

const EVERY = 10 * 60_000;
const PER_FILE = 10;
const BIG = 25 * 1024 * 1024;
const TOTAL = 200 * 1024 * 1024;

export class VersionHistory {
  /** When each file was last kept in this session. */
  private kept = new Map<string, number>();

  constructor(
    private store: VersionStore,
    private settings: () => { enabled: boolean; days: number },
    private now: () => number = Date.now
  ) {}

  /** Keep `path` as it is on disk (`read`), unless it was kept less than 10 minutes ago (`force`: always). */
  async beforeSave(path: string, read: () => Promise<Uint8Array>, force = false): Promise<void> {
    if (!this.settings().enabled) return;
    const now = this.now();
    const last = this.kept.get(path);
    if (!force && last !== undefined && now - last < EVERY) return;
    this.kept.set(path, now);
    const bytes = await read();
    if (!bytes.length) return;
    await this.store.add({ path, time: now, size: bytes.length }, bytes);
    await this.prune();
  }

  /** A file's versions, newest first. */
  async list(path: string): Promise<VersionMeta[]> {
    return (await this.store.all()).filter((v) => v.path === path).sort((a, b) => b.time - a.time || b.id - a.id);
  }

  bytes(id: number): Promise<Uint8Array | null> {
    return this.store.bytes(id);
  }

  async renamed(from: string, to: string): Promise<void> {
    const ids = (await this.store.all()).filter((v) => v.path === from).map((v) => v.id);
    if (ids.length) await this.store.setPath(ids, to);
    const last = this.kept.get(from);
    this.kept.delete(from);
    if (last !== undefined) this.kept.set(to, last);
  }

  async deleted(path: string): Promise<void> {
    const ids = (await this.store.all()).filter((v) => v.path === path).map((v) => v.id);
    if (ids.length) await this.store.remove(ids);
    this.kept.delete(path);
  }

  /** Too old, too many for one file, then the oldest until everything fits in 200 MB. */
  private async prune(): Promise<void> {
    const cutoff = this.now() - this.settings().days * 86_400_000;
    const all = (await this.store.all()).sort((a, b) => b.time - a.time || b.id - a.id);
    const drop = new Set<number>();
    const counts = new Map<string, number>();
    for (const v of all) {
      const n = (counts.get(v.path) ?? 0) + 1;
      counts.set(v.path, n);
      if (v.time < cutoff || n > (v.size > BIG ? 1 : PER_FILE)) drop.add(v.id);
    }
    let total = 0;
    for (const v of all) {
      if (drop.has(v.id)) continue;
      total += v.size;
      if (total > TOTAL) drop.add(v.id);
    }
    if (drop.size) await this.store.remove([...drop]);
  }
}

/** Versions in memory (tests). */
export class MemoryVersionStore implements VersionStore {
  private next = 1;
  private items = new Map<number, { meta: VersionMeta; bytes: Uint8Array }>();

  async all(): Promise<VersionMeta[]> {
    return [...this.items.values()].map((i) => ({ ...i.meta }));
  }

  async add(meta: Omit<VersionMeta, "id">, bytes: Uint8Array): Promise<void> {
    const id = this.next++;
    this.items.set(id, { meta: { ...meta, id }, bytes });
  }

  async bytes(id: number): Promise<Uint8Array | null> {
    return this.items.get(id)?.bytes ?? null;
  }

  async remove(ids: number[]): Promise<void> {
    for (const id of ids) this.items.delete(id);
  }

  async setPath(ids: number[], path: string): Promise<void> {
    for (const id of ids) {
      const item = this.items.get(id);
      if (item) item.meta.path = path;
    }
  }
}

/** Versions in IndexedDB, one database per vault. */
export class IndexedDbVersionStore implements VersionStore {
  private db: Promise<IDBDatabase>;

  constructor(name: string) {
    this.db = new Promise((resolve, reject) => {
      const open = indexedDB.open(name, 1);
      open.onupgradeneeded = () => {
        open.result.createObjectStore("meta", { keyPath: "id", autoIncrement: true });
        open.result.createObjectStore("data");
      };
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error ?? new Error("Couldn't open version history"));
    });
  }

  private async run<T>(stores: string[], mode: IDBTransactionMode, work: (t: IDBTransaction) => IDBRequest<T> | void): Promise<T | undefined> {
    const t = (await this.db).transaction(stores, mode);
    const request = work(t);
    return new Promise((resolve, reject) => {
      t.oncomplete = () => resolve(request ? request.result : undefined);
      t.onerror = () => reject(t.error ?? new Error("Version history failed"));
      t.onabort = () => reject(t.error ?? new Error("Version history failed"));
    });
  }

  async all(): Promise<VersionMeta[]> {
    return ((await this.run(["meta"], "readonly", (t) => t.objectStore("meta").getAll())) as VersionMeta[] | undefined) ?? [];
  }

  async add(meta: Omit<VersionMeta, "id">, bytes: Uint8Array): Promise<void> {
    await this.run(["meta", "data"], "readwrite", (t) => {
      const add = t.objectStore("meta").add(meta);
      add.onsuccess = () => t.objectStore("data").put(bytes, add.result);
    });
  }

  async bytes(id: number): Promise<Uint8Array | null> {
    return ((await this.run(["data"], "readonly", (t) => t.objectStore("data").get(id))) as Uint8Array | undefined) ?? null;
  }

  async remove(ids: number[]): Promise<void> {
    await this.run(["meta", "data"], "readwrite", (t) => {
      for (const id of ids) {
        t.objectStore("meta").delete(id);
        t.objectStore("data").delete(id);
      }
    });
  }

  async setPath(ids: number[], path: string): Promise<void> {
    const all = await this.all();
    await this.run(["meta"], "readwrite", (t) => {
      for (const v of all.filter((v) => ids.includes(v.id))) t.objectStore("meta").put({ ...v, path });
    });
  }
}

/** Keep `file` as it is on disk before saving over it; failing to keep it never stops the save. */
export async function keepVersion(app: App, file: TFile, force = false): Promise<void> {
  if (!versions) return;
  try {
    await versions.beforeSave(file.path, async () => new Uint8Array(await app.vault.readBinary(file)), force);
  } catch (err) {
    console.warn("[doc-works] couldn't keep a version", err);
  }
}

/** The vault's history, set up by the plugin; null until then (and in tests of other modules). */
export let versions: VersionHistory | null = null;
export function setVersionHistory(history: VersionHistory | null): void {
  versions = history;
}
