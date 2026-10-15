/** Sequence and word diffs for comparing a kept version with the file now. */

export type DiffOp<T> = { kind: "same"; a: T; b: T } | { kind: "removed"; a: T } | { kind: "added"; b: T };

/** Above this many cells the middle isn't aligned: it's all removed, then all added. */
const MAX_CELLS = 4_000_000;

/** The longest common subsequence of `a` and `b` kept, the rest removed (first) or added. */
export function diff<T>(a: T[], b: T[], eq: (x: T, y: T) => boolean = (x, y) => x === y): DiffOp<T>[] {
  let start = 0;
  while (start < a.length && start < b.length && eq(a[start], b[start])) start++;
  let endA = a.length, endB = b.length;
  while (endA > start && endB > start && eq(a[endA - 1], b[endB - 1])) {
    endA--;
    endB--;
  }
  const ops: DiffOp<T>[] = [];
  for (let i = 0; i < start; i++) ops.push({ kind: "same", a: a[i], b: b[i] });
  const n = endA - start, m = endB - start;
  if (n * m > MAX_CELLS) {
    // ponytail: no alignment for huge rewrites; Myers' diff if that ever matters.
    for (let i = start; i < endA; i++) ops.push({ kind: "removed", a: a[i] });
    for (let j = start; j < endB; j++) ops.push({ kind: "added", b: b[j] });
  } else {
    // lcs[i * (m + 1) + j]: the common length of a[start + i..endA) and b[start + j..endB).
    const lcs = new Uint32Array((n + 1) * (m + 1));
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        lcs[i * (m + 1) + j] = eq(a[start + i], b[start + j]) ? lcs[(i + 1) * (m + 1) + j + 1] + 1 : Math.max(lcs[(i + 1) * (m + 1) + j], lcs[i * (m + 1) + j + 1]);
      }
    }
    let i = 0, j = 0;
    while (i < n || j < m) {
      if (i < n && j < m && eq(a[start + i], b[start + j])) ops.push({ kind: "same", a: a[start + i++], b: b[start + j++] });
      else if (j >= m || (i < n && lcs[(i + 1) * (m + 1) + j] >= lcs[i * (m + 1) + j + 1])) ops.push({ kind: "removed", a: a[start + i++] });
      else ops.push({ kind: "added", b: b[start + j++] });
    }
  }
  for (let i = endA; i < a.length; i++) ops.push({ kind: "same", a: a[i], b: b[i - endA + endB] });
  return ops;
}

export interface WordChange {
  kind: "same" | "removed" | "added";
  text: string;
}

/** `a` → `b` by words (spaces kept as their own pieces), runs of one kind joined. */
export function diffWords(a: string, b: string): WordChange[] {
  const words = (s: string) => s.split(/(\s+)/).filter(Boolean);
  const out: WordChange[] = [];
  for (const op of diff(words(a), words(b))) {
    const text = op.kind === "added" ? op.b : op.a;
    const last = out[out.length - 1];
    if (last?.kind === op.kind) last.text += text;
    else out.push({ kind: op.kind, text });
  }
  return out;
}
