/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * TypeScript port of libeot (https://github.com/umanwizard/libeot):
 *   Copyright (c) 2013 Brennan T. Vincent <brennanv@email.arizona.edu>
 *   This file is a part of libeot, which is licensed under the MPL license, version 2.0.
 *   Most of the code in the src/lzcomp directory was originally published by Monotype and
 *   licensed under the MTX License, which allowed relicensing it under the MPL 2.0.
 *
 * This file remains under the MPL 2.0 (the rest of Doc Works is GPL-3.0; MPL 2.0 section 3.3
 * permits the combination). Output matches libeot's eot2ttf byte for byte, quirks included.
 */

const fail = (msg: string): never => {
  throw new Error(`EOT: ${msg}`);
};

/** Big-endian reader (util/stream.c). */
class Reader {
  pos = 0;
  constructor(private b: Uint8Array) {}
  seek(p: number) {
    if (p > this.b.length) fail("offset past end of data");
    this.pos = p;
  }
  u8() {
    if (this.pos >= this.b.length) fail("unexpected end of data");
    return this.b[this.pos++];
  }
  peek() {
    const v = this.u8();
    this.pos--;
    return v;
  }
  /** Unsigned big-endian integer of n <= 4 bytes. */
  uint(n: number) {
    let v = 0;
    while (n--) v = v * 256 + this.u8();
    return v;
  }
  s16() {
    return (this.uint(2) << 16) >> 16;
  }
  bytes(n: number) {
    if (this.pos + n > this.b.length) fail("unexpected end of data");
    return this.b.slice(this.pos, (this.pos += n));
  }
}

/** Growable big-endian writer. Like libeot's Stream, `size` only grows on u8/u16 writes, not copies. */
class Writer {
  b: number[] = [];
  pos = 0;
  size = 0;
  u8(v: number) {
    this.b[this.pos++] = v & 255;
    this.size = Math.max(this.size, this.pos);
  }
  u16(v: number) {
    this.u8(v >> 8);
    this.u8(v);
  }
  copy(r: Reader, n: number) {
    for (let i = 0; i < n; i++) this.b[this.pos++] = r.u8();
    this.size = Math.max(this.size, this.pos);
  }
  done() {
    return Uint8Array.from(this.b.slice(0, this.size), (v) => v ?? 0);
  }
}

// ---- LZCOMP (lzcomp/{bitio,ahuff,lzcomp}.c) ----

/** Adaptive Huffman decoder over symbols [0, range). Node 1 is the root; nodes stay sorted by weight. */
function huffman(bit: () => number, range: number) {
  const n = 2 * range;
  const up = new Int32Array(n), left = new Int32Array(n), right = new Int32Array(n);
  const code = new Int32Array(n).fill(-1), weight = new Int32Array(n), index = new Int32Array(range);
  // Heap layout: node i has children 2i, 2i+1; leaves range..2*range-1 hold the symbols, all weight 1.
  for (let i = 2; i < n; i++) {
    up[i] = i >> 1;
    weight[i] = 1;
  }
  for (let i = 1; i < range; i++) {
    left[i] = 2 * i;
    right[i] = 2 * i + 1;
  }
  for (let i = 0; i < range; i++) {
    code[range + i] = i;
    index[i] = range + i;
  }
  for (let i = range - 1; i > 0; i--) weight[i] = weight[2 * i] + weight[2 * i + 1];

  // Swap nodes a and b (everything but their parent links), then re-point children / symbol index.
  const swap = (a: number, b: number) => {
    for (const f of [left, right, code, weight]) [f[a], f[b]] = [f[b], f[a]];
    for (const x of [a, b]) {
      if (code[x] < 0) up[left[x]] = up[right[x]] = x;
      else index[code[x]] = x;
    }
  };
  // Increment a leaf's weight up to the root, first swapping each node with the leftmost node of equal weight.
  const update = (a: number) => {
    for (; a !== 1; a = up[a]) {
      const w = weight[a];
      let b = a - 1;
      if (weight[b] === w) {
        while (weight[--b] === w);
        if (++b > 1) {
          swap(a, b);
          a = b;
        }
      }
      weight[a] = w + 1;
    }
    weight[1]++;
  };

  if (range > 256 && range < 512) {
    // Symbol coder: bias the first copy symbols and DUP2/DUP4.
    update(index[256]);
    update(index[257]);
    for (let i = 0; i < 12; i++) update(index[range - 3]);
    for (let i = 0; i < 6; i++) update(index[range - 2]);
  } else {
    for (let j = 0; j < 2; j++) for (let i = 0; i < range; i++) update(index[i]);
  }
  return () => {
    let a = 1;
    do a = bit() ? right[a] : left[a];
    while (code[a] < 0);
    const symbol = code[a];
    update(a);
    return symbol;
  };
}

// The window starts preloaded with 32x96 (k, j) pairs then each byte value 4 times, so early copies can hit it.
const PRELOAD = 2 * 32 * 96 + 4 * 256;

function unpackLz(data: Uint8Array, version: number): Uint8Array {
  let p = 0, count = 0, buf = 0;
  const bit = () => {
    if (count-- === 0) {
      if (p >= data.length) fail("truncated LZCOMP data");
      buf = data[p++];
      count = 7;
    }
    return (buf <<= 1) & 0x100;
  };
  const value = (bits: number) => {
    let v = 0;
    while (bits--) v = v * 2 + (bit() ? 1 : 0);
    return v;
  };

  const rle = version !== 1 && bit() !== 0;
  const dist = huffman(bit, 8), len = huffman(bit, 8);
  const outLen = value(24);
  // Copy distances use up to `ranges` 3-bit digits; symbols: 256 literals, 8 length codes per range, DUP2/4/6.
  let ranges = 1;
  while (8 ** ranges < outLen) ranges++;
  const dup2 = 256 + 8 * ranges;
  const sym = huffman(bit, dup2 + 3);

  const win = new Uint8Array(PRELOAD + outLen);
  let pos = 0;
  for (let k = 0; k < 32; k++) for (let j = 0; j < 96; j++) {
      win[pos++] = k;
      win[pos++] = j;
    }
  for (let j = 0; j < 256; j++) for (let i = 0; i < 4; i++) win[pos++] = j;

  // Run-length layer (RUNLENGTHCOMP): first byte is the escape; escape, n, v = v repeated n times; escape, 0 = escape.
  const out: number[] = [];
  let state = 3, escape = 0, run = 0;
  const emit = (v: number) => {
    if (!rle || (state === 0 && v !== escape)) out.push(v);
    else if (state === 0) state = 1;
    else if (state === 1 && v) {
      run = v;
      state = 2;
    } else {
      // escape, 0 -> escape; escape, n, v -> v * n; the very first byte sets the escape
      if (state === 1) out.push(escape);
      else if (state === 2) for (let i = 0; i < run; i++) out.push(v);
      else escape = v;
      state = 0;
    }
  };

  const end = PRELOAD + outLen;
  while (pos < end) {
    const s = sym();
    if (s < 256 || s >= dup2) {
      // Literal, or DUP2/DUP4/DUP6: repeat the byte 2/4/6 back.
      const v = s < 256 ? s : win[pos - 2 * (s - dup2 + 1)];
      win[pos++] = v;
      emit(v);
      continue;
    }
    // Copy item: the symbol carries the distance range count and the first 2-bit length digit;
    // bit 2 of each digit means another digit follows from the length coder.
    let bits = s - 256, length = 0, distance = 0;
    const n = (bits >> 3) + 1;
    for (;;) {
      length = length * 4 + (bits & 3);
      if (!(bits & 4)) break;
      bits = len();
    }
    length += 2;
    for (let i = 0; i < n; i++) distance = distance * 8 + dist();
    distance += 1;
    if (distance >= 512) length++;
    const start = pos - distance - length + 1;
    if (start < 0) fail("corrupt LZCOMP copy");
    for (let i = 0; i < length; i++) {
      const v = win[start + i];
      win[pos++] = v;
      emit(v);
    }
  }
  if (pos !== end) fail("corrupt LZCOMP data");
  return Uint8Array.from(out);
}

/** MTX container: version byte, copy limit, offsets of blocks 2 and 3 (24-bit BE); block 1 starts at 10. */
function unpackMtx(buf: Uint8Array): Uint8Array[] {
  const r = new Reader(buf);
  const version = r.u8();
  r.uint(3);
  const offs = [10, r.uint(3), r.uint(3), buf.length];
  return [0, 1, 2].map((i) => {
    if (offs[i] > offs[i + 1] || offs[i + 1] > buf.length) fail("corrupt MTX block offsets");
    return unpackLz(buf.subarray(offs[i], offs[i + 1]), version);
  });
}

// ---- CTF -> TrueType (ctf/parseCTF.c) ----

/** 255UShort: 253 = word follows, 255 = 253 + byte, 254 = 506 + byte. */
function r255us(r: Reader) {
  const c = r.u8();
  return c === 253 ? r.uint(2) : c === 255 ? 253 + r.u8() : c === 254 ? 506 + r.u8() : c;
}

/** 255Short: 253 = word follows, 250 = negate next, 255 = 250 + byte, 254 = 500 + byte. */
function r255s(r: Reader) {
  let c = r.u8();
  if (c === 253) return r.s16();
  const sign = c === 250 ? -1 : 1;
  if (sign < 0) c = r.u8();
  return sign * (c === 255 ? 250 + r.u8() : c === 254 ? 500 + r.u8() : c);
}

/** Expand `count` pushed values (with 0xFB/0xFC hop codes) into PUSHB/PUSHW/NPUSHB/NPUSHW runs. */
function pushes(r: Reader, w: Writer, count: number) {
  const data: number[] = [];
  let word = false, run = 0;
  const dump = () => {
    if (!run) return;
    if (run < 8) w.u8((word ? 0xb8 : 0xb0) | (run - 1));
    else {
      w.u8(word ? 0x41 : 0x40);
      w.u8(run);
    }
    for (const v of data.slice(-run)) {
      if (word) w.u16(v);
      else w.u8(v);
    }
  };
  const put = (v: number) => {
    const isWord = v < 0 || v > 255;
    if (isWord !== word || run === 255) {
      dump();
      word = isWord;
      run = 0;
    }
    data.push(v);
    run++;
  };
  for (let left = count; left > 0; ) {
    const code = r.peek();
    if (code === 0xfb || code === 0xfc) {
      // A B 0xFB C -> A B A C A;  A B 0xFC C D -> A B A C A D A
      const n = code === 0xfb ? 3 : 5;
      if (left < n || data.length < 2) fail("corrupt hop code data");
      left -= n;
      r.pos++;
      const prev = data[data.length - 2];
      put(prev);
      for (let k = 1; k < n; k += 2) {
        put(r255s(r));
        put(prev);
      }
    } else {
      put(r255s(r));
      left--;
    }
  }
  dump();
}

// Triplet encodings (triplet_encodings.c), indexed by flag & 0x7F: [xBits, yBits, deltaX, deltaY, xSign, ySign].
const sgn = (b: number) => (b ? 1 : -1);
const TRIPLETS = Array.from({ length: 128 }, (_, i) => {
  if (i < 10) return [0, 8, 0, (i >> 1) << 8, 0, sgn(i & 1)];
  if (i < 20) return [8, 0, ((i - 10) >> 1) << 8, 0, sgn(i & 1), 0];
  const b = i < 84 ? i - 20 : i < 120 ? i - 84 : i - 120, xs = sgn(b & 1), ys = sgn(b & 2);
  if (i < 84) return [4, 4, 1 + 16 * (b >> 4), 1 + 16 * ((b >> 2) & 3), xs, ys];
  if (i < 120) return [8, 8, 1 + 256 * Math.floor((b >> 2) / 3), 1 + 256 * ((b >> 2) % 3), xs, ys];
  return i < 124 ? [12, 12, 0, 0, xs, ys] : [16, 16, 0, 0, xs, ys];
});

function simpleGlyph(contours: number, bbox: number[] | null, [r, r1, r2]: Reader[], w: Writer) {
  if (contours === 0) return;
  if (contours < 0) fail("corrupt glyph");
  w.u16(contours);
  const bboxAt = w.pos;
  if (bbox) bbox.forEach((v) => w.u16(v));
  else w.pos += 8;
  let total = 1;
  for (let i = 0; i < contours; i++) {
    total += r255us(r);
    w.u16(total - 1);
  }
  const flags = r.bytes(total);
  const xs = new Int16Array(total), ys = new Int16Array(total);
  let x = 0, y = 0, minX = 0x7fff, minY = 0x7fff, maxX = -0x8000, maxY = -0x8000;
  for (let i = 0; i < total; i++) {
    const [xBits, yBits, dX, dY, xSign, ySign] = TRIPLETS[flags[i] & 0x7f];
    const v = r.uint((xBits + yBits) / 8), div = 2 ** yBits;
    xs[i] = xSign * (Math.floor(v / div) + dX);
    ys[i] = ySign * ((v % div) + dY);
    x = ((x + xs[i]) << 16) >> 16;
    y = ((y + ys[i]) << 16) >> 16;
    minX = Math.min(minX, x);
    maxX = Math.max(maxX, x);
    minY = Math.min(minY, y);
    maxY = Math.max(maxY, y);
  }
  const codeAt = w.pos;
  w.pos += 2;
  pushes(r1, w, r255us(r));
  w.copy(r2, r255us(r));
  const codeLen = w.pos - codeAt - 2;
  // Flags: 1 on-curve, 2/4 x/y short, 0x10/0x20 x/y same (or positive when short). No repeat flag.
  const bits = (v: number, first: boolean, short: number, same: number) =>
    !first && v === 0 ? same : v > -256 && v < 0 ? short : v >= 0 && v < 256 ? short | same : 0;
  for (let i = 0; i < total; i++) w.u8((flags[i] & 0x80 ? 0 : 1) | bits(xs[i], !i, 2, 0x10) | bits(ys[i], !i, 4, 0x20));
  for (const c of [xs, ys])
    c.forEach((v, i) => {
      if (i && !v) return;
      if (v > -256 && v < 256) w.u8(Math.abs(v));
      else w.u16(v);
    });
  const end = w.pos;
  w.pos = codeAt;
  w.u16(codeLen);
  if (!bbox) {
    w.pos = bboxAt;
    [minX, minY, maxX, maxY].forEach((v) => w.u16(v));
  }
  w.pos = end;
}

function compositeGlyph([r, r1, r2]: Reader[], w: Writer) {
  w.u16(-1);
  for (let i = 0; i < 4; i++) w.u16(r.s16());
  let f: number;
  do {
    f = r.uint(2);
    w.u16(f);
    // glyph index + args (words if ARG_1_AND_2_ARE_WORDS) + 2x2 / xy scale / scale
    w.copy(r, 2 + (f & 1 ? 4 : 2) + (f & 0x80 ? 8 : f & 0x40 ? 4 : f & 8 ? 2 : 0));
  } while (f & 0x20);
  // WE_HAVE_INSTRUCTIONS: length, then the bytecode (libeot leaves the length out, which browsers reject).
  if (f & 0x100) {
    const lenAt = w.pos;
    w.u16(0);
    pushes(r1, w, r255us(r));
    w.copy(r2, r255us(r));
    const end = w.pos;
    w.pos = lenAt;
    w.u16(end - lenAt - 2);
    w.pos = end;
  }
}

interface Table {
  tag: string;
  offset: number;
  length: number;
  data: Uint8Array;
}

function ctfToTtf(blocks: Uint8Array[]): Uint8Array {
  const rs = blocks.map((b) => new Reader(b)), r = rs[0];
  r.uint(4);
  const count = r.uint(2);
  r.uint(6);
  const tables: Table[] = [];
  for (let i = 0; i < count; i++) {
    const tag = String.fromCharCode(...r.bytes(4));
    r.uint(4); // checksum
    tables.push({ tag, offset: r.uint(4), length: r.uint(4), data: new Uint8Array(0) });
  }
  const find = (tag: string) => tables.filter((t) => t.tag === tag).pop();
  for (const t of tables) {
    if (["loca", "glyf", "hdmx", "VDMX"].includes(t.tag)) continue; // hdmx / VDMX aren't rebuilt (optional; dropped below)
    r.seek(t.offset);
    if (t.tag === "cvt ") {
      // Deltas: <238 literal, 238 word, 239-247 negative and 248-255 positive with a second byte.
      const n = r.uint(2), dv = new DataView((t.data = new Uint8Array(2 * n)).buffer);
      for (let i = 0, last = 0; i < n; i++) {
        const c = r.u8();
        last += c >= 248 ? 238 * (c - 247) + r.u8() : c >= 239 ? -(238 * (c - 239) + r.u8()) : c === 238 ? r.s16() : c;
        dv.setInt16(2 * i, last);
      }
    } else {
      t.data = r.bytes(t.length);
      if (t.tag === "head") {
        if (t.length < 12) fail("malformed head table");
        t.data.fill(0, 8, 12); // checkSumAdjustment, set once the whole font is written
      }
    }
  }
  let loca = find("loca");
  if (!loca) tables.push((loca = { tag: "loca", offset: 0, length: 0, data: new Uint8Array(0) }));
  const [glyf, maxp, head] = ["glyf", "maxp", "head"].map(find);
  if (!glyf || !maxp || !head || !find("hmtx")) fail("missing glyf, maxp, head or hmtx table");
  if (head!.length < 52 || maxp!.length < 32) fail("corrupt head or maxp table");
  const shortLoca = !((head!.data[50] << 8) | head!.data[51]);
  const numGlyphs = (maxp!.data[4] << 8) | maxp!.data[5];

  r.seek(glyf!.offset);
  const w = new Writer(), lw = new Writer();
  const locaEntry = (p: number) => {
    if (!shortLoca) lw.u16(p >>> 16);
    lw.u16(shortLoca ? p >> 1 : p);
  };
  locaEntry(0);
  for (let i = 0; i < numGlyphs; i++) {
    const contours = r.s16();
    if (contours < 0) compositeGlyph(rs, w);
    else if (contours === 0x7fff) simpleGlyph(r.s16(), [r.s16(), r.s16(), r.s16(), r.s16()], rs, w);
    else simpleGlyph(contours, null, rs, w);
    if (w.pos % 2) w.u8(0);
    locaEntry(w.pos);
  }
  glyf!.data = w.done();
  loca.data = lw.done();

  // SFNT container (ctf/SFNTContainer.c): offset table, directory sorted by tag, 4-byte padded tables.
  const kept = tables.filter((t) => t.data.length).sort((a, b) => (a.tag < b.tag ? -1 : 1));
  const n = kept.length, e = 31 - Math.clz32(n);
  const out = new Uint8Array(kept.reduce((s, t) => s + ((t.data.length + 3) & ~3), 12 + 16 * n));
  const dv = new DataView(out.buffer);
  dv.setUint32(0, 0x10000);
  dv.setUint16(4, n);
  dv.setUint16(6, 16 << e);
  dv.setUint16(8, e);
  dv.setUint16(10, n * 16 - (16 << e));
  let pos = 12 + 16 * n;
  let headAt = 0;
  kept.forEach((t, i) => {
    if (t.tag === "head") headAt = pos;
    out.set(t.data, pos);
    let sum = 0;
    for (let p = pos; p < pos + t.data.length; p += 4) sum = (sum + dv.getUint32(p)) >>> 0;
    const d = 12 + 16 * i;
    for (let k = 0; k < 4; k++) out[d + k] = t.tag.charCodeAt(k);
    dv.setUint32(d + 4, sum);
    dv.setUint32(d + 8, pos);
    dv.setUint32(d + 12, t.data.length);
    pos += (t.data.length + 3) & ~3;
  });
  let total = 0;
  for (let p = 0; p < out.length; p += 4) total = (total + dv.getUint32(p)) >>> 0;
  dv.setUint32(headAt + 8, (0xb1b0afba - total) >>> 0);
  return out;
}

// ---- EOT header (EOT.c, writeFontFile.c) ----

/** Convert an Embedded OpenType font to TrueType: plain, XOR-obfuscated and/or MTX-compressed data. */
export function eotToTtf(eot: Uint8Array): Uint8Array {
  const dv = new DataView(eot.buffer, eot.byteOffset, eot.byteLength), len = eot.length;
  if (len <= 12) fail("file is truncated");
  const total = dv.getUint32(0, true), size = dv.getUint32(4, true);
  if (len < ((total - size) >>> 0)) fail("file is truncated");
  const coded = [0x10000, 0x20001, 0x20002].indexOf(dv.getUint32(8, true)) + 1;
  if (!coded) fail("unsupported version");
  if (len < size + 12) fail("corrupt header");
  // Header fields after the 12-byte prefix; offsets below are relative to byte 12.
  const hlen = len - size - 12, u16 = (r: number) => dv.getUint16(12 + r, true);
  if (hlen <= 52 || u16(22) !== 0x504c) fail("corrupt header"); // 0x504C magic
  const SHORT = -1, BIG = -2;
  // Returns the font data offset when the header parses as version v (1..3), SHORT when it runs past the
  // header, BIG when it ends before EOTSize - FontDataSize. libeot uses these to fix a wrong version field.
  const parse = (v: number) => {
    let r = 70; // FamilyNameSize
    const blob = (wide: boolean) => {
      const w = wide ? 4 : 2;
      if (r + w > hlen) return false;
      const n = wide ? dv.getUint32(12 + r, true) : u16(r);
      if (!wide && n % 2) fail("corrupt header string");
      return (r += w + n) <= hlen;
    };
    // Family, style, version, full name (+ root string from v2), each after 2 bytes of padding.
    for (let i = 0; i < (v > 1 ? 5 : 4); i++) {
      if (i) r += 2;
      if (!blob(false)) return SHORT;
    }
    if (v === 3) {
      // RootStringCheckSum, EUDCCodePage, Padding6, SignatureSize + Signature, EUDCFlags, EUDCFontSize + data
      if (r + 8 >= hlen) return SHORT;
      r += 10;
      if (r + 2 >= hlen) return SHORT;
      r += 2 + u16(r);
      if (r + 4 >= hlen) return SHORT;
      r += 4;
      if (!blob(true)) return SHORT;
    }
    return r + 12 < ((total - size) >>> 0) ? BIG : r + 12;
  };
  let dir = 0, v = coded, off: number;
  while ((off = parse(v)) < 0) {
    const step = off === BIG ? 1 : -1;
    if (dir === -step || v + step < 1 || v + step > 3) fail("corrupt header");
    dir = step;
    v += step;
  }

  const flags = dv.getUint32(12, true);
  let data = eot.slice(off, off + size);
  if (flags & 0x10000000) data = data.map((b) => b ^ 0x50); // TTEMBED_XORENCRYPTDATA
  return flags & 4 ? ctfToTtf(unpackMtx(data)) : data; // TTEMBED_TTCOMPRESSED
}
