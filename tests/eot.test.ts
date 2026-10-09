import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { join } from "path";
import { eotToTtf } from "../src/infrastructure/fonts/eot";

const fixture = (file: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", "eot", file)));

/** Minimal version 0x00020001 EOT around raw font data, all name strings empty. */
function plainEot(font: Uint8Array, flags = 0): Uint8Array {
  const out = new Uint8Array(100 + font.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, out.length, true);
  dv.setUint32(4, font.length, true);
  dv.setUint32(8, 0x00020001, true);
  dv.setUint32(12, flags, true);
  dv.setUint16(34, 0x504c, true);
  out.set(flags & 0x10000000 ? font.map((b) => b ^ 0x50) : font, 100);
  return out;
}

/** Table tag -> data of an SFNT font, in directory order. */
function tables(font: Uint8Array): Map<string, Uint8Array> {
  const dv = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const out = new Map<string, Uint8Array>();
  for (let i = 0; i < dv.getUint16(4); i++) {
    const d = 12 + 16 * i;
    const tag = String.fromCharCode(...font.subarray(d, d + 4));
    out.set(tag, font.subarray(dv.getUint32(d + 8), dv.getUint32(d + 8) + dv.getUint32(d + 12)));
  }
  return out;
}

/** Each glyph's bytes, from glyf and loca. */
function glyphs(t: Map<string, Uint8Array>): Uint8Array[] {
  const loca = t.get("loca")!, glyf = t.get("glyf")!;
  const short = !new DataView(t.get("head")!.buffer, t.get("head")!.byteOffset).getUint16(50);
  const dv = new DataView(loca.buffer, loca.byteOffset, loca.byteLength);
  const at = (i: number) => (short ? dv.getUint16(2 * i) * 2 : dv.getUint32(4 * i));
  const n = loca.length / (short ? 2 : 4) - 1;
  return Array.from({ length: n }, (_, i) => glyf.subarray(at(i), at(i + 1)));
}

/** End of a composite glyph's data as read by a strict parser (like the browser's): components, then the instructions with their length. */
function compositeEnd(g: Uint8Array): number {
  const dv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  let p = 10, flags: number;
  do {
    flags = dv.getUint16(p);
    p += 4 + (flags & 1 ? 4 : 2) + (flags & 0x80 ? 8 : flags & 0x40 ? 4 : flags & 8 ? 2 : 0);
  } while (flags & 0x20);
  if (flags & 0x100) p += 2 + dv.getUint16(p);
  return p;
}

describe("eotToTtf", () => {
  // References are libeot's eot2ttf output for these MTX-compressed Google Slides exports. libeot's own output
  // isn't a valid font (browsers reject it): it drops composite glyphs' instruction length and the end of glyf.
  for (const name of ["Average-regular", "Oswald-regular", "Oswald-bold"]) {
    describe(name, () => {
      const font = eotToTtf(fixture(`${name}.eot`));
      const ours = tables(font);
      const ref = tables(fixture(`${name}.ttf`));

      it("keeps every other table exactly as libeot decodes it", () => {
        for (const [tag, data] of ref) {
          if (["glyf", "loca", "head", "hdmx", "VDMX"].includes(tag)) continue;
          expect(Buffer.from(ours.get(tag)!).equals(Buffer.from(data)), tag).toBe(true);
        }
        const head = ours.get("head")!.slice();
        head.fill(0, 8, 12);
        expect(Buffer.from(head).equals(Buffer.from(ref.get("head")!))).toBe(true);
      });

      it("decodes every simple glyph exactly as libeot does, and gives composite glyphs their instruction length", () => {
        const a = glyphs(ours), b = glyphs(ref);
        expect(a).toHaveLength(b.length);
        a.forEach((g, i) => {
          const composite = g.length >= 2 && new DataView(g.buffer, g.byteOffset).getInt16(0) < 0;
          if (!composite) expect(Buffer.from(g).equals(Buffer.from(b[i])), `glyph ${i}`).toBe(true);
          else expect(compositeEnd(g), `glyph ${i}`).toBeLessThanOrEqual(g.length);
        });
      });

      it("is a well-formed font: glyf matches loca, sorted directory, no empty tables, valid checksum", () => {
        const g = glyphs(ours);
        const last = g[g.length - 1];
        expect(last.byteOffset + last.length - ours.get("glyf")!.byteOffset).toBe(ours.get("glyf")!.length);
        const tags = [...ours.keys()];
        expect(tags).toEqual([...tags].sort());
        expect([...ours.values()].every((t) => t.length > 0)).toBe(true);
        const dv = new DataView(font.buffer, font.byteOffset, font.byteLength);
        let sum = 0;
        for (let p = 0; p < font.length; p += 4) sum = (sum + dv.getUint32(p)) >>> 0;
        expect(sum).toBe(0xb1b0afba);
      });
    });
  }

  it("unwraps plain and XOR-obfuscated font data", () => {
    const ttf = fixture("Average-regular.ttf");
    expect(eotToTtf(plainEot(ttf))).toEqual(ttf);
    expect(eotToTtf(plainEot(ttf, 0x10000000))).toEqual(ttf);
  });

  it("throws on corrupt input", () => {
    expect(() => eotToTtf(new Uint8Array(40))).toThrow(/EOT/);
    const eot = fixture("Oswald-bold.eot");
    expect(() => eotToTtf(eot.subarray(0, eot.length >> 1))).toThrow(/EOT/);
    const bad = eot.slice();
    bad[34] = 0;
    expect(() => eotToTtf(bad)).toThrow(/EOT/);
  });
});
