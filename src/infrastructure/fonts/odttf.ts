/**
 * Word's embedded fonts (word/fonts/*.odttf) are TrueType files whose first 32 bytes are XORed with the
 * font's key, a GUID, read as 16 bytes from its last hex pair to its first (ECMA-376 Part 2, §15.2.2).
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, "");
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) return data;
  const key = Array.from({ length: 16 }, (_, i) => parseInt(hex.slice(30 - 2 * i, 32 - 2 * i), 16));
  const out = data.slice();
  for (let i = 0; i < Math.min(32, out.length); i++) out[i] ^= key[i % 16];
  return out;
}
