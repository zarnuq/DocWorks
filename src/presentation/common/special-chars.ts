import { App, Modal } from "obsidian";

/** Google's Insert › Special characters, a short list: "char name" per line, by category. */
const CATEGORIES: [string, string][] = [
  ["Symbols", `© Copyright sign
® Registered sign
™ Trade mark sign
§ Section sign
¶ Pilcrow sign
† Dagger
‡ Double dagger
• Bullet
· Middle dot
… Horizontal ellipsis
° Degree sign
′ Prime
″ Double prime
‰ Per mille sign
№ Numero sign
℃ Degree Celsius
℉ Degree Fahrenheit
✓ Check mark
✔ Heavy check mark
✗ Ballot x
✘ Heavy ballot x
☐ Ballot box
☑ Ballot box with check
☒ Ballot box with x
★ Black star
☆ White star
♠ Black spade suit
♣ Black club suit
♥ Black heart suit
♦ Black diamond suit
♪ Eighth note
☎ Black telephone
✉ Envelope
☀ Sun
☁ Cloud
☂ Umbrella
❄ Snowflake
⚠ Warning sign`],
  ["Punctuation", `– En dash
— Em dash
‘ Left single quotation mark
’ Right single quotation mark
“ Left double quotation mark
” Right double quotation mark
« Left-pointing double angle quotation mark
» Right-pointing double angle quotation mark
‹ Single left-pointing angle quotation mark
› Single right-pointing angle quotation mark
¡ Inverted exclamation mark
¿ Inverted question mark
‽ Interrobang
⁂ Asterism
  No-break space`],
  ["Arrows", `← Leftwards arrow (arrow left)
→ Rightwards arrow (arrow right)
↑ Upwards arrow (arrow up)
↓ Downwards arrow (arrow down)
↔ Left right arrow
↕ Up down arrow
↖ North west arrow
↗ North east arrow
↘ South east arrow
↙ South west arrow
⇐ Leftwards double arrow
⇒ Rightwards double arrow (implies)
⇔ Left right double arrow (if and only if)
↩ Leftwards arrow with hook
↪ Rightwards arrow with hook
↻ Clockwise open circle arrow
➜ Heavy round-tipped rightwards arrow
▲ Black up-pointing triangle
▼ Black down-pointing triangle
◀ Black left-pointing triangle
▶ Black right-pointing triangle`],
  ["Math", `± Plus-minus sign
× Multiplication sign
÷ Division sign
− Minus sign
≈ Almost equal to
≠ Not equal to
≡ Identical to
≤ Less-than or equal to
≥ Greater-than or equal to
≪ Much less-than
≫ Much greater-than
∞ Infinity
√ Square root
∛ Cube root
∑ N-ary summation
∏ N-ary product
∫ Integral
∂ Partial differential
∆ Increment
∇ Nabla
∈ Element of
∉ Not an element of
⊂ Subset of
⊃ Superset of
∪ Union
∩ Intersection
∅ Empty set
∀ For all
∃ There exists
¬ Not sign
∧ Logical and
∨ Logical or
∴ Therefore
∵ Because
∝ Proportional to
∠ Angle
⊥ Perpendicular
∥ Parallel to
½ Vulgar fraction one half
⅓ Vulgar fraction one third
¼ Vulgar fraction one quarter
¾ Vulgar fraction three quarters
² Superscript two
³ Superscript three
¹ Superscript one
ⁿ Superscript n
₀ Subscript zero
₁ Subscript one
₂ Subscript two`],
  ["Greek", `α Greek small letter alpha
Α Greek capital letter alpha
β Greek small letter beta
γ Greek small letter gamma
Γ Greek capital letter gamma
δ Greek small letter delta
Δ Greek capital letter delta
ε Greek small letter epsilon
ζ Greek small letter zeta
η Greek small letter eta
θ Greek small letter theta
Θ Greek capital letter theta
λ Greek small letter lambda
Λ Greek capital letter lambda
μ Greek small letter mu
π Greek small letter pi
Π Greek capital letter pi
ρ Greek small letter rho
σ Greek small letter sigma
Σ Greek capital letter sigma
τ Greek small letter tau
φ Greek small letter phi
Φ Greek capital letter phi
χ Greek small letter chi
ψ Greek small letter psi
Ψ Greek capital letter psi
ω Greek small letter omega
Ω Greek capital letter omega`],
  ["Currency", `$ Dollar sign
¢ Cent sign
€ Euro sign
£ Pound sign
¥ Yen sign
₹ Indian rupee sign
₽ Ruble sign
₩ Won sign
₿ Bitcoin sign
₺ Turkish lira sign
₫ Dong sign
฿ Thai baht
¤ Currency sign`],
  ["Latin", `á Latin small letter a with acute
à Latin small letter a with grave
â Latin small letter a with circumflex
ä Latin small letter a with diaeresis
ã Latin small letter a with tilde
å Latin small letter a with ring above
æ Latin small letter ae
ç Latin small letter c with cedilla
é Latin small letter e with acute
è Latin small letter e with grave
ê Latin small letter e with circumflex
ë Latin small letter e with diaeresis
í Latin small letter i with acute
ï Latin small letter i with diaeresis
ñ Latin small letter n with tilde
ó Latin small letter o with acute
ô Latin small letter o with circumflex
ö Latin small letter o with diaeresis
ø Latin small letter o with stroke
œ Latin small ligature oe
ß Latin small letter sharp s
ú Latin small letter u with acute
ü Latin small letter u with diaeresis
ý Latin small letter y with acute
É Latin capital letter e with acute
Ñ Latin capital letter n with tilde
Ö Latin capital letter o with diaeresis
Ü Latin capital letter u with diaeresis`],
  ["Emoji", `😀 Grinning face
😂 Face with tears of joy
😊 Smiling face with smiling eyes
😍 Smiling face with heart-eyes
😎 Smiling face with sunglasses
🤔 Thinking face
😢 Crying face
😮 Face with open mouth
👍 Thumbs up
👎 Thumbs down
👏 Clapping hands
🙏 Folded hands
👋 Waving hand
💪 Flexed biceps
❤️ Red heart
🔥 Fire
⭐ Star
✨ Sparkles
🎉 Party popper
✅ Check mark button
❌ Cross mark
❗ Exclamation mark
❓ Question mark
💡 Light bulb
📌 Pushpin
📎 Paperclip
📅 Calendar
📈 Chart increasing
📉 Chart decreasing
📝 Memo
🔒 Locked
🔑 Key
⏰ Alarm clock
🚀 Rocket`],
];

const ALL = CATEGORIES.flatMap(([category, lines]) =>
  lines.split("\n").map((line) => {
    const at = line.indexOf(" ");
    return { category, char: line.slice(0, at), name: line.slice(at + 1) };
  })
);

/** Insert › Special characters: a searchable grid; every click hands a character to `pick`, the dialog stays open as Google's does. */
export function specialCharacters(app: App, pick: (char: string) => void): void {
  const modal = new Modal(app);
  modal.titleEl.setText("Special characters");
  const bar = modal.contentEl.createDiv({ cls: "dw-chars-bar" });
  const category = bar.createEl("select", { cls: "dropdown", attr: { "aria-label": "Category" } });
  category.createEl("option", { text: "All", value: "" });
  for (const [name] of CATEGORIES) category.createEl("option", { text: name, value: name });
  const search = bar.createEl("input", { type: "search", attr: { placeholder: "Search by name", "aria-label": "Search by name" } });
  const grid = modal.contentEl.createDiv({ cls: "dw-chars" });
  const draw = () => {
    grid.empty();
    const words = search.value.toLowerCase().split(/\s+/).filter(Boolean);
    for (const c of ALL) {
      if (category.value && c.category !== category.value) continue;
      const name = c.name.toLowerCase();
      if (!words.every((w) => name.includes(w))) continue;
      const b = grid.createEl("button", { cls: "dw-char", text: c.char, attr: { "aria-label": c.name } });
      b.addEventListener("click", () => {
        pick(c.char);
        // Inserting can focus the document (a cell editor, a text box): keep focus here so Escape closes the dialog.
        b.focus();
      });
    }
    if (!grid.childElementCount) grid.createDiv({ cls: "dw-chars-empty", text: "No matches" });
  };
  search.addEventListener("input", draw);
  category.addEventListener("change", draw);
  draw();
  modal.open();
  search.focus();
}
