/*
 * The equation editor both editors share: LaTeX with a live preview and Google's symbol palette. Documents store
 * the result as OMML in the text, slides as PowerPoint's a14:m math.
 */
import { App, Modal, Setting } from "obsidian";
import { DOMSerializer } from "prosemirror-model";
import { latexToOmml, mathSpec } from "../../domain/docx/math";

/** Equation palette rows: label, LaTeX to insert. */
const EQUATION_PALETTE: [string, string][][] = [
  [["a/b", "\\frac{}{}"], ["√", "\\sqrt{}"], ["ⁿ√", "\\sqrt[]{}"], ["xⁿ", "^{}"], ["xₙ", "_{}"], ["∑", "\\sum_{}^{}{}"], ["∫", "\\int_{}^{}{}"], ["lim", "\\lim_{}{}"], ["( )", "\\left(\\right)"], ["[ ]", "\\left[\\right]"], ["x̂", "\\hat{}"], ["x̄", "\\overline{}"], ["vec", "\\vec{}"], ["▦", "\\begin{matrix}{}&\\\\&\\end{matrix}"]],
  [["α", "\\alpha "], ["β", "\\beta "], ["γ", "\\gamma "], ["δ", "\\delta "], ["ε", "\\varepsilon "], ["θ", "\\theta "], ["λ", "\\lambda "], ["μ", "\\mu "], ["π", "\\pi "], ["σ", "\\sigma "], ["φ", "\\varphi "], ["ω", "\\omega "], ["Δ", "\\Delta "], ["Σ", "\\Sigma "], ["Ω", "\\Omega "]],
  [["±", "\\pm "], ["×", "\\times "], ["÷", "\\div "], ["·", "\\cdot "], ["∞", "\\infty "], ["∂", "\\partial "], ["∇", "\\nabla "], ["≤", "\\leq "], ["≥", "\\geq "], ["≠", "\\neq "], ["≈", "\\approx "], ["≡", "\\equiv "], ["∈", "\\in "], ["⊂", "\\subset "], ["∪", "\\cup "], ["∩", "\\cap "], ["∀", "\\forall "], ["∃", "\\exists "], ["→", "\\to "], ["⇒", "\\Rightarrow "], ["⇔", "\\Leftrightarrow "]],
];

/**
 * Edit `latex` (null: a new equation). Resolves the LaTeX to use, "" for Delete (editing only), null when
 * cancelled.
 */
export function equationDialog(app: App, latex: string | null): Promise<string | null> {
  return new Promise((resolve) => {
    let result: string | null = null;
    const editing = latex !== null;
    const modal = new Modal(app);
    modal.titleEl.setText(editing ? "Edit equation" : "Insert equation");
    modal.modalEl.addClass("docx-equation-modal");
    const input = modal.contentEl.createEl("textarea", {
      cls: "docx-equation-input",
      attr: { rows: "3", spellcheck: "false", placeholder: "Type LaTeX, e.g. \\frac{a}{b}, x^2, \\alpha, \\sum_{i=1}^{n}" },
    });
    input.value = latex ?? "";
    const preview = modal.contentEl.createDiv({ cls: "docx-equation-preview" });
    const draw = () => preview.replaceChildren(DOMSerializer.renderSpec(document, mathSpec(latexToOmml(input.value))).dom);
    input.addEventListener("input", draw);
    // Symbols and structures, as in Google Docs' equation toolbar; the cursor lands in the first empty box.
    const insert = (text: string) => {
      const at = input.selectionStart;
      input.setRangeText(text, at, input.selectionEnd, "end");
      const box = text.indexOf("{}");
      if (box >= 0) input.setSelectionRange(at + box + 1, at + box + 1);
      input.focus();
      draw();
    };
    const palette = modal.contentEl.createDiv({ cls: "docx-equation-palette" });
    for (const row of EQUATION_PALETTE) {
      const line = palette.createDiv({ cls: "docx-equation-row" });
      for (const [label, text] of row) line.createEl("button", { text: label, attr: { "aria-label": text.trim() } }).addEventListener("click", () => insert(text));
    }
    const apply = (value: string) => {
      result = value;
      modal.close();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        apply(input.value);
      }
    });
    const buttons = new Setting(modal.contentEl).addButton((b) =>
      b
        .setButtonText(editing ? "Save" : "Insert")
        .setCta()
        .onClick(() => apply(input.value))
    );
    if (editing) buttons.addButton((b) => b.setButtonText("Delete").onClick(() => apply("")));
    buttons.addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
    draw();
    input.focus();
  });
}
