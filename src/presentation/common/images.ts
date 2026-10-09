/** Open the image picker and pass the selected file to the editor. */
export function pickImage(onPick: (file: File) => void): void {
  const input = document.body.createEl("input", { type: "file", attr: { accept: "image/*" } });
  input.setCssStyles({ display: "none" });
  input.onchange = () => {
    const file = input.files?.[0];
    if (file) onPick(file);
    input.remove();
  };
  input.click();
}
