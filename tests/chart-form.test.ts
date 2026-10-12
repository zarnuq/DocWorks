import { describe, expect, it } from "vitest";
import { chartFormatForm, defaultChartFormat } from "../src/presentation/common/chart-form";

const data = { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] };
const visible = (host: HTMLElement, label: string) => {
  const row = host.querySelector(`[aria-label="${label}"]`)!.parentElement as HTMLElement;
  return row.style.display !== "none";
};

describe("chart editor form", () => {
  it("hides axis settings for pie and doughnut charts, as Google's editor does", () => {
    const host = document.body.createDiv();
    const format = { ...defaultChartFormat(), kind: "doughnut" as const };
    chartFormatForm(host, format, data, ["#4285F4"]);
    for (const label of ["Gridlines", "Category axis title", "Value axis title", "Value axis minimum", "Value axis maximum"]) expect(visible(host, label)).toBe(false);
    const picker = host.querySelector<HTMLSelectElement>('[aria-label="Chart type"]')!;
    picker.value = "column:none";
    picker.dispatchEvent(new Event("change"));
    expect(visible(host, "Value axis title")).toBe(true);
    expect(visible(host, "Data labels")).toBe(true);
  });
});
