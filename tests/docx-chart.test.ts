import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { chartAt, chartRun, editedChartRun } from "../src/domain/docx/chart";
import { defaultChartFormat } from "../src/presentation/common/chart-form";
import { fixture } from "./helpers";

const data = { categories: ["Q1", "Q2", "Q3"], series: [{ name: "Sales", values: [10, 20, 15] }] };

describe("Word charts", () => {
  it("inserts a native chart: an inline drawing related to a chart part with its own embedded workbook", async () => {
    const d = await DocxDocument.load(fixture("sample.docx"));
    const run = await chartRun(d, data, { ...defaultChartFormat(), title: "Revenue" });
    expect(run).toContain("wp:inline");
    expect(run).toContain('uri="http://schemas.openxmlformats.org/drawingml/2006/chart"');
    const found = chartAt(d, run)!;
    expect(found.model.data).toEqual(data);
    expect(found.model.format?.title).toBe("Revenue");
    expect(found.path).toMatch(/^word\/charts\/chart\d*\.xml$/);
    const book = d.rels(found.path).find((r) => r.type === "package")!;
    expect(book.target).toMatch(/^word\/embeddings\/.+\.xlsx$/);
    const zip = await JSZip.loadAsync(await d.save());
    expect(await zip.file("[Content_Types].xml")!.async("text")).toContain('ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"');
    const sheet = await (await JSZip.loadAsync(await zip.file(book.target)!.async("uint8array"))).file("xl/worksheets/sheet1.xml")!.async("text");
    expect(sheet).toContain("Sales");
  });

  it("edits into a new chart part (so undo restores the old run, still pointing at the untouched old chart)", async () => {
    const d = await DocxDocument.load(fixture("sample.docx"));
    const run = await chartRun(d, data, defaultChartFormat());
    const before = chartAt(d, run)!;
    const edited = await editedChartRun(d, run, { ...data, series: [{ name: "Sales", values: [1, 2, 3] }] }, { ...defaultChartFormat(), kind: "pie" });
    const after = chartAt(d, edited)!;
    expect(after.path).not.toBe(before.path);
    expect(after.model.data.series[0].values).toEqual([1, 2, 3]);
    expect(after.model.format?.kind).toBe("pie");
    expect(chartAt(d, run)!.model.data.series[0].values).toEqual([10, 20, 15]);
    // The new chart's workbook relationship id matches its externalData reference.
    const ref = d.xml(after.path)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/chart", "externalData")[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    expect(d.rels(after.path).find((r) => r.id === ref)?.type).toBe("package");
  });

  it("is null for a drawing that isn't a chart", async () => {
    const d = await DocxDocument.load(fixture("sample.docx"));
    expect(chartAt(d, `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>x</w:t></w:r>`)).toBeNull();
  });
});
