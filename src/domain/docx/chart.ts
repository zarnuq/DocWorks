/** Native Word charts: an inline drawing whose graphic is a chart part with its own embedded workbook. */
import { NS } from "../../core/utils/xml";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import type { DocxDocument } from "../../infrastructure/docx/docx-document";
import { addChartPart, chartData, ChartData, ChartModel, chartWorkbook, CHART_NS, CHART_TYPE, prepareChartEdit, WORKBOOK_TYPE } from "../pptx/chart";
import { chartPalette, ChartFormat, readChartFormat, setChartFormat } from "../pptx/chart-format";
import type { ColorContext } from "../pptx/color";
import { themeColors } from "../xlsx/styles";

/** Word's default chart size: 6 × 3.5 inches. */
const SIZE = { cx: 5486400, cy: 3200400 };

/** The document theme's colours, for default series colours. */
export function docxChartColors(docx: DocxDocument): ColorContext {
  return { scheme: themeColors(docx.themeXml()), map: { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" } };
}

const chartRef = (run: Document) => run.getElementsByTagNameNS(CHART_NS, "chart")[0] ?? null;

/** The chart a run's drawing shows (relationships of the main part), or null when it isn't a chart. */
export function chartAt(docx: DocxDocument, runXml: string): { path: string; model: ChartModel; size: { cx: number; cy: number } } | null {
  if (!runXml.includes(CHART_NS)) return null;
  const run = parseXml(runXml);
  const ref = run && chartRef(run);
  const path = ref && docx.relTarget(docx.mainPath, ref.getAttributeNS(NS.r, "id") ?? "");
  const xml = path ? docx.xml(path) : null;
  if (!path || !xml) return null;
  const extent = run.getElementsByTagNameNS(NS.wp, "extent")[0];
  const colors = docxChartColors(docx);
  const type = Array.from(xml.getElementsByTagNameNS(CHART_NS, "plotArea")[0]?.children ?? []).find((e) => e.localName.endsWith("Chart"))?.localName ?? "unknown";
  return {
    path,
    model: { path, type, data: chartData(xml), format: readChartFormat(xml, colors), palette: chartPalette(colors) },
    size: { cx: Number(extent?.getAttribute("cx")) || SIZE.cx, cy: Number(extent?.getAttribute("cy")) || SIZE.cy },
  };
}

/** Insert › Chart: a new chart part and the run showing it inline, as wide as Word's default or the text. */
export async function chartRun(docx: DocxDocument, data: ChartData, format: ChartFormat): Promise<string> {
  const workbook = await chartWorkbook(data);
  const path = addChartPart(docx, "word", data, "barChart", workbook, (f) => Object.assign(f, structuredClone(format)));
  setChartFormat(docx.xml(path)!, format, docxChartColors(docx));
  const rId = docx.relate(docx.mainPath, "chart", path);
  const cx = Math.min(SIZE.cx, docx.textWidth() * 635);
  const cy = Math.round((SIZE.cy * cx) / SIZE.cx);
  const id = docx.nextDocPrId();
  return (
    `<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:c="${CHART_NS}" xmlns:r="${NS.r}"><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="Chart ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="${rId}"/></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r>`
  );
}

/**
 * The run after editing its chart: the edit goes into a new chart part (and workbook), so undoing the run's change
 * brings back the old chart untouched. Other relationships (styles, colours) are shared with the old part.
 */
export async function editedChartRun(docx: DocxDocument, runXml: string, data: ChartData, format: ChartFormat): Promise<string> {
  const found = chartAt(docx, runXml);
  if (!found) throw new Error("Not a chart");
  const update = await prepareChartEdit(docx, found.path, data);
  setChartFormat(update.xml, format, docxChartColors(docx));
  const path = docx.uniquePath("word/charts/chart", ".xml");
  const ids = new Map<string, string>();
  for (const rel of docx.rels(found.path).filter((r) => !r.external)) {
    let target = rel.target;
    if (rel.target === update.workbookPath && update.workbook) {
      target = docx.uniquePath("word/embeddings/chartData", ".xlsx");
      docx.putBinary(target, update.workbook, WORKBOOK_TYPE);
    }
    ids.set(rel.id, docx.relate(path, rel.uri, target));
  }
  // Point the copy's references at its own relationship ids.
  for (const el of Array.from(update.xml.getElementsByTagName("*"))) {
    for (const attr of Array.from(el.attributes)) {
      if (attr.namespaceURI === NS.r && ids.has(attr.value)) attr.value = ids.get(attr.value)!;
    }
  }
  docx.putXml(path, update.xml, CHART_TYPE);
  const run = parseXml(runXml)!;
  chartRef(run).setAttributeNS(NS.r, "r:id", docx.relate(docx.mainPath, "chart", path));
  return new XMLSerializer().serializeToString(run);
}
