/** Chart caches and their embedded workbook are edited together; unrelated chart formatting is retained. */
import JSZip from "jszip";
import { NS, escapeXml } from "../../core/utils/xml";
import { parseXml, resolvePath } from "../../infrastructure/ooxml/ooxml-package";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { nextShapeId } from "./slide-edit";
import { axesXml, ChartFormat, readChartFormat, setChartFormat } from "./chart-format";
export const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export interface ChartData {
  categories: string[];
  series: {
    name: string;
    values: number[];
  }[];
}
export interface ChartModel {
  path: string;
  type: string;
  data: ChartData;
  format?: ChartFormat;
  /** Default series colours (the theme's accents). */
  palette?: string[];
}
const elements = (e: Element | Document, name: string) => Array.from(e.getElementsByTagNameNS(CHART_NS, name));
const first = (e: Element | Document, name: string) => elements(e, name)[0];
const child = (e: Element, name: string) => Array.from(e.children).find(c => c.namespaceURI === CHART_NS && c.localName === name);
const points = (e?: Element) => e ? elements(e, "pt").sort((a, b) => Number(a.getAttribute("idx")) - Number(b.getAttribute("idx"))).map(p => first(p, "v")?.textContent ?? "") : [];
export function chartData(xml: Document): ChartData {
  const series = elements(xml, "ser");
  const category = series[0] && (child(series[0], "cat") ?? child(series[0], "xVal"));
  return { categories: points(category), series: series.map((ser, i) => ({ name: points(child(ser, "tx"))[0] ?? first(child(ser, "tx") ?? ser, "v")?.textContent ?? `Series ${i + 1}`, values: points(child(ser, "val") ?? child(ser, "yVal")).map(Number) })) };
}
function validate(data: ChartData): void {
  if (!data.categories.length || !data.series.length || data.series.some(s => s.values.length !== data.categories.length || s.values.some(v => !Number.isFinite(v))))
    throw new Error("Every chart series needs one finite number per category");
}
const col = (index: number): string => index < 26 ? String.fromCharCode(65 + index) : col(Math.floor(index / 26) - 1) + col(index % 26);
const column = (name: string) => [...name].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
function cellXml(ref: string, value: string | number): string { return typeof value === "number" ? `<c r="${ref}"><v>${value}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`; }
export async function chartWorkbook(data: ChartData): Promise<Uint8Array> {
  validate(data);
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<Types xmlns="${NS.contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
  zip.file("_rels/.rels", `<Relationships xmlns="${NS.pkgRels}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `<workbook xmlns="${S}" xmlns:r="${NS.r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="${NS.pkgRels}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  const rows = [["", ...data.series.map(s => s.name)], ...data.categories.map((name, i) => [name, ...data.series.map(s => s.values[i])])];
  zip.file("xl/worksheets/sheet1.xml", `<worksheet xmlns="${S}"><sheetData>${rows.map((row, i) => `<row r="${i + 1}">${row.map((v, j) => cellXml(col(j) + (i + 1), v)).join("")}</row>`).join("")}</sheetData></worksheet>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
const cache = (values: (string | number)[], numeric: boolean) => `<c:${numeric ? "numCache" : "strCache"}>${numeric ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`).join("")}</c:${numeric ? "numCache" : "strCache"}>`;
/** Add a chart of `type` (c:barChart...) to `slide`, then apply `format` (kind, stacking, title...) on top. */
export function addChart(doc: PptxDocument, slide: string, data: ChartData, type: string, workbook: Uint8Array, format?: (f: ChartFormat) => void): Element {
  validate(data);
  if (!["barChart", "lineChart", "pieChart", "doughnutChart", "areaChart"].includes(type))
    throw new Error("Unsupported chart type");
  const path = doc.uniquePath("ppt/charts/chart", ".xml");
  const book = doc.uniquePath("ppt/embeddings/chartData", ".xlsx");
  doc.putBinary(book, workbook, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const rid = doc.addSlideRel(path, "package", book);
  const ser = data.series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${col(i + 1)}$1</c:f>${cache([s.name], false)}</c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${data.categories.length + 1}</c:f>${cache(data.categories, false)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${col(i + 1)}$2:$${col(i + 1)}$${data.categories.length + 1}</c:f>${cache(s.values, true)}</c:numRef></c:val></c:ser>`).join("");
  const axes = type === "pieChart" || type === "doughnutChart" ? "" : '<c:axId val="1"/><c:axId val="2"/>';
  const axisXml = axes ? axesXml(1, 2) : '';
  doc.putXml(path, parseXml(`<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:chart><c:plotArea><c:layout/><c:${type}>${type === "barChart" ? '<c:barDir val="col"/><c:grouping val="clustered"/>' : type === "lineChart" || type === "areaChart" ? '<c:grouping val="standard"/>' : ''}${ser}${type === "doughnutChart" ? '<c:holeSize val="50"/>' : ''}${axes}</c:${type}>${axisXml}</c:plotArea><c:legend><c:legendPos val="b"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="${rid}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`)!, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
  if (format) {
    const xml = doc.xml(path)!;
    const f = readChartFormat(xml);
    format(f);
    setChartFormat(xml, f);
  }
  const xml = doc.xml(slide)!;
  const id = nextShapeId(xml);
  const chartRel = doc.addSlideRel(slide, "chart", path);
  const { cx, cy } = doc.slideSize;
  const frame = xml.importNode(parseXml(`<p:graphicFrame xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:c="${CHART_NS}" xmlns:r="${NS.r}"><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round(cx * .15)}" y="${Math.round(cy * .15)}"/><a:ext cx="${Math.round(cx * .7)}" cy="${Math.round(cy * .7)}"/></p:xfrm><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="${chartRel}"/></a:graphicData></a:graphic></p:graphicFrame>`)!.documentElement, true);
  const tree = xml.getElementsByTagNameNS(NS.p, "spTree")[0];
  tree.insertBefore(frame, Array.from(tree.children).find(e => e.localName === "extLst") ?? null);
  return frame;
}
/** Apply values to the original referenced cells, retaining other sheets, styles and workbook parts. */
export async function prepareChartEdit(doc: PptxDocument, path: string, data: ChartData): Promise<{
  xml: Document;
  workbookPath: string | null;
  workbook: Uint8Array | null;
}> {
  validate(data);
  const xml = doc.xml(path)!.cloneNode(true) as Document;
  const series = elements(xml, "ser");
  if (series.length !== data.series.length)
    throw new Error("Keep the imported chart's series count");
  const external = first(xml, "externalData")?.getAttributeNS(NS.r, "id");
  const rel = external ? doc.rels(path).find(r => r.id === external) : null;
  if (rel?.external)
    throw new Error("This chart is linked to an external workbook. Embed its data in Office before editing it here.");
  const workbookPath = rel?.target ?? null;
  const zip = workbookPath ? await JSZip.loadAsync(await doc.media(workbookPath)) : null;
  const sheets = new Map<string, Document>();
  const wb = zip && parseXml(await zip.file("xl/workbook.xml")!.async("text"));
  const wbRels = zip && parseXml(await zip.file("xl/_rels/workbook.xml.rels")!.async("text"));
  const setValues = async (source: Element | undefined, values: (string | number)[]) => {
    if (!source)
      return;
    const ref = first(source, "f");
    const formula = ref?.textContent;
    const match = formula && /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(formula);
    if (ref && !match)
      throw new Error("Chart data uses a named or unsupported range; use a cell range in Office first");
    const numeric = !!first(source, "numRef") || !!first(source, "numLit");
    if (numeric && values.some(v => !Number.isFinite(Number(v))))
      throw new Error("This chart needs numeric categories");
    if (zip && match) {
      const sheetName = (match[1] ?? match[2]).replace(/''/g, "'");
      const sheet = Array.from(wb!.getElementsByTagNameNS(S, "sheet")).find(e => e.getAttribute("name") === sheetName);
      const rid = sheet?.getAttributeNS(NS.r, "id");
      const target = Array.from(wbRels!.documentElement.children).find(e => e.getAttribute("Id") === rid)?.getAttribute("Target");
      if (!target)
        throw new Error("Chart workbook sheet is missing");
      const sheetPath = resolvePath("xl/", target);
      let sx = sheets.get(sheetPath);
      if (!sx) {
        sx = parseXml(await zip.file(sheetPath)!.async("text"))!;
        sheets.set(sheetPath, sx);
      }
      const sheetData = sx.getElementsByTagNameNS(S, "sheetData")[0];
      const horizontal = !!match[5] && match[3] !== match[5];
      const startCol = column(match[3]);
      const startRow = Number(match[4]);
      values.forEach((v, i) => {
        const r = startRow + (horizontal ? 0 : i);
        const c = col(startCol + (horizontal ? i : 0)) + r;
        let row = Array.from(sheetData.children).find(e => Number(e.getAttribute("r")) === r);
        if (!row) {
          row = sx.createElementNS(S, "row");
          row.setAttribute("r", String(r));
          sheetData.insertBefore(row, Array.from(sheetData.children).find(e => Number(e.getAttribute("r")) > r) ?? null);
        }
        const old = Array.from(row.children).find(e => e.getAttribute("r") === c);
        const cell = sx.importNode(parseXml(`<root xmlns="${S}">${cellXml(c, numeric ? Number(v) : v)}</root>`)!.documentElement.firstElementChild!, true);
        if (old?.hasAttribute("s"))
          cell.setAttribute("s", old.getAttribute("s")!);
        if (old)
          row.replaceChild(cell, old);
        else
          row.insertBefore(cell, Array.from(row.children).find(e => column(e.getAttribute("r")!.replace(/\d+/g, "")) > column(c.replace(/\d+/g, ""))) ?? null);
      });
      const endCol = col(startCol + (horizontal ? values.length - 1 : 0));
      const endRow = startRow + (horizontal ? 0 : values.length - 1);
      ref.textContent = `${formula.slice(0, formula.lastIndexOf("!") + 1)}$${match[3]}$${startRow}${values.length > 1 ? `:$${endCol}$${endRow}` : ""}`;
      sx.getElementsByTagNameNS(S, "dimension")[0]?.remove();
    }
    const container = first(source, "numRef") ?? first(source, "strRef") ?? first(source, "numLit") ?? first(source, "strLit");
    if (!container) {
      const v = first(source, "v");
      if (v)
        v.textContent = String(values[0]);
      return;
    }
    const oldCache = child(container, numeric ? "numCache" : "strCache");
    const replacement = parseXml(`<root xmlns:c="${CHART_NS}">${cache(values, numeric)}</root>`)!.documentElement.firstElementChild!;
    const format = oldCache && child(oldCache, "formatCode");
    if (format && numeric) child(replacement, "formatCode")!.textContent = format.textContent;
    if (container.localName.endsWith("Lit"))
      container.replaceChildren(...Array.from(xml.importNode(replacement, true).children));
    else if (oldCache)
      container.replaceChild(xml.importNode(replacement, true), oldCache);
    else
      container.append(xml.importNode(replacement, true));
  };
  for (let i = 0; i < series.length; i++) {
    await setValues(child(series[i], "tx"), [data.series[i].name]);
    await setValues(child(series[i], "cat") ?? child(series[i], "xVal"), data.categories);
    await setValues(child(series[i], "val") ?? child(series[i], "yVal"), data.series[i].values);
  }
  if (zip)
    for (const [path, sheet] of sheets)
      zip.file(path, new XMLSerializer().serializeToString(sheet));
  return { xml, workbookPath, workbook: zip ? await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }) : null };
}
