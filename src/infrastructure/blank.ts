import JSZip from "jszip";
import { NS } from "../core/utils/xml";
import { PptxDocument } from "./pptx/pptx-document";

/** Minimal Office-style packages for "New presentation" / "New document". */

const CT = "application/vnd.openxmlformats-officedocument.";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const NSP = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

const types = (overrides: [string, string][]) =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  overrides.map(([part, ct]) => `<Override PartName="/${part}" ContentType="${CT}${ct}"/>`).join("") +
  `</Types>`;

const rels = (list: [string, string][]) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  list.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${REL}${type}" Target="${target}"/>`).join("") +
  `</Relationships>`;

async function zip(files: Record<string, string>): Promise<Uint8Array> {
  const z = new JSZip();
  for (const [path, xml] of Object.entries(files)) z.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`);
  return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// ---------------------------------------------------------------------------
// pptx
// ---------------------------------------------------------------------------

const solid = (clr: string) => `<a:solidFill><a:schemeClr val="${clr}"/></a:solidFill>`;
const sys = (name: string, rgb: string) => `<a:${name}><a:sysClr val="${name === "dk1" ? "windowText" : "window"}" lastClr="${rgb}"/></a:${name}>`;
const srgb = (name: string, rgb: string) => `<a:${name}><a:srgbClr val="${rgb}"/></a:${name}>`;
const font = (latin: string) => `<a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/>`;
const ln = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid("phClr")}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

const THEME =
  `<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office">${sys("dk1", "000000")}${sys("lt1", "FFFFFF")}${srgb("dk2", "0E2841")}${srgb("lt2", "E8E8E8")}` +
  `${srgb("accent1", "156082")}${srgb("accent2", "E97132")}${srgb("accent3", "196B24")}${srgb("accent4", "0F9ED5")}` +
  `${srgb("accent5", "A02B93")}${srgb("accent6", "4EA72E")}${srgb("hlink", "467886")}${srgb("folHlink", "96607D")}</a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont>${font("Aptos Display")}</a:majorFont><a:minorFont>${font("Aptos")}</a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst>${solid("phClr").repeat(3)}</a:fillStyleLst>` +
  `<a:lnStyleLst>${ln(12700)}${ln(19050)}${ln(25400)}</a:lnStyleLst>` +
  `<a:effectStyleLst>${"<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3)}</a:effectStyleLst>` +
  `<a:bgFillStyleLst>${solid("phClr").repeat(3)}</a:bgFillStyleLst></a:fmtScheme>` +
  `</a:themeElements></a:theme>`;

const GRP = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`;

/** A placeholder shape; `ph` is the attribute string of p:ph, `lst` an optional list style. */
function ph(id: number, name: string, attrs: string, [x, y, w, h]: number[], anchor = "", lst = ""): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr${anchor ? ` anchor="${anchor}"` : ""}/><a:lstStyle>${lst}</a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

const TITLE_BOX = [838200, 365125, 10515600, 1325563];
const BODY_BOX = [838200, 1825625, 10515600, 4351338];

const BODY_STYLE = [2800, 2400, 2000, 1800, 1800]
  .map(
    (sz, i) =>
      `<a:lvl${i + 1}pPr marL="${228600 + i * 457200}" indent="-228600" algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc>` +
      `<a:spcBef><a:spcPts val="${i ? 500 : 1000}"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/>` +
      `<a:defRPr sz="${sz}" kern="1200">${solid("tx1")}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`
  )
  .join("");

const MASTER =
  `<p:sldMaster ${NSP}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}` +
  ph(2, "Title Placeholder 1", `type="title"`, TITLE_BOX, "ctr") +
  ph(3, "Text Placeholder 2", `type="body" idx="1"`, BODY_BOX) +
  `</p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/></p:sldLayoutIdLst>` +
  `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>` +
  `<a:defRPr sz="4400" kern="1200">${solid("tx1")}<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle>${BODY_STYLE}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const layout = (type: string, name: string, shapes: string) =>
  `<p:sldLayout ${NSP} type="${type}" preserve="1"><p:cSld name="${name}"><p:spTree>${GRP}${shapes}</p:spTree></p:cSld>` +
  `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const centered = (sz: number) => `<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="${sz}"/></a:lvl1pPr>`;

const LAYOUTS = [
  layout(
    "title",
    "Title Slide",
    ph(2, "Title 1", `type="ctrTitle"`, [1524000, 1122363, 9144000, 2387600], "b", centered(6000)) +
      ph(3, "Subtitle 2", `type="subTitle" idx="1"`, [1524000, 3602038, 9144000, 1655762], "", centered(2400))
  ),
  layout("obj", "Title and Content", ph(2, "Title 1", `type="title"`, TITLE_BOX) + ph(3, "Content Placeholder 2", `idx="1"`, BODY_BOX)),
  layout("blank", "Blank", ""),
];

/** A 16:9 deck with one title slide. */
export async function blankPptx(): Promise<Uint8Array> {
  const files: Record<string, string> = {
    "[Content_Types].xml": types([
      ["ppt/presentation.xml", "presentationml.presentation.main+xml"],
      ["ppt/presProps.xml", "presentationml.presProps+xml"],
      ["ppt/slideMasters/slideMaster1.xml", "presentationml.slideMaster+xml"],
      ...LAYOUTS.map((_, i): [string, string] => [`ppt/slideLayouts/slideLayout${i + 1}.xml`, "presentationml.slideLayout+xml"]),
      ["ppt/theme/theme1.xml", "theme+xml"],
    ]),
    "_rels/.rels": rels([["officeDocument", "ppt/presentation.xml"]]),
    "ppt/presentation.xml":
      `<p:presentation ${NSP} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst/><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": rels([
      ["slideMaster", "slideMasters/slideMaster1.xml"],
      ["theme", "theme/theme1.xml"],
      ["presProps", "presProps.xml"],
    ]),
    "ppt/presProps.xml": `<p:presentationPr ${NSP}/>`,
    "ppt/slideMasters/slideMaster1.xml": MASTER,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels([
      ...LAYOUTS.map((_, i): [string, string] => ["slideLayout", `../slideLayouts/slideLayout${i + 1}.xml`]),
      ["theme", "../theme/theme1.xml"],
    ]),
    "ppt/theme/theme1.xml": THEME,
  };
  LAYOUTS.forEach((xml, i) => {
    files[`ppt/slideLayouts/slideLayout${i + 1}.xml`] = xml;
    files[`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`] = rels([["slideMaster", "../slideMasters/slideMaster1.xml"]]);
  });
  const doc = await PptxDocument.load(await zip(files));
  doc.addSlide(-1, "ppt/slideLayouts/slideLayout1.xml");
  return doc.save();
}

// ---------------------------------------------------------------------------
// docx
// ---------------------------------------------------------------------------

/** A Letter-size document with one empty paragraph. */
export function blankDocx(): Promise<Uint8Array> {
  const w = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;
  return zip({
    "[Content_Types].xml": types([
      ["word/document.xml", "wordprocessingml.document.main+xml"],
      ["word/styles.xml", "wordprocessingml.styles+xml"],
    ]),
    "_rels/.rels": rels([["officeDocument", "word/document.xml"]]),
    "word/_rels/document.xml.rels": rels([["styles", "styles.xml"]]),
    "word/document.xml":
      `<w:document ${w}><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    "word/styles.xml":
      `<w:styles ${w}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Aptos" w:cs="Aptos"/>` +
      `<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>`,
  });
}

// ---------------------------------------------------------------------------
// xlsx
// ---------------------------------------------------------------------------

/** A workbook with one empty sheet, in Google Sheets' default font. */
export function blankXlsx(): Promise<Uint8Array> {
  const s = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS.r}"`;
  return zip({
    "[Content_Types].xml": types([
      ["xl/workbook.xml", "spreadsheetml.sheet.main+xml"],
      ["xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"],
      ["xl/styles.xml", "spreadsheetml.styles+xml"],
    ]),
    "_rels/.rels": rels([["officeDocument", "xl/workbook.xml"]]),
    "xl/_rels/workbook.xml.rels": rels([
      ["worksheet", "worksheets/sheet1.xml"],
      ["styles", "styles.xml"],
    ]),
    "xl/workbook.xml": `<workbook ${s}><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`,
    "xl/worksheets/sheet1.xml": `<worksheet ${s}><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`,
    "xl/styles.xml":
      `<styleSheet ${s}><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  });
}
