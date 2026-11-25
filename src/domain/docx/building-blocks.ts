/**
 * Google Docs' Insert › Building blocks: ready-made meeting notes, an email draft and tracker tables, built from
 * ordinary paragraphs, tables and smart chips (so they save as plain Word content).
 */
import { Node as PMNode } from "prosemirror-model";
import { docxSchema as schema } from "./schema";
import { newChip } from "./chips";

export type BuildingBlock = "meeting" | "email" | "roadmap" | "review" | "assets" | "launch";

export const BUILDING_BLOCKS: [BuildingBlock, string][] = [
  ["meeting", "Meeting notes"],
  ["email", "Email draft"],
  ["roadmap", "Product roadmap"],
  ["review", "Review tracker"],
  ["assets", "Project assets"],
  ["launch", "Content launch tracker"],
];

const PROJECT_STATUS = ["Not started", "In progress", "Blocked", "Done"];
const REVIEW_STATUS = ["Not started", "In review", "Changes requested", "Approved"];
/** The note export's task box font. */
const BOX_FONT = "Segoe UI Symbol";

/** What a tracker's first row holds in a column: nothing, a status dropdown or a date chip. */
type Cell = "" | "status" | "review" | "date";

/** Columns of each tracker, and what its first row holds (text is left blank; chips are placed). */
const TRACKERS: Record<Exclude<BuildingBlock, "meeting" | "email">, [header: string, cell: Cell][]> = {
  roadmap: [["Feature", ""], ["Assignee", ""], ["Status", "status"], ["Notes", ""]],
  review: [["Reviewer", ""], ["Status", "review"], ["Notes", ""]],
  assets: [["Asset", ""], ["Owner", ""], ["Status", "status"], ["Notes", ""]],
  launch: [["Content", ""], ["Channel", ""], ["Owner", ""], ["Status", "status"], ["Launch date", "date"]],
};

const para = (content: PMNode[] = [], attrs: Record<string, unknown> = {}) => schema.nodes.paragraph.create(attrs, content);
const text = (s: string, bold = false) => schema.text(s, bold ? [schema.marks.bold.create()] : []);
const cell = (content: PMNode[]) => schema.nodes.table_cell.create(null, para(content));
const table = (rows: PMNode[][][]) => schema.nodes.table.create({ borders: "all" }, rows.map((r) => schema.nodes.table_row.create(null, r.map(cell))));

function chipFor(c: Cell, today: string): PMNode[] {
  if (c === "status" || c === "review") return [schema.nodes.chip.create(newChip("dropdown", { options: c === "status" ? PROJECT_STATUS : REVIEW_STATUS }))];
  if (c === "date") return [schema.nodes.chip.create(newChip("date", { date: today }))];
  return [];
}

/**
 * The blocks of `kind`. `taskList` is the numId of the document's checklist definition (for action items);
 * `today` an ISO date for date chips. A block ending in a table is followed by an empty paragraph to type in.
 */
export function buildingBlock(kind: BuildingBlock, opts: { taskList: string; bulletList: string; today: string }): PMNode[] {
  if (kind === "meeting") {
    const date = schema.nodes.chip.create(newChip("date", { date: opts.today }));
    return [
      para([date, text(" | Meeting title")], { style: "Heading1" }),
      para([text("Attendees: ", true)]),
      para([text("Notes")], { style: "Heading2" }),
      para([], { numId: opts.bulletList, ilvl: 0 }),
      para([text("Action items")], { style: "Heading2" }),
      para([schema.text("☐", [schema.marks.font.create({ font: BOX_FONT })]), text(" ")], { numId: opts.taskList, ilvl: 0 }),
    ];
  }
  if (kind === "email") {
    return [table(["To", "Cc", "Bcc", "Subject"].map((label) => [[text(label, true)], []])), para()];
  }
  const cols = TRACKERS[kind];
  return [table([cols.map(([h]) => [text(h, true)]), cols.map(([, c]) => chipFor(c, opts.today))]), para()];
}
