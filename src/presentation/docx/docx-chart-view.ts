import type { Node as PMNode } from "prosemirror-model";
import type { EditorView, NodeView } from "prosemirror-view";
import type { ChartModel } from "../../domain/pptx/chart";
import { renderChart } from "../pptx/chart-renderer";

const EMU_PER_PX = 9525;

/** A native Word chart, drawn from its chart part; double-click opens the chart editor. */
export class ChartView implements NodeView {
  dom: HTMLElement;
  constructor(private node: PMNode, _view: EditorView, getPos: () => number | undefined, private find: (xml: string) => { model: ChartModel; size: { cx: number; cy: number } } | null, edit: (pos: number) => void) {
    this.dom = createSpan({ cls: "docx-chart", attr: { contenteditable: "false" } });
    this.draw();
    this.dom.addEventListener("dblclick", (e) => {
      e.preventDefault();
      const pos = getPos();
      if (pos !== undefined) edit(pos);
    });
  }

  private draw(): void {
    const found = this.find(this.node.attrs.xml as string);
    this.dom.empty();
    if (!found) return;
    const w = found.size.cx / EMU_PER_PX, h = found.size.cy / EMU_PER_PX;
    this.dom.setCssStyles({ width: `${w}px`, height: `${h}px` });
    // The renderer works in points.
    renderChart(this.dom, found.model, { w: w * 0.75, h: h * 0.75 });
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    if (node.attrs.xml !== this.node.attrs.xml) {
      this.node = node;
      this.draw();
    }
    return true;
  }

  ignoreMutation(): boolean { return true; }
}
