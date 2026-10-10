import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { EditorView, NodeView } from "prosemirror-view";
import { Origins, ReadContext } from "../../domain/docx/docx-reader";
import { WriteContext } from "../../domain/docx/docx-writer";
import { readTextBox, writeTextBox } from "../../domain/docx/textbox";
import { docxPlugins, EditorContext } from "./docx-editor";

const editors = new WeakMap<Element, EditorView>();

/**
 * A Word text box, editable in place: a small editor of its own inside the box. Each edit writes the box's content
 * back into the run XML the document keeps for it. Other run content we don't edit is drawn as usual.
 */
export class TextBoxView implements NodeView {
  dom: HTMLElement;
  private sub: EditorView | null = null;
  private xml: string;
  private origins: Origins = new Map();

  /** The editor inside a text box element (for tests). */
  static editorOf(el: Element): EditorView | null {
    return editors.get(el) ?? null;
  }

  constructor(private node: PMNode, outer: EditorView, getPos: () => number | undefined, ctx: EditorContext, read: ReadContext, write: WriteContext) {
    this.xml = node.attrs.xml as string;
    this.dom = DOMSerializer.renderSpec(document, node.type.spec.toDOM!(node)).dom;
    const content = readTextBox(this.xml, read);
    if (!content) return;
    this.origins = content.origins;
    this.dom.replaceChildren();
    const sub: EditorView = new EditorView(this.dom, {
      state: EditorState.create({ doc: content.doc, plugins: docxPlugins(ctx) }),
      dispatchTransaction: (tr) => {
        sub.updateState(sub.state.apply(tr));
        const pos = getPos();
        if (!tr.docChanged || pos === undefined) return;
        this.xml = writeTextBox(this.xml, sub.state.doc, this.origins, write);
        outer.dispatch(outer.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, xml: this.xml }));
      },
    });
    this.sub = sub;
    editors.set(this.dom, sub);
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    // Someone else changed the box (undo in the document): draw it afresh.
    if (node.attrs.xml !== this.xml) return false;
    this.node = node;
    return true;
  }

  stopEvent(e: Event): boolean {
    return !!this.sub && this.sub.dom.contains(e.target as Node);
  }

  ignoreMutation(): boolean {
    return true;
  }

  destroy(): void {
    this.sub?.destroy();
  }
}
