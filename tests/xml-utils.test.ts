import { describe, expect, it } from "vitest";
import { insertInOrder, parseElement } from "../src/core/utils/xml";

const names = (el: Element) => Array.from(el.children).map((c) => c.localName);

describe("xml utils", () => {
  it("parseElement returns the root element", () => {
    const el = parseElement(`<a x="1"><b/></a>`);
    expect(el.localName).toBe("a");
    expect(el.getAttribute("x")).toBe("1");
  });

  it("insertInOrder puts an element before the first later sibling", () => {
    const order = ["a", "b", "c", "d", "extLst"];
    const root = parseElement(`<r><a/><c/><extLst/></r>`);
    insertInOrder(root, root.ownerDocument.createElement("b"), order);
    expect(names(root)).toEqual(["a", "b", "c", "extLst"]);
    insertInOrder(root, root.ownerDocument.createElement("d"), order);
    expect(names(root)).toEqual(["a", "b", "c", "d", "extLst"]);
  });

  it("insertInOrder ranks siblings by a key", () => {
    const root = parseElement(`<r><solidFill/><ln/></r>`);
    const key = (e: Element) => (e.localName.endsWith("Fill") ? "fill" : e.localName);
    insertInOrder(root, root.ownerDocument.createElement("prstGeom"), ["xfrm", "prstGeom", "fill", "ln"], key);
    expect(names(root)).toEqual(["prstGeom", "solidFill", "ln"]);
  });
});
