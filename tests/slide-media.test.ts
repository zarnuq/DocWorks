import { it, expect } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addMedia, readMedia, setMediaPlayback } from "../src/domain/pptx/media";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { renderSlide } from "../src/presentation/pptx/slide-renderer";
it.each(["audio", "video"] as const)("embeds %s, edits playback settings and round trips", async (kind) => {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  const frame = addMedia(doc, path, new Uint8Array([1, 2, 3]), kind, kind === "audio" ? "mp3" : "mp4", "Demo");
  expect(readMedia(doc, path, frame)).toMatchObject({ kind, autoPlay: false, loop: false });
  setMediaPlayback(doc.xml(path)!, frame, { autoPlay: true, loop: true });
  const model = parseSlide(doc, path);
  expect(model.shapes.find(s => s.el === frame)?.media).toMatchObject({ autoPlay: true, loop: true });
  const host = document.createElement("div");
  renderSlide(host, model, doc.slideSize, { width: 800, editable: false, playMedia: true, imageUrl: () => "blob:demo" });
  expect(host.querySelector(kind)?.getAttribute("src")).toBe("blob:demo");
  expect(host.querySelector(kind)?.hasAttribute("controls")).toBe(true);
  const back = await PptxDocument.load(await doc.save());
  expect(parseSlide(back, path).shapes.find(s => s.media)?.media).toMatchObject({ kind, autoPlay: true, loop: true });
});

import { linkedMediaCandidates } from "../src/domain/pptx/media";

it("finds linked (not embedded) media only as vault files next to the deck, never on the web or elsewhere on disk", () => {
  const deck = "Talks/2026/deck.pptx";
  expect(linkedMediaCandidates("clip.mp4", deck)).toEqual(["Talks/2026/clip.mp4"]);
  expect(linkedMediaCandidates("media/clip%20one.mp4", deck)).toEqual(["Talks/2026/media/clip one.mp4", "Talks/2026/clip one.mp4"]);
  expect(linkedMediaCandidates("../shared/song.mp3", deck)).toEqual(["Talks/shared/song.mp3", "Talks/2026/song.mp3"]);
  // Absolute paths from the author's machine: only a same-named file beside the deck (PowerPoint's own fallback).
  expect(linkedMediaCandidates("file:///C:/Users/Ann/Videos/demo.mp4", deck)).toEqual(["Talks/2026/demo.mp4"]);
  expect(linkedMediaCandidates("C:\\Users\\Ann\\demo.webm", deck)).toEqual(["Talks/2026/demo.webm"]);
  expect(linkedMediaCandidates("/home/ann/demo.wav", deck)).toEqual(["Talks/2026/demo.wav"]);
  // Leaving the vault is never followed.
  expect(linkedMediaCandidates("../../../../etc/clip.mp4", deck)).toEqual(["Talks/2026/clip.mp4"]);
  // Web links and non-media files are refused.
  expect(linkedMediaCandidates("https://www.youtube.com/embed/abc", deck)).toEqual([]);
  expect(linkedMediaCandidates("http://example.com/a.mp4", deck)).toEqual([]);
  expect(linkedMediaCandidates("notes.txt", deck)).toEqual([]);
  expect(linkedMediaCandidates("clip.mp4", "deck.pptx")).toEqual(["clip.mp4"]);
});

it("plays linked media when the view can resolve it, and shows a label otherwise", async () => {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  addMedia(doc, path, new Uint8Array([1]), "video", "mp4", "Linked");
  const model = parseSlide(doc, path);
  const shape = model.shapes.find((s) => s.media)!;
  shape.media = { ...shape.media!, external: true, path: "clip.mp4" };
  const host = document.createElement("div");
  renderSlide(host, model, doc.slideSize, { width: 800, editable: false, playMedia: true, imageUrl: (p) => (p === "clip.mp4" ? "app://vault/clip.mp4" : null) });
  expect(host.querySelector("video")?.getAttribute("src")).toBe("app://vault/clip.mp4");
  const none = document.createElement("div");
  renderSlide(none, model, doc.slideSize, { width: 800, editable: false, playMedia: true, imageUrl: () => null });
  expect(none.querySelector("video")).toBeNull();
  expect(none.querySelector(".pptx-media-label")).not.toBeNull();
});
