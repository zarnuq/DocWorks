/** Transition and object-animation authoring controls. All writes go through the view's undo history. */
import { Transition, TransitionPatch, TRANSITIONS } from "../../domain/pptx/animation";
import { animationParts, AnimationPart, AnimationEntry, AnimationSettings, AnimationType, ANIMATION_TYPES } from "../../domain/pptx/animation-edit";

export interface MotionControl {
  transition: Transition | null;
  entries: AnimationEntry[];
  selectedId: string | null;
  selectedName: string;
  setTransition(patch: TransitionPatch): void;
  applyTransitionToAll(): void;
  addAnimation(settings: AnimationSettings): void;
  updateAnimation(entry: AnimationEntry, patch: Partial<AnimationSettings>): void;
  removeAnimation(entry: AnimationEntry): void;
  updatePart?(part: AnimationPart, patch: Partial<Pick<AnimationPart,"dur"|"delay"|"motionPath">>): void;
  moveAnimation(entry: AnimationEntry, direction: -1 | 1): void;
  preview(): void;
  back(): void;
}

const STARTS: [string, string][] = [["On click", "click"], ["With previous", "with"], ["After previous", "after"]];
const EFFECTS: [string, string][] = ANIMATION_TYPES.map(([name, value]) => [name, value]);
const DIRECTIONS: [string, string][] = [["To left", "l"], ["To right", "r"], ["Up", "u"], ["Down", "d"]];

function section(host: HTMLElement, name: string) {
  host.createDiv({ cls: "pptx-props-title", text: name });
  return host.createDiv({ cls: "pptx-props-section" });
}
function row(host: HTMLElement, name: string) {
  const label = host.createEl("label", { cls: "pptx-props-row" });
  label.createSpan({ text: name });
  return label;
}
function select(host: HTMLElement, name: string, options: [string, string][], value: string, onChange: (value: string) => void) {
  const field = row(host, name).createEl("select", { cls: "dropdown", attr: { "data-field": name } });
  for (const [text, v] of options) field.createEl("option", { text, value: v });
  field.value = value;
  field.addEventListener("change", () => onChange(field.value));
  return field;
}
function seconds(host: HTMLElement, name: string, ms: number, minimum: number, change: (ms: number) => void) {
  const field = row(host, name).createEl("input", { type: "number", value: String(ms / 1000), attr: { "data-field": name, min: String(minimum), step: "0.1" } });
  field.addEventListener("change", () => {
    const n = Number(field.value);
    if (field.value !== "" && Number.isFinite(n) && n >= minimum) change(Math.round(n * 1000));
  });
  return field;
}
function button(host: HTMLElement, text: string, fn: () => void, disabled = false) {
  const b = host.createEl("button", { text });
  b.disabled = disabled;
  b.addEventListener("click", fn);
  return b;
}

export function buildMotionPanel(host: HTMLElement, ctl: MotionControl): void {
  host.replaceChildren();
  host.createDiv({ cls: "pptx-props-heading", text: "Motion" });
  const actions = host.createDiv({ cls: "pptx-props-buttons" });
  button(actions, "Play preview", () => ctl.preview());
  button(actions, "Back to format", () => ctl.back());

  const t = ctl.transition;
  const transition = section(host, "Slide transition");
  const options = TRANSITIONS.map(([name, value]): [string, string] => [name, value]);
  if (t && !options.some(([, value]) => value === t.type)) options.push([`Current: ${t.type}`, t.type]);
  select(transition, "Transition", options, t?.type ?? "cut", (type) => ctl.setTransition(type === "cut" ? { type } : { type, dir: "l", dur: t?.dur || 500 }));
  if (t && t.type !== "cut") {
    seconds(transition, "Transition duration (s)", t.dur, 0.1, (dur) => ctl.setTransition({ dur }));
    if (["push", "wipe", "cover", "pull"].includes(t.type)) {
      const dirs = [...DIRECTIONS];
      if (t.dir && !dirs.some(([, v]) => v === t.dir)) dirs.push([`Current: ${t.dir}`, t.dir]);
      select(transition, "Direction", dirs, t.dir ?? "l", (dir) => ctl.setTransition({ dir }));
    }
  }
  button(transition, "Apply to all slides", () => ctl.applyTransitionToAll(), !!t && !TRANSITIONS.some(([, type]) => type === t.type));
  transition.createDiv({ cls: "pptx-props-hint", text: "Plays when entering this slide in present mode." });

  const add = section(host, "Object animations");
  add.createDiv({ cls: "pptx-props-hint", text: ctl.selectedId ? `Selected: ${ctl.selectedName || "Object"}` : "Select one object on the slide to add an animation." });
  const settings: AnimationSettings = { type: "fade-in", trigger: "click", dur: 500, delay: 0 };
  select(add, "New effect", EFFECTS, settings.type, (v) => { settings.type = v as AnimationType; });
  select(add, "New start", STARTS, settings.trigger, (v) => { settings.trigger = v as AnimationSettings["trigger"]; });
  const paragraph = row(add, "By paragraph").createEl("input", { type: "checkbox", attr: { "data-field": "By paragraph" } });
  paragraph.addEventListener("change", () => { settings.byParagraph = paragraph.checked; });
  const duration = seconds(add, "New duration (s)", settings.dur, 0.1, (dur) => { settings.dur = dur; });
  const delay = seconds(add, "New delay (s)", settings.delay, 0, (ms) => { settings.delay = ms; });
  button(add, "Add animation", () => {
    if (duration.value !== "" && duration.checkValidity() && delay.value !== "" && delay.checkValidity()) ctl.addAnimation(settings);
  }, !ctl.selectedId);

  if (!ctl.entries.length) add.createDiv({ cls: "pptx-props-hint", text: "No object animations on this slide." });
  ctl.entries.forEach((entry, i) => {
    const s = section(host, `${i + 1}. ${entry.label} · ${entry.targetName}`);
    if (entry.editable) {
      if (entry.simple) select(s, "Effect", EFFECTS, entry.type, (v) => ctl.updateAnimation(entry, { type: v as AnimationType }));
      if(entry.simple){const checkbox=row(s,"By paragraph").createEl("input",{type:"checkbox"});checkbox.checked=!!entry.byParagraph;checkbox.addEventListener("change",()=>ctl.updateAnimation(entry,{byParagraph:checkbox.checked}));}
      select(s, "Start", STARTS, entry.trigger, (v) => ctl.updateAnimation(entry, { trigger: v as AnimationSettings["trigger"] }));
      if (entry.type !== "appear") seconds(s, "Duration (s)", entry.dur, 0.1, (dur) => ctl.updateAnimation(entry, { dur }));
      seconds(s, "Delay (s)", entry.delay, 0, (delay) => ctl.updateAnimation(entry, { delay }));
    } else s.createDiv({ cls: "pptx-props-hint", text: "This imported step contains effects the editor keeps intact." });
    if (!entry.simple && ctl.updatePart) {
      for (const part of animationParts(entry)) {
        const detail=s.createEl("details"); detail.createEl("summary",{text:part.label});
        seconds(detail,"Effect duration (s)",part.dur,.1,dur=>ctl.updatePart!(part,{dur}));
        seconds(detail,"Effect delay (s)",part.delay,0,delay=>ctl.updatePart!(part,{delay}));
        if(part.motionPath!==undefined){const field=row(detail,"Motion path").createEl("textarea",{text:part.motionPath});field.addEventListener("change",()=>ctl.updatePart!(part,{motionPath:field.value}));}
      }
    }
    const buttons = s.createDiv({ cls: "pptx-props-buttons" });
    button(buttons, "Move earlier", () => ctl.moveAnimation(entry, -1), !entry.editable || !ctl.entries[i - 1]?.editable);
    button(buttons, "Move later", () => ctl.moveAnimation(entry, 1), !entry.editable || !ctl.entries[i + 1]?.editable);
    button(buttons, "Remove", () => ctl.removeAnimation(entry), !entry.editable);
  });
}
