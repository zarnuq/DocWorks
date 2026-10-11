import { describe, expect, it, vi } from "vitest";
import { buildMotionPanel, MotionControl } from "../src/presentation/pptx/motion-panel";

function panel(patch: Partial<MotionControl> = {}) {
  const ctl: MotionControl = { transition: null, entries: [], selectedId: "3", selectedName: "Title", setTransition: vi.fn(), applyTransitionToAll: vi.fn(),
    addAnimation: vi.fn(), updateAnimation: vi.fn(), removeAnimation: vi.fn(), moveAnimation: vi.fn(), preview: vi.fn(), back: vi.fn(), ...patch };
  const host = document.createElement("div");
  buildMotionPanel(host, ctl);
  const change = (name: string, value: string) => {
    const field = host.querySelector<HTMLInputElement>(`[data-field="${name}"]`)!;
    field.value = value;
    field.dispatchEvent(new Event("change"));
  };
  const click = (text: string) => Array.from(host.querySelectorAll("button")).find((b) => b.textContent === text)!.click();
  return { host, ctl, change, click };
}

describe("motion panel", () => {
  it("adds an animation with the effect, trigger, duration and delay chosen for the selection", () => {
    const { ctl, change, click } = panel();
    change("New effect", "fly-left");
    change("New start", "after");
    change("New duration (s)", "1.2");
    change("New delay (s)", "0.3");
    click("Add animation");
    expect(ctl.addAnimation).toHaveBeenCalledWith({ type: "fly-left", trigger: "after", dur: 1200, delay: 300 });
  });

  it("requires a selected object and doesn't edit XML just by opening the panel", () => {
    const { ctl, host } = panel({ selectedId: null });
    expect(Array.from(host.querySelectorAll("button")).find((b) => b.textContent === "Add animation")?.disabled).toBe(true);
    expect(ctl.addAnimation).not.toHaveBeenCalled();
    expect(ctl.setTransition).not.toHaveBeenCalled();
  });

  it("sets and removes a transition and provides apply-all and present preview", () => {
    const { ctl, change, click } = panel();
    change("Transition", "push");
    expect(ctl.setTransition).toHaveBeenCalledWith({ type: "push", dir: "l", dur: 500 });
    click("Apply to all slides");
    click("Play preview");
    expect(ctl.applyTransitionToAll).toHaveBeenCalledOnce();
    expect(ctl.preview).toHaveBeenCalledOnce();
    change("Transition", "cut");
    expect(ctl.setTransition).toHaveBeenLastCalledWith({ type: "cut" });
  });

  it("preserves imported transitions when editing their duration", () => {
    const { ctl, host, change } = panel({ transition: { type: "vortex", dir: "r", dur: 2000, advanceAfter: 3000 } });
    expect(host.querySelector<HTMLSelectElement>('[data-field="Transition"]')!.value).toBe("vortex");
    change("Transition duration (s)", "1.4");
    expect(ctl.setTransition).toHaveBeenCalledWith({ dur: 1400 });
  });

  it("ignores empty or invalid timing values", () => {
    const { ctl, change } = panel({ transition: { type: "fade", dir: null, dur: 500, advanceAfter: null } });
    change("Transition duration (s)", "");
    change("Transition duration (s)", "-2");
    expect(ctl.setTransition).not.toHaveBeenCalled();
  });
});
