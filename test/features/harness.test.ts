import { expect, it } from "vitest";
import { AttributeError, NotImplementedError, ValueError } from "../../src/errors.js";
import { Inches, Presentation } from "../../src/index.js";
import { $append, $getattr, $m } from "../pyshim.js";
import { parseFeature, runScenario, step } from "./runner.js";
import "./steps/text.js";
import "./steps/shape.js";

let input: Record<string, unknown>;
step("given", "a review context", [], (context) => Object.assign(context, input));

function check(text: string, values: Record<string, unknown>): void {
  input = values;
  const feature = parseFeature(`Feature: exception checks\nScenario: check\nGiven a review context\nThen ${text}`);
  runScenario(feature, feature.scenarios[0]);
}

it.each([
  ["accessing shape.click_action raises TypeError", "clickAction", TypeError],
  ["shape.shadow raises NotImplementedError", "shadow", NotImplementedError],
] as const)("%s accepts only the upstream exception type", (text, property, expected) => {
  const throwing = (error: Error) => ({ shape: { get [property]() { throw error; } } });
  expect(() => check(text, throwing(new expected()))).not.toThrow();
  expect(() => check(text, throwing(new ValueError("wrong exception")))).toThrow();
  expect(() => check(text, { shape: { [property]: null } })).toThrow(/did not raise/);
});

it.each(["before", "after"])("space_%s checks null and EMU spacing without swallowing the sentinel", (direction) => {
  const paragraph = (spacing: unknown) => ({ paragraph: { [`space${direction === "before" ? "Before" : "After"}`]: spacing } });
  expect(() => check(`paragraph.space_${direction}.pt raises AttributeError`, paragraph(null))).not.toThrow();
  expect(() => check(`paragraph.space_${direction}.pt raises AttributeError`, paragraph(76200))).toThrow(/did not raise/);
  expect(() => check(`paragraph.space_${direction}.pt == 6.0`, paragraph(76200))).not.toThrow();
  expect(() => check(`paragraph.space_${direction}.pt == 6.0`, paragraph(38100))).toThrow();
  const wrong = { get pt() { throw new TypeError("wrong exception"); } };
  expect(() => check(`paragraph.space_${direction}.pt raises AttributeError`, paragraph(wrong))).toThrow(/wrong exception/);
});

it("line spacing distinguishes upstream null, line multiples, and fixed EMU values", () => {
  for (const lineSpacing of [null, 1.5, 2.0]) {
    expect(() => check("paragraph.line_spacing.pt raises AttributeError", { paragraph: { lineSpacing } })).not.toThrow();
  }
  for (const [lineSpacing, points] of [[254000, 20], [304800, 24]]) {
    expect(() => check(`paragraph.line_spacing.pt is ${points}.0`, { paragraph: { lineSpacing } })).not.toThrow();
    expect(() => check("paragraph.line_spacing.pt raises AttributeError", { paragraph: { lineSpacing } })).toThrow();
  }
  expect(() => check("paragraph.line_spacing.pt is 20.0", { paragraph: { lineSpacing: 1.5 } })).toThrow();
});

it("missing shim attributes raise AttributeError across getter and method dispatch", () => {
  for (const value of [null, {}]) {
    expect(() => $getattr(value, "missing")).toThrow(AttributeError);
    expect(() => $m(value, "missing")).toThrow(AttributeError);
    expect(() => $append(value, 1)).toThrow(AttributeError);
  }
  expect($getattr(null, "missing", "default")).toBe("default");
  expect($getattr(254000, "pt")).toBe(20);
  const object = { count: 2, increment() { return ++this.count; } };
  expect($getattr(object, "increment")()).toBe(3);
  expect($m(object, "increment")).toBe(4);
  const values: number[] = [];
  $append(values, 1);
  expect(values).toEqual([1]);
});

it("parses required outline syntax with tables and triple-quoted doc strings", () => {
  const feature = parseFeature(`Feature: required syntax
Background:
  Given setup
Scenario Outline: case <value>
  When change <value>
  And continue
  Then check
  But also check
    | name | value |
    | item | <value> |
    """
    body <value>
    """
Examples:
  | value |
  | first |
  | second |`);
  expect(feature.background[0].text).toBe("setup");
  expect(feature.scenarios).toHaveLength(2);
  expect(feature.scenarios[0].steps.map((s) => s.kind)).toEqual(["when", "when", "then", "then"]);
  expect(feature.scenarios[0].steps[3].table).toEqual([["name", "value"], ["item", "first"]]);
  expect(feature.scenarios[1].steps[3].docString).toBe("body second");
});

it.each(["Scenario Template: extra", "Scenarios: extra", "* extra", "```"])("does not accept alternate syntax %s", (line) => {
  expect(() => parseFeature(`Feature: minimal syntax\nScenario: check\nGiven setup\n${line}`)).toThrow(/unrecognized line/);
});

it.each(["shape", "textFrame"])("%s.text replaces contents and creates paragraphs and soft breaks", (target) => {
  const prs = Presentation();
  const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
  const shape = slide.shapes.addTextbox(Inches(1), Inches(1), Inches(1), Inches(1));
  shape.text = "old\ncontents\nremoved";
  if (target === "shape") shape.text = "a\nb\vc";
  else shape.textFrame.text = "a\nb\vc";
  expect(shape.text).toBe("a\nb\vc");
  expect(shape.textFrame.paragraphs.map((p) => p.text)).toEqual(["a", "b\vc"]);
});
