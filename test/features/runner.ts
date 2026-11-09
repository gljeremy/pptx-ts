// A small Gherkin runner for vitest, enough for python-pptx's behave features: Feature,
// Background, Scenario, Scenario Outline with Examples, Given/When/Then/And/But steps, and
// step tables and doc strings (as behave's `context.table` and `context.text`). Steps match
// like behave's default "parse" matcher: `{name}` captures `.+?`, the whole step text must
// match, case-insensitively, and the first registered step that matches wins.
import { describe, it } from "vitest";

type Kind = "given" | "when" | "then" | "step";
type StepDef = { kind: Kind; pattern: string; re: RegExp; fn: (context: any, ...args: string[]) => unknown };
const registry: StepDef[] = [];

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Register a step, as behave's `@given/@when/@then("pattern")`. behave passes the pattern's
 * named fields as keyword arguments, so `params` (the python function's parameter names after
 * `context`) say which field goes to which parameter.
 */
export function step(kind: Kind, pattern: string, params: string[], fn: StepDef["fn"]): void {
  const parts = pattern.split(/(\{\w*\})/);
  const re = new RegExp("^" + parts.map((p, i) => (i % 2 ? "(.+?)" : escapeRe(p))).join("") + "$", "is");
  const fields = parts.filter((_p, i) => i % 2).map((p) => p.slice(1, -1));
  const order = fields.every((f) => f !== "") ? params.map((p) => fields.indexOf(p)) : params.map((_p, i) => i);
  if (order.includes(-1) || params.length !== fields.length) throw new Error(`step ${JSON.stringify(pattern)}: parameters ${params} don't match its fields`);
  const bound = (context: any, ...args: string[]) => fn(context, ...order.map((j) => args[j]));
  Object.defineProperty(bound, "name", { value: fn.name });
  registry.push({ kind, pattern, re, fn: bound });
}

export interface Step {
  keyword: string;
  kind: Kind;
  text: string;
  line: number;
  table?: string[][];
  docString?: string;
}
export interface Scenario {
  name: string;
  line: number;
  steps: Step[];
}
export interface Feature {
  name: string;
  background: Step[];
  scenarios: Scenario[];
}

const cells = (line: string) =>
  line
    .trim()
    .slice(1, -1)
    .split(/(?<!\\)\|/)
    .map((c) => c.trim().replace(/\\\|/g, "|"));

/** Parse the text of a .feature file; Scenario Outlines are expanded, one scenario per example row. */
export function parseFeature(text: string): Feature {
  const lines = text.split(/\r?\n/);
  const feature: Feature = { name: "", background: [], scenarios: [] };
  let steps: Step[] | null = null;
  let outline: { name: string; line: number; steps: Step[] } | null = null;
  let examplesHeader: string[] | null = null;
  let lastKind: Kind = "given";
  let lastStep: Step | null = null;
  const endOutline = () => {
    outline = null;
    examplesHeader = null;
  };
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const line = raw.trim();
    if (line === "" || line.startsWith("#") || line.startsWith("@")) continue;
    let m: RegExpExecArray | null;
    if ((m = /^Feature:\s*(.*)$/.exec(line))) {
      feature.name = m[1];
      continue;
    }
    if (/^Background:/.test(line)) {
      steps = feature.background;
      endOutline();
      continue;
    }
    if ((m = /^Scenario Outline:\s*(.*)$/.exec(line))) {
      outline = { name: m[1], line: i + 1, steps: [] };
      examplesHeader = null;
      steps = outline.steps;
      continue;
    }
    if ((m = /^Scenario:\s*(.*)$/.exec(line))) {
      endOutline();
      const sc: Scenario = { name: m[1], line: i + 1, steps: [] };
      feature.scenarios.push(sc);
      steps = sc.steps;
      continue;
    }
    if (/^Examples:/.test(line)) {
      examplesHeader = null;
      steps = null;
      continue;
    }
    if ((m = /^(Given|When|Then|And|But)\s+(.*)$/.exec(line))) {
      const kw = m[1];
      const kind: Kind = kw === "Given" ? "given" : kw === "When" ? "when" : kw === "Then" ? "then" : lastKind;
      lastKind = kind;
      lastStep = { keyword: kw, kind, text: m[2], line: i + 1 };
      if (!steps) throw new Error(`step outside a scenario at line ${i + 1}`);
      steps.push(lastStep);
      continue;
    }
    if (line.startsWith('"""')) {
      const indent = raw.indexOf('"""');
      const body: string[] = [];
      for (i++; i < lines.length && lines[i].trim() !== '"""'; i++) body.push(lines[i].slice(indent));
      lastStep!.docString = body.join("\n");
      continue;
    }
    if (line.startsWith("|")) {
      const row = cells(line);
      if (outline && steps === null) {
        // an Examples table row
        if (!examplesHeader) {
          examplesHeader = row;
          continue;
        }
        const o = outline as { name: string; line: number; steps: Step[] };
        const sub = (s: string) => s.replace(/<([^>]+)>/g, (all, k) => (examplesHeader!.includes(k) ? row[examplesHeader!.indexOf(k)] : all));
        const n = feature.scenarios.filter((s) => s.line === o.line).length + 1;
        feature.scenarios.push({
          name: `${sub(o.name)} -- @${n} ${row.join(" | ")}`,
          line: o.line,
          steps: o.steps.map((s) => ({
            ...s,
            text: sub(s.text),
            table: s.table?.map((r) => r.map(sub)),
            docString: s.docString === undefined ? undefined : sub(s.docString),
          })),
        });
        continue;
      }
      lastStep!.table = [...(lastStep!.table ?? []), row];
      continue;
    }
    // free-text description of the feature
    if (steps === null && !outline && feature.scenarios.length === 0 && feature.background.length === 0) continue;
    throw new Error(`unrecognized line ${i + 1}: ${line}`);
  }
  return feature;
}

/** behave's `context.table`: rows of the step table after its heading row. */
function table(rows: string[][]): any {
  const [headings, ...body] = rows;
  const tbl = body.map((r) => Object.assign(Object.create({ get: (k: string) => r[headings.indexOf(k)] }), Object.fromEntries(headings.map((h, j) => [h, r[j]])), { cells: r }));
  return Object.assign(tbl, { headings, rows: tbl });
}

function findStep(s: Step): { def: StepDef; args: string[] } {
  for (const def of registry) {
    if (def.kind !== s.kind && def.kind !== "step") continue;
    const m = def.re.exec(s.text);
    if (m) return { def, args: m.slice(1) };
  }
  throw new Error(`undefined step: ${s.keyword} ${s.text}`);
}

/** Run one scenario's steps (background first) with a fresh context. */
export function runScenario(feature: Feature, scenario: Scenario): void {
  const context: any = {
    /** behave's `context.execute_steps()`: run steps given as Gherkin text, keywords in any case */
    executeSteps(text: string) {
      const gherkin = String(text).replace(/^(\s*)(given|when|then|and|but)\b/gim, (_m, sp, kw) => sp + kw[0].toUpperCase() + kw.slice(1).toLowerCase());
      const steps = parseFeature(`Scenario: execute_steps\n${gherkin}`).scenarios[0].steps;
      runSteps(context, steps);
    },
  };
  runSteps(context, [...feature.background, ...scenario.steps]);
}

function runSteps(context: any, steps: Step[]): void {
  for (const s of steps) {
    const { def, args } = findStep(s);
    context.table = s.table ? table(s.table) : null;
    context.text = s.docString ?? null;
    try {
      def.fn(context, ...args);
    } catch (e: any) {
      if (e && typeof e === "object") e.message = `${s.keyword} ${s.text} (line ${s.line}, ${def.fn.name})\n${e.message}`;
      throw e;
    }
  }
}

/** Define a vitest suite for a feature: one test per scenario; `skips` maps scenario names to reasons. */
export function defineFeature(file: string, text: string, skips: Record<string, string> = {}): void {
  const feature = parseFeature(text);
  describe(`${file}: ${feature.name}`, () => {
    for (const sc of feature.scenarios) {
      const reason = skips[sc.name];
      (reason ? it.skip : it)(sc.name, () => runScenario(feature, sc));
    }
  });
}
