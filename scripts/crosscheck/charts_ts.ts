// Compare pptx-ts's chart API with python-pptx's on every chart in the upstream test decks.
// Run scripts/crosscheck/charts_py.py first; both sides read the paths in chart_props.json.
// Usage: npx tsx scripts/crosscheck/charts_ts.ts charts_py.json <python-pptx checkout>
import { readFileSync } from "node:fs";
import { Presentation } from "../../src/index.js";
import { BaseEnum } from "../../src/enum/base.js";

const PROPS = JSON.parse(readFileSync(new URL("./chart_props.json", import.meta.url), "utf8"));
const py = JSON.parse(readFileSync(process.argv[2] ?? "charts_py.json", "utf8"));
const root = (process.argv[3] ?? "../python-pptx").replace(/\\/g, "/").replace(/\/?$/, "/");

const camel = (s: string) => s.replace(/_+([a-z0-9])/g, (_m, c) => c.toUpperCase());

function norm(v: any): any {
  if (v instanceof BaseEnum) return v.name;
  if (v instanceof String) return v.valueOf();
  if (v === undefined) return "N/A";
  if (v === null || typeof v === "boolean" || typeof v === "number" || typeof v === "string") return v;
  if (Array.isArray(v)) return v.map(norm);
  return `<${v?.constructor?.name}>`;
}

class Missing extends Error {}

function readRaw(obj: any, path: string): any {
  for (const part of path.split(".")) {
    // python `None.x` is an AttributeError, as JS's TypeError for `null.x` is below
    if (obj === null || obj === undefined) throw new TypeError("Cannot read properties of null");
    const k = camel(part.endsWith("()") ? part.slice(0, -2) : part);
    if (!(k in obj)) throw new Missing(k);
    obj = part.endsWith("()") ? [...obj[k]()] : obj[k];
  }
  return obj;
}

function read(obj: any, path: string): any {
  try {
    if (path.startsWith("len(") && path.endsWith(")")) return readRaw(obj, path.slice(4, -1)).length;
    return norm(readRaw(obj, path));
  } catch (e: any) {
    if (e instanceof Missing) return "N/A";
    // JS's TypeError for reading a property of null/undefined is python's AttributeError
    if (e instanceof TypeError && /of (null|undefined)/.test(e.message)) return "ERR:AttributeError";
    return `ERR:${e?.name ?? "Error"}`;
  }
}

const props = (obj: any, kind: string) => Object.fromEntries(PROPS[kind].map((p: string) => [p, read(obj, p)]));

function axisInfo(chart: any, name: string): any {
  let ax: any;
  try {
    ax = chart[camel(name)];
  } catch (e: any) {
    return `ERR:${e.name}`;
  }
  const d: any = props(ax, "axis");
  if (name === "value_axis") Object.assign(d, props(ax, "value_axis"));
  if (ax.hasTitle) d.title = props(ax.axisTitle, "axis_title");
  return d;
}

function chartInfo(chart: any): any {
  const d: any = props(chart, "chart");
  if (chart.hasTitle) d.title = props(chart.chartTitle, "chart_title");
  if (chart.hasLegend) d.legend = props(chart.legend, "legend");
  d.plots = [...chart.plots].map((plot: any) => {
    const p: any = props(plot, "plot");
    if (plot.hasDataLabels) p.data_labels = props(plot.dataLabels, "data_labels");
    p.series = [...plot.series].map((s: any) => {
      const sd: any = props(s, "series");
      if (["XySeries", "BubbleSeries"].includes(s.constructor.name)) Object.assign(sd, props(s, "xy_series"));
      try {
        const pts = s.points;
        sd.points = Array.from({ length: Math.min(pts.length, 3) }, (_, i) => props(pts.at(i), "point"));
      } catch (e: any) {
        sd.points = `ERR:${e.name}`;
      }
      return sd;
    });
    return p;
  });
  d.category_axis = axisInfo(chart, "category_axis");
  d.value_axis = axisInfo(chart, "value_axis");
  d.series_names = [...chart.series].map((s: any) => s.name);
  return d;
}

/** JSON equality with numbers compared as numbers (python 1.0 and JS 1 agree). */
const canon = (v: any) => JSON.stringify(v, (_k, x) => (typeof x === "number" ? Number(x.toPrecision(15)) : x));

let ok = 0;
let bad = 0;
for (const key of Object.keys(py)) {
  const charts: any[] = [];
  const prs = Presentation(new Uint8Array(readFileSync(root + key)));
  [...prs.slides].forEach((slide: any, si: number) => {
    for (const sh of slide.shapes as Iterable<any>) {
      if (!sh.hasChart) continue;
      try {
        charts.push({ slide: si, shape: sh.name, ...chartInfo(sh.chart) });
      } catch (e: any) {
        charts.push({ slide: si, shape: sh.name, error: `ERR:${e.name}: ${String(e.message).slice(0, 80)}` });
      }
    }
  });
  const want = py[key];
  for (let i = 0; i < Math.max(charts.length, want.length); i++) {
    const a = charts[i];
    const b = want[i];
    if (canon(a) === canon(b)) {
      ok++;
      continue;
    }
    bad++;
    console.log(`MISMATCH ${key} chart ${i} (${b?.shape})`);
    const diff = (x: any, y: any, path: string) => {
      if (canon(x) === canon(y)) return;
      if (x && y && typeof x === "object" && typeof y === "object") {
        for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) diff(x[k], y[k], `${path}.${k}`);
        return;
      }
      console.log(`   ${path}\n      ts: ${canon(x)?.slice(0, 200)}\n      py: ${canon(y)?.slice(0, 200)}`);
    };
    diff(a, b, "");
  }
}
console.log(`${ok}/${ok + bad} charts match`);
