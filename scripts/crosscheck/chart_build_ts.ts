// Build the charts in chart_build.json with pptx-ts and compare the deck with python-pptx's,
// part by part. Embedded workbooks are compared part by part too, except docProps/core.xml
// (XlsxWriter stamps the current time there).
// Usage: python scripts/crosscheck/chart_build_py.py charts_py.pptx
//        npx tsx scripts/crosscheck/chart_build_ts.ts charts_py.pptx [charts_ts.pptx]
// Expected difference: the `multi` data's 39.0 is a python float written "39.0"; JS has one
// number type and writes 39 as "39" (PARITY.md). Everything else is byte-identical.
import { readFileSync, writeFileSync } from "node:fs";
import { unzipSync } from "fflate";
import * as pptx from "../../src/index.js";

const { Presentation, BytesIO, Inches, Pt, RGBColor, XL_CHART_TYPE, CategoryChartData, XyChartData, BubbleChartData } = pptx;
const SPEC = JSON.parse(readFileSync(new URL("./chart_build.json", import.meta.url), "utf8"));
const camel = (s: string) => s.replace(/_+([a-z0-9])/g, (_m, c) => c.toUpperCase());

function label(v: any): any {
  if (v && typeof v === "object" && "date" in v) return new Date(`${v.date}T00:00:00Z`);
  if (v && typeof v === "object" && "datetime" in v) return new Date(`${v.datetime}Z`);
  return v;
}

function makeData(d: any): any {
  if (d.kind === "category") {
    const cd = new CategoryChartData(d.number_format ?? "General");
    if (d.categories.multi) {
      for (const [top, subs] of d.categories.multi) {
        const c = cd.addCategory(top);
        for (const s of subs) c.addSubCategory(s);
      }
    } else cd.categories = d.categories.map(label);
    for (const s of d.series) cd.addSeries(s[0], s[1], s[2] ?? null);
    return cd;
  }
  const cd = d.kind === "xy" ? new XyChartData(d.number_format ?? "General") : new BubbleChartData(d.number_format ?? "General");
  for (const s of d.series) {
    const ser = cd.addSeries(s[0], s[2] ?? null);
    for (const pt of s[1]) (ser.addDataPoint as any)(...pt);
  }
  return cd;
}

function value(v: any): any {
  if (typeof v === "string" && v.startsWith("RGB:")) return RGBColor.fromString(v.slice(4));
  if (typeof v === "string" && v.startsWith("PT:")) return Pt(Number(v.slice(3)));
  if (typeof v === "string" && v.startsWith("ENUM:")) {
    const [e, m] = v.slice(5).split(".");
    return (pptx as any)[e][m];
  }
  return v;
}

function apply(chart: any, path: string, ...args: any[]): void {
  if (path === "replace_data") {
    chart.replaceData(makeData(SPEC.data[args[0]]));
    return;
  }
  const parts = path.split(".");
  let obj = chart;
  for (const part of parts.slice(0, -1)) obj = /^\d+$/.test(part) ? obj.at(Number(part)) : obj[camel(part)];
  const last = parts.at(-1)!;
  if (last.endsWith("()")) obj[camel(last.slice(0, -2))]();
  else obj[camel(last)] = value(args[0]);
}

const prs = Presentation();
for (const c of SPEC.charts) {
  const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
  const gf = slide.shapes.addChart((XL_CHART_TYPE as any)[c.type], Inches(1), Inches(1), Inches(6), Inches(4.5), makeData(SPEC.data[c.data]));
  for (const op of c.ops ?? []) apply(gf.chart, op[0], ...op.slice(1));
}
const buf = new BytesIO();
prs.save(buf);
if (process.argv[3]) writeFileSync(process.argv[3], buf.getvalue());

// -- compare -------------------------------------------------------------------------------------
const dec = new TextDecoder();
const ts = unzipSync(buf.getvalue());
const py = unzipSync(new Uint8Array(readFileSync(process.argv[2] ?? "charts_py.pptx")));
let same = 0;
const diffs: string[] = [];
const eq = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((x, i) => x === b[i]);
const firstDiff = (a: string, b: string) => {
  let i = 0;
  while (i < a.length && a[i] === b[i]) i++;
  return `@${i}\n      ts: ${JSON.stringify(a.slice(Math.max(0, i - 60), i + 80))}\n      py: ${JSON.stringify(b.slice(Math.max(0, i - 60), i + 80))}`;
};
for (const name of [...new Set([...Object.keys(ts), ...Object.keys(py)])].sort()) {
  const a = ts[name];
  const b = py[name];
  if (!a || !b) {
    diffs.push(`${name}: only in ${a ? "ts" : "py"}`);
    continue;
  }
  if (name.endsWith(".xlsx")) {
    const za = unzipSync(a);
    const zb = unzipSync(b);
    const names = [...new Set([...Object.keys(za), ...Object.keys(zb)])].sort();
    const bad = names.filter((n) => n !== "docProps/core.xml" && !(za[n] && zb[n] && eq(za[n], zb[n])));
    if (Object.keys(za).join() !== Object.keys(zb).join()) bad.push(`(member order: ts ${Object.keys(za).join(",")} / py ${Object.keys(zb).join(",")})`);
    if (bad.length === 0) same++;
    else for (const n of bad) diffs.push(`${name} > ${n}: ${za[n] && zb[n] ? firstDiff(dec.decode(za[n]), dec.decode(zb[n])) : "missing"}`);
    continue;
  }
  if (eq(a, b)) same++;
  else diffs.push(`${name}: ${firstDiff(dec.decode(a), dec.decode(b))}`);
}
for (const d of diffs) console.log("DIFF", d);
console.log(`${same}/${same + diffs.length} parts identical (${SPEC.charts.length} charts)`);
