// Bundle the built package for the browser, as an app would, and check it: the main entry must
// bundle without Node-only imports, stay under a size budget, leave the OLE icons out (they come
// with `pptx-ts/ole-icons`) and still work: the bundled code opens a deck with a chart, reads it
// and saves it. (A bundler drops modules it thinks are unused; the element classes register
// themselves on import, so the package must not claim to be free of side effects.)
// Usage: npm run build && node scripts/check_bundle.mjs
import { build } from "esbuild";
import { readFileSync } from "node:fs";
import { gzipSync } from "node:zlib";

const BUDGET_GZIP = 150 * 1024; // main entry, minified + gzip

async function bundle(contents) {
  const r = await build({
    stdin: { contents, resolveDir: process.cwd(), loader: "js" },
    bundle: true,
    platform: "browser",
    format: "esm",
    minify: true,
    write: false,
    logLevel: "silent",
  });
  const code = r.outputFiles[0].contents;
  return { bytes: code.length, gzip: gzipSync(code).length, code: new TextDecoder().decode(code) };
}

const kb = (n) => `${(n / 1024).toFixed(1)} KB`;
const main = await bundle(`
import { Presentation, BytesIO } from "./dist/index.js";
export function run(bytes) {
  const prs = Presentation(bytes);
  const chart = [...prs.slides.at(0).shapes].find((s) => s.hasChart).chart;
  const out = new BytesIO();
  prs.save(out);
  return { chartType: chart.chartType.name, series: chart.plots.at(0).series.length, size: out.getvalue().length };
}`);
const ole = await bundle('import "./dist/ole-icons.js"; import { Presentation } from "./dist/index.js"; export const p = Presentation();');
console.log(`pptx-ts              ${kb(main.bytes)} minified, ${kb(main.gzip)} gzip (budget ${kb(BUDGET_GZIP)})`);
console.log(`+ pptx-ts/ole-icons  ${kb(ole.bytes)} minified, ${kb(ole.gzip)} gzip`);

let failed = false;
const fail = (msg) => {
  console.error(`FAIL: ${msg}`);
  failed = true;
};
// "AQAAAG" starts the base64 of every EMF icon (its EMR_HEADER record); only icons.ts has it
if (main.code.includes("AQAAAG")) fail("the OLE icons are in the main bundle");
if (main.gzip > BUDGET_GZIP) fail(`main bundle ${kb(main.gzip)} gzip is over the ${kb(BUDGET_GZIP)} budget`);
if (ole.bytes - main.bytes < 300 * 1024) fail("pptx-ts/ole-icons doesn't add the icons");

// run the bundle (it has no Node imports, so a data: URL module works)
const { run } = await import(`data:text/javascript;base64,${Buffer.from(main.code).toString("base64")}`);
const deck = new Uint8Array(readFileSync("test/features/test_files/cht-chart-props.pptx"));
const r = run(deck);
console.log(`bundled code opened a ${r.chartType} chart with ${r.series} series and saved ${kb(r.size)}`);
if (!(r.series > 0 && r.size > 10000)) fail("the bundled code didn't read and save the chart deck");
process.exit(failed ? 1 : 0);
