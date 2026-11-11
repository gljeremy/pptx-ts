// Bundle the built package for the browser, as an app would, and check it: the main entry must
// bundle without Node-only imports, stay under a size budget, leave the OLE icons out (they come
// with `@gljeremy/pptx-ts/ole-icons`) and still work: the bundled code opens a deck with a chart, reads it
// and saves it. (A bundler drops modules it thinks are unused; the element classes register
// themselves on import, so the package must not claim to be free of side effects.)
// Usage: npm run build && node scripts/check_bundle.mjs
import assert from "node:assert/strict";
import { unzipSync } from "fflate";
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
const ole = await bundle(`
import { Presentation, BytesIO, PROG_ID } from "./dist/index.js";
const save = (prs) => {
  const out = new BytesIO();
  prs.save(out);
  return out.getvalue();
};
export async function exerciseOle(payload, customIcon) {
  const cases = [];
  for (const grouped of [false, true]) {
    for (const progId of [PROG_ID.DOCX, "Acme.Document", PROG_ID.PPTX, PROG_ID.XLSX]) {
      const prs = Presentation();
      const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
      const shapes = grouped ? slide.shapes.addGroupShape().shapes : slide.shapes;
      const before = save(prs);
      let error = null;
      try {
        shapes.addOleObject(payload, progId, 0, 0);
      } catch (e) {
        error = { name: e.name, message: e.message };
      }
      cases.push({ prs, shapes, progId, before, error, afterFailure: save(prs) });
    }
  }
  const custom = Presentation();
  const customSlide = custom.slides.addSlide(custom.slideLayouts.at(6));
  customSlide.shapes.addOleObject(payload, "Acme.Document", 0, 0, null, null, customIcon);
  const customSaved = save(custom);
  await import("./dist/ole-icons.js");
  return {
    customSaved,
    cases: cases.map(({ prs, shapes, progId, ...result }) => {
      shapes.addOleObject(payload, progId, 0, 0);
      const saved = save(prs);
      const reopened = Presentation(saved);
      const root = reopened.slides.at(0).shapes;
      const objects = root.at(0).shapes ?? root;
      return { ...result, saved, count: objects.length, progId: objects.at(0).oleFormat.progId };
    }),
  };
}`);
console.log(`@gljeremy/pptx-ts              ${kb(main.bytes)} minified, ${kb(main.gzip)} gzip (budget ${kb(BUDGET_GZIP)})`);
console.log(`+ @gljeremy/pptx-ts/ole-icons  ${kb(ole.bytes)} minified, ${kb(ole.gzip)} gzip`);

let failed = false;
const fail = (msg) => {
  console.error(`FAIL: ${msg}`);
  failed = true;
};
// "AQAAAG" starts the base64 of every EMF icon (its EMR_HEADER record); only icons.ts has it
if (main.code.includes("AQAAAG")) fail("the OLE icons are in the main bundle");
if (main.gzip > BUDGET_GZIP) fail(`main bundle ${kb(main.gzip)} gzip is over the ${kb(BUDGET_GZIP)} budget`);


// run the bundle (it has no Node imports, so a data: URL module works)
const { run } = await import(`data:text/javascript;base64,${Buffer.from(main.code).toString("base64")}`);
const deck = new Uint8Array(readFileSync("test/features/test_files/cht-chart-props.pptx"));
const r = run(deck);
console.log(`bundled code opened a ${r.chartType} chart with ${r.series} series and saved ${kb(r.size)}`);
if (!(r.series > 0 && r.size > 10000)) fail("the bundled code didn't read and save the chart deck");
const { exerciseOle } = await import(`data:text/javascript;base64,${Buffer.from(ole.code).toString("base64")}`);
const payload = new Uint8Array(readFileSync("test/features/test_files/shp-embedded-xlsx.xlsx"));
const iconNames = ["docx-icon.emf", "generic-icon.emf", "pptx-icon.emf", "xlsx-icon.emf"];
const progIds = ["Word.Document.12", "Acme.Document", "PowerPoint.Show.12", "Excel.Sheet.12"];
const iconFile = (name) => new Uint8Array(readFileSync(`templates/${name}`));
const result = await exerciseOle(payload, iconFile("generic-icon.emf"));
const entries = (bytes) => Object.fromEntries(
  Object.entries(unzipSync(bytes)).sort(([a], [b]) => a.localeCompare(b)).map(([name, data]) => [name, Buffer.from(data)]),
);
const checkEmbedded = (bytes, expectedIcon) => {
  const files = entries(bytes);
  const attachments = Object.keys(files).filter((name) => name.startsWith("ppt/embeddings/"));
  const images = Object.keys(files).filter((name) => name.startsWith("ppt/media/"));
  assert.equal(attachments.length, 1, "exactly one attachment after insertion or retry");
  assert.deepEqual(files[attachments[0]], Buffer.from(payload));
  assert.equal(images.length, 1, "exactly one OLE icon");
  assert.deepEqual(files[images[0]], Buffer.from(expectedIcon), "saved icon matches the shipped EMF file");
};
for (const [i, c] of result.cases.entries()) {
  assert.equal(c.error?.name, "KeyError");
  assert.match(c.error.message, /import "@gljeremy\/pptx-ts\/ole-icons"/);
  assert.deepEqual(entries(c.afterFailure), entries(c.before), "missing-icon failure leaves every ZIP entry unchanged");
  assert.equal(c.count, 1);
  assert.equal(c.progId, progIds[i % 4]);
  checkEmbedded(c.saved, iconFile(iconNames[i % 4]));
}
checkEmbedded(result.customSaved, iconFile("generic-icon.emf"));
console.log("browser OLE: all four icons on slides and groups fail cleanly, register, save and reopen; custom icons work without registration");
process.exit(failed ? 1 : 0);
