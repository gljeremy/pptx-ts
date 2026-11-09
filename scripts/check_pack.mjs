// Pack the library with `npm pack`, install the tarball in a fresh Vite project (as a user
// would) and check it: `vite build` bundles it for the browser with the OLE icons in their own
// chunk, and Node opens, edits and saves decks through both package entries.
// This checks browser build compatibility, not the file-input UI or a download.
// Usage: npm run build && node scripts/check_pack.mjs
import { execSync } from "node:child_process";
import { cpSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const dir = mkdtempSync(path.join(tmpdir(), "pptx-ts-pack-"));
const sh = (cmd, cwd = dir) => execSync(cmd, { cwd, stdio: ["ignore", "pipe", "inherit"] }).toString();

try {
  const tgz = sh(`npm pack --ignore-scripts --pack-destination "${dir}"`, root).trim().split(/\r?\n/).at(-1);
  console.log(`packed ${tgz}`);
  mkdirSync(path.join(dir, "public"));
  cpSync(path.join(root, "test/features/test_files/cht-chart-props.pptx"), path.join(dir, "public/sample.pptx"));
  cpSync(path.join(root, "test/features/test_files/shp-embedded-xlsx.xlsx"), path.join(dir, "public/sheet.xlsx"));
  writeFileSync(path.join(dir, "package.json"), JSON.stringify({ name: "pptx-ts-pack-check", private: true, type: "module" }));
  writeFileSync(
    path.join(dir, "index.html"),
    '<!doctype html><html><body><input type="file" id="file"><pre id="status"></pre><script type="module" src="/main.js"></script></body></html>',
  );
  writeFileSync(
    path.join(dir, "main.js"),
    `import { Presentation, BytesIO } from "pptx-ts";
document.getElementById("file").addEventListener("change", async (e) => {
  const prs = Presentation(new Uint8Array(await e.target.files[0].arrayBuffer()));
  prs.slides.addSlide(prs.slideLayouts.at(0));
  const out = new BytesIO();
  prs.save(out);
  document.getElementById("status").textContent = String(out.getvalue().length);
});
window.ole = () => import("pptx-ts/ole-icons");
`,
  );
  sh(`npm install --no-audit --no-fund --loglevel=error ./${tgz} vite@5.4.11`);

  // browser: vite build
  sh("npx vite build --logLevel warn");
  const assets = readdirSync(path.join(dir, "dist/assets"));
  console.log(`vite build: ${assets.join(", ")}`);
  if (!assets.some((a) => a.startsWith("ole-icons-"))) throw new Error("the OLE icons are not in a chunk of their own");

  // node: both entries, from the installed package
  writeFileSync(
    path.join(dir, "node-check.mjs"),
    `import { readFileSync } from "node:fs";
import { BytesIO, Presentation, PROG_ID } from "pptx-ts";
const prs = Presentation(new Uint8Array(readFileSync("public/sample.pptx")));
const charts = [...prs.slides].flatMap((s) => [...s.shapes]).filter((sh) => sh.hasChart).length;
const slide = prs.slides.addSlide(prs.slideLayouts.at(0));
slide.shapes.addOleObject(new BytesIO(readFileSync("public/sheet.xlsx")), PROG_ID.XLSX, 914400, 914400);
const out = new BytesIO();
prs.save(out);
const re = Presentation(out.getvalue());
const ole = [...re.slides.at(-1).shapes][0];
await import("pptx-ts/ole-icons");
if (charts !== 4 || re.slides.length !== 3 || ole.oleFormat.progId !== "Excel.Sheet.12") throw new Error("unexpected result");
console.log("node: opened 4 charts, added an OLE object, saved and reopened");
`,
  );
  process.stdout.write(sh("node node-check.mjs"));
  console.log("pack check passed");
} finally {
  rmSync(dir, { recursive: true, force: true });
}
