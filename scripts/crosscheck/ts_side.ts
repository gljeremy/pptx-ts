// Compare pptx-ts with python-pptx on every upstream test deck: the same API answers and
// byte-identical parts after re-saving. Run scripts/crosscheck/py_side.py first.
// Usage: npx tsx scripts/crosscheck/ts_side.ts py.json <python-pptx checkout>
// Expected difference: decks without core properties get a fresh `modified` time on each side.
import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { unzipSync } from "fflate";
import { Presentation, BytesIO } from "../../src/index.js";
const py = JSON.parse(readFileSync(process.argv[2] ?? "py.json", "utf8"));
const root = (process.argv[3] ?? "../python-pptx").replace(/\\/g, "/").replace(/\/?$/, "/");
const pyStr = (v: any) => {
  if (v === null || v === undefined) return "None";
  if (v instanceof Date) {
    const p = (n: number) => String(n).padStart(2, "0");
    return `${v.getUTCFullYear()}-${p(v.getUTCMonth() + 1)}-${p(v.getUTCDate())} ${p(v.getUTCHours())}:${p(v.getUTCMinutes())}:${p(v.getUTCSeconds())}`;
  }
  return String(v);
};
const shapeInfo = (sh: any) => {
  const d: any = { name: sh.name, id: sh.shapeId, ph: sh.isPlaceholder };
  for (const a of ["left", "top", "width", "height"]) { try { d[a] = sh[a]; } catch { d[a] = "ERR"; } }
  if (sh.isPlaceholder) { const pf = sh.placeholderFormat; d.phIdx = pf.idx; d.phType = pf.type?.name ?? null; }
  return d;
};
const camelCore = { author: "author", category: "category", comments: "comments", content_status: "contentStatus", created: "created", identifier: "identifier", keywords: "keywords", language: "language", last_modified_by: "lastModifiedBy", last_printed: "lastPrinted", modified: "modified", revision: "revision", subject: "subject", title: "title", version: "version" } as const;
let ok = 0, bad = 0;
for (const key of Object.keys(py)) {
  let info: any;
  try {
    const prs = Presentation(new Uint8Array(readFileSync(root + key)));
    info = {
      slides: [...prs.slides].map((s) => ({ id: s.slideId, name: s.name, layout: s.slideLayout.name, shapes: [...s.shapes].map(shapeInfo), placeholders: [...s.placeholders].map((p: any) => p.placeholderFormat.idx), hasNotes: s.hasNotesSlide, followBg: s.followMasterBackground })),
      masters: [...prs.slideMasters].map((m) => ({ layouts: [...m.slideLayouts].map((l) => l.name), shapes: [...m.shapes].map((x) => x.name) })),
      size: [prs.slideWidth, prs.slideHeight],
      core: Object.fromEntries(Object.entries(camelCore).map(([k, c]) => [k, pyStr((prs.coreProperties as any)[c])])),
    };
    const notes: any[] = [];
    for (const s of prs.slides) if (s.hasNotesSlide) { const tf = s.notesSlide.notesTextFrame; notes.push(tf !== null ? tf.text : null); }
    info.notes = notes;
    const buf = new BytesIO(); prs.save(buf);
    const z = unzipSync(buf.getvalue());
    info.saved = Object.fromEntries(Object.entries(z).map(([n, b]) => [n, createHash("sha1").update(b).digest("hex")]));
  } catch (e: any) {
    info = { error: `${e.name}: ${String(e.message).slice(0, 120)}` };
  }
  const want = py[key];
  const a = JSON.stringify(info), b = JSON.stringify(want);
  if (a === b) { ok++; continue; }
  bad++;
  console.log("MISMATCH", key);
  for (const k of new Set([...Object.keys(want), ...Object.keys(info)])) {
    const x = JSON.stringify(info[k]), y = JSON.stringify(want[k]);
    if (x !== y) console.log("  ", k, "\n     ts:", x?.slice(0, 300), "\n     py:", y?.slice(0, 300));
  }
}
console.log(`${ok}/${ok + bad} decks match`);
