// Summarize a vitest JSON report: per file pass/fail counts and the first line of each failure.
// Usage: npx vitest run --reporter=json --outputFile=.vitest.json [files]; node scripts/failures.mjs [.vitest.json] [maxPerFile]
import fs from "node:fs";
const r = JSON.parse(fs.readFileSync(process.argv[2] ?? ".vitest.json", "utf8"));
const max = Number(process.argv[3] ?? 6);
let pass = 0, fail = 0;
for (const f of r.testResults) {
  const rel = f.name.replace(/.*pptx-ts[\/]/, "");
  const fails = f.assertionResults.filter((a) => a.status === "failed");
  pass += f.assertionResults.filter((a) => a.status === "passed").length;
  fail += fails.length;
  if (!fails.length && f.status !== "failed") continue;
  console.log(`\n== ${rel}: ${fails.length} failed / ${f.assertionResults.length}` + (f.message ? `  [${f.message.split("\n")[0].slice(0, 200)}]` : ""));
  for (const a of fails.slice(0, max)) console.log(`  x ${a.fullName.slice(0, 110)}\n      ${(a.failureMessages[0] ?? "").split("\n")[0].slice(0, 220)}`);
}
console.log(`\npassed ${pass}, failed ${fail}`);
