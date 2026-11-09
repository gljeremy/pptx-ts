// python-pptx's behave acceptance tests (features/*.feature), run against pptx-ts with the step
// definitions translated from features/steps/*.py. Scenarios that can't pass for a documented
// reason are listed in skips.json (scenario name -> reason; see PARITY.md).
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineFeature } from "./runner.js";

// behave imports the step modules in sorted order; the first matching step wins
const modules = import.meta.glob("./steps/*.ts", { eager: true });
void modules;

const dir = path.dirname(fileURLToPath(import.meta.url));
const skips: Record<string, Record<string, string>> = JSON.parse(readFileSync(path.join(dir, "skips.json"), "utf8"));
for (const file of readdirSync(dir).filter((f) => f.endsWith(".feature")).sort()) {
  defineFeature(file, readFileSync(path.join(dir, file), "utf8"), skips[file] ?? {});
}
