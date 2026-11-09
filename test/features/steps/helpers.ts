// Port of python-pptx features/steps/helpers.py: paths to the acceptance test files and a
// scratch directory for saved decks (a temp directory instead of features/_scratch).
import { mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const thisdir = path.dirname(fileURLToPath(import.meta.url));
const testFilesDir = path.join(thisdir, "..", "test_files");
export const scratch_dir = path.join(tmpdir(), `pptx-ts-features-${process.pid}`);
mkdirSync(scratch_dir, { recursive: true });

// legacy test pptx files
export const no_core_props_pptx_path = path.join(thisdir, "..", "..", "test_files", "no-core-props.pptx");

// scratch test pptx file
export const saved_pptx_path = path.join(scratch_dir, "test_out.pptx");

export const test_text = "python-pptx was here!";

const srcModules = import.meta.glob("../../../src/**/*.ts", { eager: true }) as Record<string, Record<string, unknown>>;

/** python `"%s.%s" % (obj.__module__, obj.__class__.__name__)`: the src module that defines the class */
export function cls_qname(obj: object): string {
  const cls = obj.constructor;
  for (const [file, exports] of Object.entries(srcModules)) {
    if (Object.values(exports).includes(cls)) {
      const mod = file.replace(/^.*\/src\//, "").replace(/(\/index)?\.ts$/, "").replace(/\//g, ".");
      return `pptx.${mod}.${cls.name}`;
    }
  }
  return `?.${cls.name}`;
}

/** python `itertools.count()` */
export function* count(start = 0, step = 1): Generator<number> {
  let n = start;
  for (;;) {
    yield n;
    n += step;
  }
}

/** Absolute path to `filename` in the acceptance test_files directory. */
export function test_file(filename: string): string {
  return path.join(testFilesDir, filename);
}

/** Absolute path to image file `filename` in the test_files directory. */
export function test_image(filename: string): string {
  return path.join(testFilesDir, filename);
}

/** Absolute path to test .pptx file with root name `name`. */
export function test_pptx(name: string): string {
  return path.join(testFilesDir, `${name}.pptx`);
}
