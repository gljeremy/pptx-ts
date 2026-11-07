// Port of python-pptx tests/unitutil/file.py: loading files for unit testing.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const thisdir = path.dirname(fileURLToPath(import.meta.url));
export const testFileDir = path.resolve(thisdir, "..", "test_files");
export const test_file_dir = testFileDir;

export function absjoin(...paths: string[]): string {
  return path.resolve(path.join(...paths));
}

/** Bytes read from snippet file `name` (stripped). */
export function snippetBytes(name: string): Uint8Array {
  const b = fs.readFileSync(path.join(testFileDir, "snippets", `${name}.txt`));
  let s = 0;
  let e = b.length;
  while (s < e && /\s/.test(String.fromCharCode(b[s]))) s++;
  while (e > s && /\s/.test(String.fromCharCode(b[e - 1]))) e--;
  return new Uint8Array(b.subarray(s, e));
}

/** Text snippets read from snippet file `name`, delimited by a blank line. */
export function snippetSeq(name: string, offset = 0, count = Number.MAX_SAFE_INTEGER): string[] {
  const text = fs.readFileSync(path.join(testFileDir, "snippets", `${name}.txt`), "utf8");
  return text.split("\n\n").slice(offset, offset + count);
}

export function snippetText(name: string): string {
  return fs.readFileSync(path.join(testFileDir, "snippets", `${name}.txt`), "utf8");
}

/** Absolute path to test file `name`. */
export function testfile(name: string): string {
  return absjoin(testFileDir, name);
}

/** Bytes of the file at the path formed by adding `segments` to the test file dir. */
export function testfileBytes(...segments: string[]): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(testFileDir, ...segments)));
}
