// Access to the files bundled from python-pptx's pptx/templates (see data.ts).
//
// The OLE object icons (*.emf) are not in data.ts, to keep them out of browser bundles that don't
// embed OLE objects. `import "@gljeremy/pptx-ts/ole-icons"` registers them; in Node they are read from the
// package's templates/ directory when not registered, so `addOleObject()` works there as is.
import { KeyError } from "../errors.js";
import { nodeFs } from "../fileio.js";
import { BINARY, TEXT } from "./data.js";

const registered: Record<string, string> = {};

/** Text of the XML template `name` (without the .xml extension). */
export function templateText(name: string): string {
  const t = TEXT[name];
  if (t === undefined) throw new KeyError(`no template named '${name}'`);
  return t;
}

/** Make base64 template files (the OLE icons) available to `templateBytes()`. */
export function registerTemplateFiles(files: Record<string, string>): void {
  Object.assign(registered, files);
}

/** Bytes of the binary template file `filename` (e.g. "default.pptx"). */
export function templateBytes(filename: string): Uint8Array {
  const b64 = BINARY[filename] ?? registered[filename];
  if (b64 !== undefined) return base64Decode(b64);
  if (filename.endsWith(".emf")) {
    const bytes = readPackageTemplate(filename);
    if (bytes) return bytes;
    throw new KeyError(
      `the OLE object icon '${filename}' is not loaded: \`import "@gljeremy/pptx-ts/ole-icons"\` once before ` +
        "addOleObject(), or pass an icon file",
    );
  }
  throw new KeyError(`no template file named '${filename}'`);
}

/** templates/<filename> next to this package's dist/ (or src/), in Node; null elsewhere. */
function readPackageTemplate(filename: string): Uint8Array | null {
  const fs = nodeFs();
  if (!fs || !import.meta.url.startsWith("file:")) return null;
  // (not `new URL(\`...${filename}\`, import.meta.url)`: bundlers turn that into an asset glob)
  const proc = (globalThis as any).process;
  const { fileURLToPath } = proc.getBuiltinModule("node:url");
  const { dirname, join } = proc.getBuiltinModule("node:path");
  const path = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "templates", filename);
  return fs.existsSync(path) ? new Uint8Array(fs.readFileSync(path)) : null;
}

export function base64Decode(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
