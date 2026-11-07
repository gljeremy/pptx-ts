// Access to the files bundled from python-pptx's pptx/templates (see data.ts).
import { KeyError } from "../errors.js";
import { BINARY, TEXT } from "./data.js";

/** Text of the XML template `name` (without the .xml extension). */
export function templateText(name: string): string {
  const t = TEXT[name];
  if (t === undefined) throw new KeyError(`no template named '${name}'`);
  return t;
}

/** Bytes of the binary template file `filename` (e.g. "default.pptx"). */
export function templateBytes(filename: string): Uint8Array {
  const b64 = BINARY[filename];
  if (b64 === undefined) throw new KeyError(`no template file named '${filename}'`);
  return base64Decode(b64);
}

export function base64Decode(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
