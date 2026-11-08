// File access that works in browsers and Node. Paths need Node (`node:fs` is loaded on
// demand); everywhere else pass bytes or a file-like object.
import { FileNotFoundError } from "./errors.js";

/** An object with a python-style `read()` (and optionally `seek()`), e.g. `BytesIO`. */
export interface ReadableFile {
  read(size?: number): Uint8Array;
  seek?(pos: number): unknown;
}

/** An object with a python-style `write(bytes)`, e.g. `BytesIO`. */
export interface WritableFile {
  write(data: Uint8Array): unknown;
}

/** A path (Node only), bytes, or a readable file-like object. */
export type FileSource = string | Uint8Array | ArrayBuffer | ReadableFile;
/** A path (Node only) or a writable file-like object. */
export type FileTarget = string | WritableFile;

/** In-memory binary stream (python `io.BytesIO`). */
export class BytesIO implements ReadableFile, WritableFile {
  private buf: Uint8Array;
  private len: number;
  private pos = 0;

  constructor(initial?: Uint8Array | ArrayBuffer) {
    this.buf = initial ? new Uint8Array(initial instanceof Uint8Array ? initial : new Uint8Array(initial)).slice() : new Uint8Array(256);
    this.len = initial ? this.buf.length : 0;
  }

  read(size = -1): Uint8Array {
    const end = size < 0 ? this.len : Math.min(this.len, this.pos + size);
    const out = this.buf.slice(this.pos, end);
    this.pos = end;
    return out;
  }

  write(data: Uint8Array | ArrayBuffer): number {
    const b = data instanceof Uint8Array ? data : new Uint8Array(data);
    const need = this.pos + b.length;
    if (need > this.buf.length) {
      const nb = new Uint8Array(Math.max(need, this.buf.length * 2));
      nb.set(this.buf.subarray(0, this.len));
      this.buf = nb;
    }
    this.buf.set(b, this.pos);
    this.pos = need;
    if (need > this.len) this.len = need;
    return b.length;
  }

  /** No-op close (python `BytesIO.close()`; the buffer stays readable here). */
  close(): void {}

  seek(pos: number, whence = 0): number {
    this.pos = whence === 0 ? pos : whence === 1 ? this.pos + pos : this.len + pos;
    return this.pos;
  }

  tell(): number {
    return this.pos;
  }

  getvalue(): Uint8Array {
    return this.buf.slice(0, this.len);
  }
}

type NodeFs = {
  readFileSync(p: string): Uint8Array;
  writeFileSync(p: string, d: Uint8Array): void;
  statSync(p: string, o?: { throwIfNoEntry?: boolean }): { isDirectory(): boolean; isFile(): boolean } | undefined;
  existsSync(p: string): boolean;
};

let fsMod: NodeFs | null | undefined;

/** `node:fs`, or null when not running in Node. */
export function nodeFs(): NodeFs | null {
  if (fsMod === undefined) {
    const proc = (globalThis as any).process;
    fsMod = proc?.getBuiltinModule ? (proc.getBuiltinModule("node:fs") as NodeFs) : null;
  }
  return fsMod;
}

function requireFs(what: string): NodeFs {
  const fs = nodeFs();
  if (!fs) throw new TypeError(`${what} needs Node.js file system access; pass bytes or a file-like object instead`);
  return fs;
}

export function readFileBytes(path: string): Uint8Array {
  const fs = requireFs("reading a file by path");
  try {
    return new Uint8Array(fs.readFileSync(path));
  } catch (e: any) {
    if (e?.code === "ENOENT") throw new FileNotFoundError(`[Errno 2] No such file or directory: '${path}'`);
    throw e;
  }
}

export function writeFileBytes(path: string, data: Uint8Array): void {
  requireFs("writing a file by path").writeFileSync(path, data);
}

export function isDir(path: string): boolean {
  return !!requireFs("opening a path").statSync(path, { throwIfNoEntry: false })?.isDirectory();
}

export function isFile(path: string): boolean {
  return !!requireFs("opening a path").statSync(path, { throwIfNoEntry: false })?.isFile();
}

export function pathExists(path: string): boolean {
  return requireFs("opening a path").existsSync(path);
}

/** Bytes of `src`: read from a path, taken from bytes, or read from a file-like object (after seek(0)). */
export function bytesOf(src: FileSource): Uint8Array {
  if (typeof src === "string") return readFileBytes(src);
  if (src instanceof Uint8Array) return src;
  if (src instanceof ArrayBuffer) return new Uint8Array(src);
  if (typeof src.seek === "function") src.seek(0);
  const b = src.read();
  return b instanceof Uint8Array ? b : new Uint8Array(b as any);
}

export const isZipBytes = (b: Uint8Array): boolean => b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && (b[2] === 3 || b[2] === 5 || b[2] === 7);

/** python `os.path.basename()` (accepts both / and \ separators). */
export function basename(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return path.slice(i + 1);
}

/** python `os.path.splitext(path)[1]`, e.g. ".mp4" ("" when there's none). */
export function extname(path: string): string {
  const base = basename(path);
  const i = base.lastIndexOf(".");
  return i <= 0 || /^\.+$/.test(base.slice(0, i + 1)) ? "" : base.slice(i);
}
