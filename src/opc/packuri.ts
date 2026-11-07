// The PackURI value type and known pack-URI strings (`pptx.opc.packuri`).
import { ValueError } from "../errors.js";
import { pyReprStr } from "../util.js";

// -- posixpath helpers ---------------------------------------------------------------------------

export const posixpath = {
  join(a: string, ...parts: string[]): string {
    let path = a;
    for (const b of parts) {
      if (b.startsWith("/")) path = b;
      else if (path === "" || path.endsWith("/")) path += b;
      else path += "/" + b;
    }
    return path;
  },
  normpath(path: string): string {
    if (path === "") return ".";
    const initial = path.startsWith("/") ? (path.startsWith("//") && !path.startsWith("///") ? 2 : 1) : 0;
    const comps: string[] = [];
    for (const c of path.split("/")) {
      if (c === "" || c === ".") continue;
      if (c !== ".." || (!initial && !comps.length) || (comps.length && comps[comps.length - 1] === "..")) comps.push(c);
      else if (comps.length) comps.pop();
    }
    const out = "/".repeat(initial) + comps.join("/");
    return out || ".";
  },
  /** abspath with "/" as the working directory */
  abspath(path: string): string {
    return posixpath.normpath(path.startsWith("/") ? path : "/" + path);
  },
  split(p: string): [string, string] {
    const i = p.lastIndexOf("/") + 1;
    let head = p.slice(0, i);
    const tail = p.slice(i);
    if (head && head !== "/".repeat(head.length)) head = head.replace(/\/+$/, "");
    return [head, tail];
  },
  splitext(p: string): [string, string] {
    const sep = p.lastIndexOf("/");
    const dot = p.lastIndexOf(".");
    if (dot > sep) {
      // skip leading dots of the filename
      let i = sep + 1;
      while (i < dot && p[i] === ".") i++;
      if (i < dot) return [p.slice(0, dot), p.slice(dot)];
    }
    return [p, ""];
  },
  relpath(path: string, start: string): string {
    const s = posixpath.abspath(start).split("/").filter(Boolean);
    const p = posixpath.abspath(path).split("/").filter(Boolean);
    let i = 0;
    while (i < s.length && i < p.length && s[i] === p[i]) i++;
    const rel = [...Array(s.length - i).fill(".."), ...p.slice(i)];
    return rel.length ? rel.join("/") : ".";
  },
};

const FILENAME_RE = /^([a-zA-Z]+)([0-9][0-9]*)?/;

/**
 * Proxy for a pack URI (partname), e.g. "/ppt/slides/slide1.xml". Behaves as a string
 * (`String(uri)`, template literals, string methods) and adds the OPC-specific properties.
 */
export class PackURI extends String {
  constructor(packUriStr: string | String) {
    const s = String(packUriStr);
    if (s[0] !== "/") throw new ValueError(`PackURI must begin with slash, got ${pyReprStr(s)}`);
    super(s);
  }

  /** Absolute pack URI formed by translating `relativeRef` onto `baseURI`. */
  static fromRelRef(baseURI: string | String, relativeRef: string): PackURI {
    const joined = posixpath.join(String(baseURI), relativeRef);
    return new PackURI(posixpath.abspath(joined));
  }

  /** The directory portion, e.g. "/ppt/slides" for "/ppt/slides/slide1.xml"; "/" for "/". */
  get baseURI(): string {
    return posixpath.split(this.valueOf())[0];
  }

  /** The extension, e.g. "xml"; no leading period. */
  get ext(): string {
    const raw = posixpath.splitext(this.valueOf())[1];
    return raw.startsWith(".") ? raw.slice(1) : raw;
  }

  /** The "filename" portion, e.g. "slide1.xml"; '' for "/". */
  get filename(): string {
    return posixpath.split(this.valueOf())[1];
  }

  /** Partname index, e.g. 21 for "/ppt/slides/slide21.xml"; null for a singleton partname. */
  get idx(): number | null {
    const filename = this.filename;
    if (!filename) return null;
    const namePart = posixpath.splitext(filename)[0];
    const m = FILENAME_RE.exec(namePart);
    if (!m) return null;
    return m[2] ? parseInt(m[2], 10) : null;
  }

  /** The pack URI with the leading slash stripped off (the zip membername). */
  get membername(): string {
    return this.valueOf().slice(1);
  }

  /** Relative reference to this package item from `baseURI`. */
  relativeRef(baseURI: string | String): string {
    const b = String(baseURI);
    return b === "/" ? this.valueOf().slice(1) : posixpath.relpath(this.valueOf(), b);
  }

  /** The pack URI of the .rels part corresponding to this pack URI. */
  get relsUri(): PackURI {
    return new PackURI(posixpath.join(this.baseURI, "_rels", `${this.filename}.rels`));
  }

  equals(other: unknown): boolean {
    return (typeof other === "string" || other instanceof String) && String(other) === this.valueOf();
  }

  lower(): string {
    return this.valueOf().toLowerCase();
  }
}

export const PACKAGE_URI = new PackURI("/");
export const CONTENT_TYPES_URI = new PackURI("/[Content_Types].xml");
