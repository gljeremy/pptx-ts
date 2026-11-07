// Namespace related objects (`pptx.oxml.ns`).
import { KeyError } from "../errors.js";
import { setPreferredPrefixes } from "./xml.js";

/** Namespace prefix -> namespace name for all known PowerPoint XML namespaces. */
export const _nsmap: Record<string, string> = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  dc: "http://purl.org/dc/elements/1.1/",
  dcmitype: "http://purl.org/dc/dcmitype/",
  dcterms: "http://purl.org/dc/terms/",
  ep: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  i: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  mo: "http://schemas.microsoft.com/office/mac/office/2008/main",
  mv: "urn:schemas-microsoft-com:mac:vml",
  o: "urn:schemas-microsoft-com:office:office",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  pd: "http://schemas.openxmlformats.org/drawingml/2006/presentationDrawing",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  pr: "http://schemas.openxmlformats.org/package/2006/relationships",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  sl: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  v: "urn:schemas-microsoft-com:vml",
  ve: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w10: "urn:schemas-microsoft-com:office:word",
  wne: "http://schemas.microsoft.com/office/word/2006/wordml",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
};

export const pfxmap: Record<string, string> = Object.fromEntries(Object.entries(_nsmap).map(([k, v]) => [v, k]));

setPreferredPrefixes(pfxmap);

const clarkCache = new Map<string, string>();

function nsLookup(pfx: string): string {
  const uri = _nsmap[pfx];
  if (uri === undefined) throw new KeyError(`'${pfx}'`);
  return uri;
}

/** Value object that knows the semantics of an XML tag having a namespace prefix. */
export class NamespacePrefixedTag extends String {
  readonly _pfx: string;
  readonly _localPart: string;
  readonly _nsUri: string;

  constructor(nstag: string) {
    super(nstag);
    const parts = nstag.split(":");
    if (parts.length !== 2) throw new Error(`ValueError: not enough values to unpack (expected 2, got ${parts.length})`);
    [this._pfx, this._localPart] = parts;
    this._nsUri = nsLookup(this._pfx);
  }

  static fromClarkName(clarkName: string): NamespacePrefixedTag {
    const [nsuri, localName] = clarkName.slice(1).split("}");
    return new NamespacePrefixedTag(`${pfxmap[nsuri]}:${localName}`);
  }

  get clarkName(): string {
    return `{${this._nsUri}}${this._localPart}`;
  }

  /** The local part of the tag, e.g. 'foobar' for tag 'f:foobar'. */
  get localPart(): string {
    return this._localPart;
  }

  /** A single-member mapping of this tag's prefix to its namespace URI. */
  get nsmap(): Record<string, string> {
    return { [this._pfx]: this._nsUri };
  }

  /** The namespace prefix, e.g. 'f' for tag 'f:foobar'. */
  get nspfx(): string {
    return this._pfx;
  }

  /** The namespace URI for the tag. */
  get nsuri(): string {
    return this._nsUri;
  }

  equals(other: unknown): boolean {
    return String(other) === this.valueOf();
  }
}

/** Mapping of the given prefixes to their namespace URIs, e.g. namespaces('a', 'r', 'p'). */
export function namespaces(...prefixes: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const p of prefixes) out[p] = nsLookup(p);
  return out;
}

export const nsmap = namespaces;

/** `xmlns:pfx="uri"` declarations for the given prefixes, space-separated. */
export function nsdecls(...prefixes: string[]): string {
  return prefixes.map((p) => `xmlns:${p}="${nsLookup(p)}"`).join(" ");
}

/** The namespace URI for `nspfx`. */
export function nsuri(nspfx: string): string {
  return nsLookup(nspfx);
}

/** Clark-notation qualified name for a prefixed tag like 'p:cSld'. */
export function qn(namespacePrefixedTag: string): string {
  let v = clarkCache.get(namespacePrefixedTag);
  if (v === undefined) {
    const i = namespacePrefixedTag.indexOf(":");
    if (i < 0) throw new Error(`ValueError: not enough values to unpack (expected 2, got 1)`);
    v = `{${nsLookup(namespacePrefixedTag.slice(0, i))}}${namespacePrefixedTag.slice(i + 1)}`;
    clarkCache.set(namespacePrefixedTag, v);
  }
  return v;
}
