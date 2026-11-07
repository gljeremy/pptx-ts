// A small, fast XML tree with the parts of the lxml API python-pptx relies on:
// element class lookup by tag, the text/tail model, namespace declarations
// that serialize the way libxml2 does, and an XPath subset (see xpath.ts).
//
// Tags and attribute names use Clark notation: "{namespace-uri}local".

export const XML_NS = "http://www.w3.org/XML/1998/namespace";

export type NsDecl = [prefix: string | null, uri: string];

/** Python-protocol hooks shared with the test shim (`obj[k]`, `len(obj)`, `x in obj`). */
export const PY_GETITEM = Symbol.for("py.getitem");
export const PY_LEN = Symbol.for("py.len");
export const PY_CONTAINS = Symbol.for("py.contains");

export class XmlError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "XMLSyntaxError";
  }
}

/** Non-element children (comments, processing instructions) keep their place in the tree. */
export class XmlComment {
  tail: string | null = null;
  _parent: XmlElement | null = null;
  constructor(public text: string) {}
  get tag(): string {
    return "#comment";
  }
  getparent(): XmlElement | null {
    return this._parent;
  }
}

export class XmlPI {
  tail: string | null = null;
  _parent: XmlElement | null = null;
  constructor(
    public target: string,
    public text: string,
  ) {}
  get tag(): string {
    return "#pi";
  }
  getparent(): XmlElement | null {
    return this._parent;
  }
}

export type XmlNode = XmlElement | XmlComment | XmlPI;

export const isElement = (n: unknown): n is XmlElement => n instanceof XmlElement;

/** True if an ElementPath expression has more than one step ("/" outside `{uri}` parts). */
function isPath(tag: string): boolean {
  if (tag.charCodeAt(0) === 123 && tag.indexOf("/", tag.indexOf("}")) < 0) return false;
  return tag.replace(/\{[^}]*\}/g, "").includes("/");
}

function splitPath(tag: string): string[] {
  const out: string[] = [];
  let cur = "";
  let depth = 0;
  for (const ch of tag) {
    if (ch === "{") depth++;
    else if (ch === "}") depth--;
    if (ch === "/" && depth === 0) {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

// -- element class lookup ------------------------------------------------------------------------

type ElementClass = new () => XmlElement;
const classLookup = new Map<string, ElementClass>();
let defaultClass: ElementClass | null = null;

/** Register `cls` to be instantiated for elements with Clark-name `tag`. */
export function registerClass(tag: string, cls: ElementClass): void {
  classLookup.set(tag, cls);
}

export function setDefaultElementClass(cls: ElementClass): void {
  defaultClass = cls;
}

export function lookupClass(tag: string): ElementClass {
  return classLookup.get(tag) ?? defaultClass ?? XmlElement;
}

/** Split a Clark name into [uri | null, local]. */
export function splitClark(name: string): [string | null, string] {
  if (name.charCodeAt(0) !== 123 /* { */) return [null, name];
  const i = name.indexOf("}");
  return [name.slice(1, i), name.slice(i + 1)];
}

// -- the element ---------------------------------------------------------------------------------

export class XmlElement implements Iterable<XmlNode> {
  tag!: string;
  /** prefix this element was written with (null = default namespace or none) */
  _prefix!: string | null;
  /** namespace declarations made on this element */
  _nsDecls!: NsDecl[] | null;
  _attrs!: Map<string, string>;
  _children!: XmlNode[];
  _parent!: XmlElement | null;
  _text!: string | null;
  tail!: string | null;

  /** Create an element of the registered class for `tag` (lxml `makeelement`). */
  static make(tag: string, nsmap?: Record<string, string> | NsDecl[] | null, prefix?: string | null): XmlElement {
    const cls = lookupClass(tag);
    const el = new cls();
    el._init(tag, prefix ?? null);
    if (nsmap) {
      const decls: NsDecl[] = Array.isArray(nsmap)
        ? nsmap.slice()
        : Object.entries(nsmap).map(([p, u]) => [p === "" || p === "null" ? null : p, u] as NsDecl);
      if (decls.length) el._nsDecls = decls;
      if (prefix === undefined) {
        const [uri] = splitClark(tag);
        if (uri !== null) {
          const d = decls.find((x) => x[1] === uri);
          if (d) el._prefix = d[0];
        }
      }
    }
    return el;
  }

  _init(tag: string, prefix: string | null): void {
    this.tag = tag;
    this._prefix = prefix;
    this._nsDecls = null;
    this._attrs = new Map();
    this._children = [];
    this._parent = null;
    this._text = null;
    this.tail = null;
  }

  // -- text --------------------------------------------------------------------------------------
  get text(): string | null {
    return this._text;
  }
  set text(v: string | null) {
    this._text = v === null || v === undefined ? null : String(v);
  }

  // -- attributes --------------------------------------------------------------------------------
  /** Attribute value or `dflt` (lxml `get`). */
  get(name: string, dflt: string | null = null): string | null {
    const v = this._attrs.get(name);
    return v === undefined ? dflt : v;
  }

  set(name: string, value: string): void {
    this._attrs.set(name, String(value));
  }

  /** The live attribute map (lxml `attrib`). */
  get attrib(): Map<string, string> {
    return this._attrs;
  }

  keys(): string[] {
    return [...this._attrs.keys()];
  }

  values(): string[] {
    return [...this._attrs.values()];
  }

  items(): [string, string][] {
    return [...this._attrs.entries()];
  }

  // -- namespace info ----------------------------------------------------------------------------
  get prefix(): string | null {
    return this._prefix;
  }

  /** In-scope namespace mapping, prefix -> uri (`null` key for the default namespace). */
  get nsmap(): Record<string, string> {
    const out: Record<string, string> = {};
    const chain: XmlElement[] = [];
    for (let e: XmlElement | null = this; e; e = e._parent) chain.push(e);
    for (let i = chain.length - 1; i >= 0; i--) {
      const d = chain[i]._nsDecls;
      if (d) for (const [p, u] of d) out[p === null ? "null" : p] = u;
    }
    return out;
  }

  /** In-scope namespace declarations as [prefix, uri] pairs, nearest element first. */
  _inScopeDecls(): NsDecl[] {
    const seen = new Set<string | null>();
    const out: NsDecl[] = [];
    for (let e: XmlElement | null = this; e; e = e._parent) {
      const d = e._nsDecls;
      if (!d) continue;
      for (const decl of d) {
        if (seen.has(decl[0])) continue;
        seen.add(decl[0]);
        out.push(decl);
      }
    }
    return out;
  }

  _lookupNs(prefix: string | null): string | undefined {
    for (let e: XmlElement | null = this; e; e = e._parent) {
      const d = e._nsDecls;
      if (!d) continue;
      for (const [p, u] of d) if (p === prefix) return u;
    }
    if (prefix === "xml") return XML_NS;
    return undefined;
  }

  // -- tree navigation ---------------------------------------------------------------------------
  get length(): number {
    return this._children.length;
  }

  [Symbol.iterator](): Iterator<XmlNode> {
    return this._children.slice()[Symbol.iterator]();
  }

  [PY_GETITEM](i: number): XmlNode {
    const n = this._children.length;
    const j = i < 0 ? n + i : i;
    if (j < 0 || j >= n) throw new RangeError("list index out of range");
    return this._children[j];
  }

  [PY_LEN](): number {
    return this._children.length;
  }

  [PY_CONTAINS](x: unknown): boolean {
    return this._children.includes(x as XmlNode);
  }

  /** Child node at index `i` (negative counts from the end). */
  child(i: number): XmlNode {
    return this[PY_GETITEM](i);
  }

  getchildren(): XmlNode[] {
    return this._children.slice();
  }

  getparent(): XmlElement | null {
    return this._parent;
  }

  getnext(): XmlNode | null {
    const p = this._parent;
    if (!p) return null;
    const i = p._children.indexOf(this);
    return p._children[i + 1] ?? null;
  }

  getprevious(): XmlNode | null {
    const p = this._parent;
    if (!p) return null;
    const i = p._children.indexOf(this);
    return i > 0 ? p._children[i - 1] : null;
  }

  getroottree(): XmlElement {
    let e: XmlElement = this;
    while (e._parent) e = e._parent;
    return e;
  }

  index(child: XmlNode): number {
    const i = this._children.indexOf(child);
    if (i < 0) throw new Error("ValueError: element is not a child");
    return i;
  }

  /** First child element with Clark-name `tag` (lxml `find` for a plain tag / "*"). */
  find(tag: string): XmlElement | null {
    if (isPath(tag)) return (this.findall(tag)[0] as XmlElement) ?? null;
    const ch = this._children;
    for (let i = 0; i < ch.length; i++) {
      const c = ch[i];
      if (c instanceof XmlElement && (tag === "*" || c.tag === tag)) return c;
    }
    return null;
  }

  /** Child elements with Clark-name `tag`. Supports simple "a/b" paths. */
  findall(tag: string): XmlElement[] {
    if (isPath(tag)) {
      let cur: XmlElement[] = [this];
      for (const step of splitPath(tag)) {
        if (step === "." || step === "") continue;
        const next: XmlElement[] = [];
        for (const e of cur) next.push(...e.findall(step));
        cur = next;
      }
      return cur;
    }
    const out: XmlElement[] = [];
    const ch = this._children;
    for (let i = 0; i < ch.length; i++) {
      const c = ch[i];
      if (c instanceof XmlElement && (tag === "*" || c.tag === tag)) out.push(c);
    }
    return out;
  }

  findtext(tag: string, dflt: string | null = null): string | null {
    const e = this.find(tag);
    return e ? (e.text ?? "") : dflt;
  }

  /** Child elements, optionally only those with one of `tags`. */
  *iterchildren(...tags: string[]): Generator<XmlElement> {
    for (const c of this._children.slice()) {
      if (c instanceof XmlElement && (tags.length === 0 || tags.includes(c.tag))) yield c;
    }
  }

  /** Depth-first iteration over this element and its descendants (lxml `iter`). */
  *iter(...tags: string[]): Generator<XmlElement> {
    const all = tags.length === 0 || tags.includes("*");
    const stack: XmlElement[] = [this];
    while (stack.length) {
      const e = stack.pop()!;
      if (all || tags.includes(e.tag)) yield e;
      const ch = e._children;
      for (let i = ch.length - 1; i >= 0; i--) {
        const c = ch[i];
        if (c instanceof XmlElement) stack.push(c);
      }
    }
  }

  *iterdescendants(...tags: string[]): Generator<XmlElement> {
    for (const e of this.iter(...tags)) if (e !== this) yield e;
  }

  *iterancestors(...tags: string[]): Generator<XmlElement> {
    for (let e = this._parent; e; e = e._parent) if (tags.length === 0 || tags.includes(e.tag)) yield e;
  }

  *itersiblings(opts: { preceding?: boolean } = {}): Generator<XmlNode> {
    const p = this._parent;
    if (!p) return;
    const i = p._children.indexOf(this);
    if (opts.preceding) for (let j = i - 1; j >= 0; j--) yield p._children[j];
    else for (let j = i + 1; j < p._children.length; j++) yield p._children[j];
  }

  // -- tree mutation -----------------------------------------------------------------------------
  _adopt(node: XmlNode): void {
    const old = node._parent;
    if (old) {
      const i = old._children.indexOf(node);
      if (i >= 0) old._children.splice(i, 1);
    }
    node._parent = this;
    if (node instanceof XmlElement) node._stripRedundantNsDecls();
  }

  /**
   * Drop declarations whose namespace URI an ancestor already declares, possibly under another
   * prefix (libxml2's moveNodeToDocument does this, matching by href); the element then uses
   * the ancestor's prefix.
   */
  _stripRedundantNsDecls(): void {
    const d = this._nsDecls;
    if (!d || !this._parent) return;
    const kept: NsDecl[] = [];
    for (const [p, u] of d) {
      const other = this._parent._prefixForUri(u);
      if (other === undefined) kept.push([p, u]);
      else if (this._prefix === p) this._prefix = other;
    }
    this._nsDecls = kept.length ? kept : null;
  }

  /** In-scope prefix (null = default) bound to `uri`, nearest first; undefined if none. */
  _prefixForUri(uri: string): string | null | undefined {
    const shadowed = new Set<string | null>();
    for (let e: XmlElement | null = this; e; e = e._parent) {
      const d = e._nsDecls;
      if (!d) continue;
      for (const [p, u] of d) {
        if (shadowed.has(p)) continue;
        if (u === uri) return p;
        shadowed.add(p);
      }
    }
    return undefined;
  }

  append(node: XmlNode): void {
    if (node === (this as XmlNode)) throw new XmlError("cannot append parent to itself");
    this._adopt(node);
    this._children.push(node);
  }

  extend(nodes: Iterable<XmlNode>): void {
    for (const n of [...nodes]) this.append(n);
  }

  insert(index: number, node: XmlNode): void {
    this._adopt(node);
    const n = this._children.length;
    let i = index < 0 ? Math.max(0, n + index) : Math.min(index, n);
    this._children.splice(i, 0, node);
  }

  remove(node: XmlNode): void {
    const i = this._children.indexOf(node);
    if (i < 0) throw new Error("ValueError: Element is not a child of this node.");
    this._children.splice(i, 1);
    node._parent = null;
    node.tail = null;
  }

  replace(oldNode: XmlNode, newNode: XmlNode): void {
    const i = this._children.indexOf(oldNode);
    if (i < 0) throw new Error("ValueError: Element is not a child of this node.");
    this._adopt(newNode);
    const j = this._children.indexOf(oldNode);
    this._children[j] = newNode;
    oldNode._parent = null;
  }

  /** Insert `node` as the preceding sibling of this element. */
  addprevious(node: XmlNode): void {
    const p = this._parent;
    if (!p) throw new TypeError("Only processing instructions and comments can be siblings of the root element");
    p._adopt(node);
    p._children.splice(p._children.indexOf(this), 0, node);
  }

  /** Insert `node` as the following sibling of this element. */
  addnext(node: XmlNode): void {
    const p = this._parent;
    if (!p) throw new TypeError("Only processing instructions and comments can be siblings of the root element");
    p._adopt(node);
    p._children.splice(p._children.indexOf(this) + 1, 0, node);
  }

  /** Remove all children, attributes and text (lxml `clear`). */
  clear(): void {
    for (const c of this._children) c._parent = null;
    this._children = [];
    this._attrs = new Map();
    this._text = null;
    this.tail = null;
  }

  /** Create an element of the registered class for `tag` (lxml `makeelement`). */
  makeelement(tag: string, attrib?: Record<string, string> | null, nsmap?: Record<string, string> | null): XmlElement {
    const el = XmlElement.make(tag, nsmap);
    if (attrib) for (const [k, v] of Object.entries(attrib)) el.set(k, v);
    return el;
  }

  // -- copying -----------------------------------------------------------------------------------
  /** Deep copy, carrying in-scope namespace declarations onto the new root (like lxml). */
  deepcopy(): this {
    const root = this._copy() as this;
    const decls = this._inScopeDecls();
    if (decls.length) {
      const own = new Set((root._nsDecls ?? []).map((d) => d[0]));
      const extra = decls.filter((d) => !own.has(d[0]));
      if (extra.length) root._nsDecls = [...(root._nsDecls ?? []), ...extra];
    }
    root.tail = null;
    return root;
  }

  _copy(): XmlElement {
    const cls = this.constructor as ElementClass;
    const el = new cls();
    el._init(this.tag, this._prefix);
    el._nsDecls = this._nsDecls ? this._nsDecls.map((d) => [d[0], d[1]] as NsDecl) : null;
    el._attrs = new Map(this._attrs);
    el._text = this._text;
    el.tail = this.tail;
    for (const c of this._children) {
      let cc: XmlNode;
      if (c instanceof XmlElement) cc = c._copy();
      else if (c instanceof XmlComment) {
        cc = new XmlComment(c.text);
        cc.tail = c.tail;
      } else {
        cc = new XmlPI(c.target, c.text);
        cc.tail = c.tail;
      }
      cc._parent = el;
      el._children.push(cc);
    }
    return el;
  }
}

// -- parsing -------------------------------------------------------------------------------------

const ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

function decodeEntities(s: string, pos: number): string {
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z_][\w.-]*);/g, (m, ent: string) => {
    if (ent[0] === "#") {
      const cp = ent[1] === "x" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
      return String.fromCodePoint(cp);
    }
    const v = ENTITIES[ent];
    if (v === undefined) throw new XmlError(`Entity '${ent}' not defined, line ${lineOf(s, pos)}`);
    return v;
  });
}

function lineOf(_s: string, _pos: number): number {
  return 1;
}

const isWs = (c: number) => c === 32 || c === 10 || c === 13 || c === 9;

function isBlank(s: string): boolean {
  for (let i = 0; i < s.length; i++) if (!isWs(s.charCodeAt(i))) return false;
  return true;
}

export interface ParseOptions {
  /** Drop whitespace-only text between elements (lxml `remove_blank_text`). Default true. */
  removeBlankText?: boolean;
}

/** Parse an XML document (string or UTF-8 bytes) and return its root element. */
export function parseXml(xml: string | Uint8Array | ArrayBuffer, opts: ParseOptions = {}): XmlElement {
  let s: string;
  if (typeof xml === "string") s = xml;
  else s = decodeBytes(xml instanceof Uint8Array ? xml : new Uint8Array(xml));
  return new Parser(s, opts.removeBlankText ?? true).parse();
}

const utf8 = new TextDecoder("utf-8");
const utf16le = new TextDecoder("utf-16le");
const utf16be = new TextDecoder("utf-16be");

function decodeBytes(b: Uint8Array): string {
  if (b[0] === 0xff && b[1] === 0xfe) return utf16le.decode(b.subarray(2));
  if (b[0] === 0xfe && b[1] === 0xff) return utf16be.decode(b.subarray(2));
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return utf8.decode(b.subarray(3));
  // look at the declaration for a non-utf8 encoding
  if (b[0] === 0x3c && b[1] === 0x3f) {
    const head = String.fromCharCode(...b.subarray(0, Math.min(200, b.length)));
    const m = /encoding\s*=\s*["']([A-Za-z0-9._-]+)["']/.exec(head.slice(0, head.indexOf("?>") + 2 || undefined));
    if (m && !/^utf-?8$/i.test(m[1])) {
      try {
        return new TextDecoder(m[1]).decode(b);
      } catch {
        /* fall through */
      }
    }
  }
  return utf8.decode(b);
}

class Parser {
  i = 0;
  constructor(
    private s: string,
    private removeBlank: boolean,
  ) {}

  err(msg: string): never {
    throw new XmlError(`${msg} (at offset ${this.i})`);
  }

  parse(): XmlElement {
    const s = this.s;
    let root: XmlElement | null = null;
    // prolog
    while (this.i < s.length) {
      this.skipWs();
      if (s.startsWith("<?", this.i)) {
        const end = s.indexOf("?>", this.i);
        if (end < 0) this.err("unterminated processing instruction");
        this.i = end + 2;
      } else if (s.startsWith("<!--", this.i)) {
        const end = s.indexOf("-->", this.i);
        if (end < 0) this.err("unterminated comment");
        this.i = end + 3;
      } else if (s.startsWith("<!DOCTYPE", this.i)) {
        this.skipDoctype();
      } else if (s.charCodeAt(this.i) === 60) {
        root = this.element(null);
        break;
      } else if (this.i < s.length) {
        this.err("Start tag expected, '<' not found");
      }
    }
    if (!root) this.err("Document is empty");
    // trailing misc
    while (this.i < s.length) {
      this.skipWs();
      if (this.i >= s.length) break;
      if (s.startsWith("<!--", this.i)) this.i = s.indexOf("-->", this.i) + 3;
      else if (s.startsWith("<?", this.i)) this.i = s.indexOf("?>", this.i) + 2;
      else this.err("Extra content at the end of the document");
    }
    return root;
  }

  skipWs(): void {
    const s = this.s;
    while (this.i < s.length && isWs(s.charCodeAt(this.i))) this.i++;
  }

  skipDoctype(): void {
    const s = this.s;
    let depth = 0;
    for (; this.i < s.length; this.i++) {
      const c = s[this.i];
      if (c === "[") depth++;
      else if (c === "]") depth--;
      else if (c === ">" && depth === 0) {
        this.i++;
        return;
      }
    }
  }

  name(): string {
    const s = this.s;
    const start = this.i;
    while (this.i < s.length) {
      const c = s.charCodeAt(this.i);
      if (isWs(c) || c === 47 || c === 62 || c === 61) break; // / > =
      this.i++;
    }
    if (this.i === start) this.err("name expected");
    return s.slice(start, this.i);
  }

  /** Parse an element starting at '<'. */
  element(parent: XmlElement | null): XmlElement {
    const s = this.s;
    this.i++; // <
    const qname = this.name();
    // attributes (raw)
    let rawAttrs: string[] | null = null;
    let decls: NsDecl[] | null = null;
    let selfClosing = false;
    for (;;) {
      this.skipWs();
      const c = s.charCodeAt(this.i);
      if (c === 62) {
        this.i++;
        break;
      }
      if (c === 47) {
        if (s.charCodeAt(this.i + 1) !== 62) this.err("expected '>'");
        this.i += 2;
        selfClosing = true;
        break;
      }
      if (this.i >= s.length) this.err("Premature end of data");
      const an = this.name();
      this.skipWs();
      if (s.charCodeAt(this.i) !== 61) this.err("Specification mandates value for attribute " + an);
      this.i++;
      this.skipWs();
      const q = s[this.i];
      if (q !== '"' && q !== "'") this.err("AttValue: \" or ' expected");
      const end = s.indexOf(q, this.i + 1);
      if (end < 0) this.err("unterminated attribute value");
      let raw = s.slice(this.i + 1, end);
      this.i = end + 1;
      // attribute-value normalization: literal whitespace chars become spaces
      if (/[\n\r\t]/.test(raw)) raw = raw.replace(/\r\n|[\n\r\t]/g, " ");
      const v = decodeEntities(raw, this.i);
      if (an === "xmlns") (decls ??= []).push([null, v]);
      else if (an.startsWith("xmlns:")) (decls ??= []).push([an.slice(6), v]);
      else (rawAttrs ??= []).push(an, v);
    }
    // resolve the element name
    const colon = qname.indexOf(":");
    const prefix = colon >= 0 ? qname.slice(0, colon) : null;
    const local = colon >= 0 ? qname.slice(colon + 1) : qname;
    const lookup = (p: string | null): string | undefined => {
      if (decls) for (const [dp, du] of decls) if (dp === p) return du;
      return parent ? parent._lookupNs(p) : p === "xml" ? XML_NS : undefined;
    };
    const uri = lookup(prefix);
    if (prefix !== null && uri === undefined) this.err(`Namespace prefix ${prefix} on ${local} is not defined`);
    const tag = uri ? `{${uri}}${local}` : local;
    const cls = lookupClass(tag);
    const el = new cls();
    el._init(tag, prefix);
    el._nsDecls = decls;
    el._parent = parent;
    if (rawAttrs) {
      const attrs = el._attrs;
      for (let k = 0; k < rawAttrs.length; k += 2) {
        const an = rawAttrs[k];
        const ac = an.indexOf(":");
        if (ac < 0) attrs.set(an, rawAttrs[k + 1]);
        else {
          const ap = an.slice(0, ac);
          const au = ap === "xml" ? XML_NS : lookup(ap);
          if (au === undefined) this.err(`Namespace prefix ${ap} for ${an.slice(ac + 1)} on ${local} is not defined`);
          attrs.set(`{${au}}${an.slice(ac + 1)}`, rawAttrs[k + 1]);
        }
      }
    }
    if (selfClosing) return el;
    // content
    let pendingText: string | null = null; // text accumulated since the last node
    const flushText = (nextIsEnd: boolean) => {
      if (pendingText === null) return;
      let t = pendingText;
      pendingText = null;
      const ch = el._children;
      if (this.removeBlank && isBlank(t)) {
        // keep only when it is the whole content of the element (libxml2 keepBlanks=0)
        if (!(ch.length === 0 && nextIsEnd)) return;
      }
      if (ch.length === 0) el._text = (el._text ?? "") + t;
      else {
        const last = ch[ch.length - 1];
        last.tail = (last.tail ?? "") + t;
      }
    };
    for (;;) {
      const lt = s.indexOf("<", this.i);
      if (lt < 0) this.err(`Premature end of data in tag ${qname}`);
      if (lt > this.i) {
        const raw = s.slice(this.i, lt);
        const t = decodeEntities(raw.indexOf("\r") >= 0 ? raw.replace(/\r\n?/g, "\n") : raw, this.i);
        pendingText = (pendingText ?? "") + t;
      }
      this.i = lt;
      const c1 = s.charCodeAt(lt + 1);
      if (c1 === 47) {
        // end tag
        flushText(true);
        this.i = lt + 2;
        const endName = this.name();
        if (endName !== qname) this.err(`Opening and ending tag mismatch: ${qname} and ${endName}`);
        this.skipWs();
        if (s.charCodeAt(this.i) !== 62) this.err("expected '>'");
        this.i++;
        return el;
      }
      if (c1 === 33) {
        if (s.startsWith("<![CDATA[", lt)) {
          const end = s.indexOf("]]>", lt);
          if (end < 0) this.err("unterminated CDATA");
          // CDATA is never "blank text" to be removed
          const t = s.slice(lt + 9, end);
          if (pendingText !== null && this.removeBlank && isBlank(pendingText) && t.length) {
            // whitespace before CDATA belongs with it
          }
          pendingText = (pendingText ?? "") + t;
          const ch = el._children;
          // write immediately so blank-removal doesn't drop it
          const txt = pendingText;
          pendingText = null;
          if (ch.length === 0) el._text = (el._text ?? "") + txt;
          else ch[ch.length - 1].tail = (ch[ch.length - 1].tail ?? "") + txt;
          this.i = end + 3;
          continue;
        }
        if (s.startsWith("<!--", lt)) {
          flushText(false);
          const end = s.indexOf("-->", lt);
          if (end < 0) this.err("unterminated comment");
          const cm = new XmlComment(s.slice(lt + 4, end));
          cm._parent = el;
          el._children.push(cm);
          this.i = end + 3;
          continue;
        }
        this.err("unexpected markup");
      }
      if (c1 === 63) {
        flushText(false);
        const end = s.indexOf("?>", lt);
        if (end < 0) this.err("unterminated processing instruction");
        const body = s.slice(lt + 2, end);
        const sp = body.search(/\s/);
        const pi = new XmlPI(sp < 0 ? body : body.slice(0, sp), sp < 0 ? "" : body.slice(sp + 1).trimStart());
        pi._parent = el;
        el._children.push(pi);
        this.i = end + 2;
        continue;
      }
      // child element
      flushText(false);
      const child = this.element(el);
      el._children.push(child);
    }
  }
}

// -- serialization -------------------------------------------------------------------------------

export interface SerializeOptions {
  /** add `<?xml ...?>` header (default false) */
  xmlDeclaration?: boolean;
  /** encoding named in the declaration (default "UTF-8") */
  encoding?: string;
  /** `standalone='yes'|'no'` in the declaration; omitted when undefined/null */
  standalone?: boolean | null;
  prettyPrint?: boolean;
  /** escape non-ASCII characters as character references (lxml default ASCII encoding) */
  ascii?: boolean;
}

function escText(s: string): string {
  if (!/[&<>\r]/.test(s)) return s;
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
}

function escAttr(s: string): string {
  if (!/[&<>"\n\r\t]/.test(s)) return s;
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#13;")
    .replace(/\t/g, "&#9;");
}

function asciiEscape(s: string): string {
  return s.replace(/[^\x00-\x7f]/gu, (ch) => `&#${ch.codePointAt(0)};`);
}

/** Known prefixes to use when a namespace must be declared during serialization. */
let preferredPrefixes: Record<string, string> = {};
export function setPreferredPrefixes(map: Record<string, string>): void {
  preferredPrefixes = map;
}

class Scope {
  constructor(
    public decls: NsDecl[],
    public parent: Scope | null,
  ) {}
  lookup(prefix: string | null): string | undefined {
    for (let s: Scope | null = this; s; s = s.parent) {
      for (const [p, u] of s.decls) if (p === prefix) return u;
    }
    return prefix === "xml" ? XML_NS : undefined;
  }
  prefixFor(uri: string, allowDefault: boolean): string | null | undefined {
    if (uri === XML_NS) return "xml";
    const shadowed = new Set<string | null>();
    for (let s: Scope | null = this; s; s = s.parent) {
      for (const [p, u] of s.decls) {
        if (shadowed.has(p)) continue;
        if (u === uri && (p !== null || allowDefault)) return p;
        shadowed.add(p);
      }
    }
    return undefined;
  }
}

/** Serialize `el` (lxml `etree.tostring`). */
export function tostring(el: XmlElement, opts: SerializeOptions = {}): string {
  const out: string[] = [];
  if (opts.xmlDeclaration) {
    const enc = opts.encoding ?? "UTF-8";
    let decl = `<?xml version='1.0' encoding='${enc}'`;
    if (opts.standalone === true) decl += " standalone='yes'";
    else if (opts.standalone === false) decl += " standalone='no'";
    out.push(decl + "?>\n");
  }
  // A subtree root carries the namespaces in scope from its ancestors, in libxml2's order:
  // its own declarations, the namespace of its name, those of its attributes, then every
  // ancestor declaration (nearest first) whose prefix isn't declared yet.
  const rootDecls: NsDecl[] = el._nsDecls ? el._nsDecls.slice() : [];
  if (el._parent) {
    const own = new Set(rootDecls.map((d) => d[0]));
    const add = (p: string | null, u: string | undefined) => {
      if (u !== undefined && !own.has(p) && p !== "xml") rootDecls.push([p, u]), own.add(p);
    };
    if (el.tag.charCodeAt(0) === 123) add(el._prefix, el._lookupNs(el._prefix));
    for (const an of el._attrs.keys()) {
      if (an.charCodeAt(0) !== 123) continue;
      const au = splitClark(an)[0]!;
      const p = el._parent._inScopeDecls().find((d) => d[1] === au && d[0] !== null)?.[0];
      if (p) add(p, au);
    }
    for (const d of el._parent._inScopeDecls()) add(d[0], d[1]);
  }
  writeElement(el, rootDecls, null, out, opts.prettyPrint ? 0 : -1);
  if (opts.prettyPrint) out.push("\n");
  let s = out.join("");
  if (opts.ascii) s = asciiEscape(s);
  return s;
}

function writeElement(el: XmlElement, decls: NsDecl[], parentScope: Scope | null, out: string[], level: number): void {
  const scope = new Scope(decls.slice(), parentScope);
  const extraDecls: NsDecl[] = [];
  const declare = (prefix: string | null, uri: string) => {
    scope.decls.push([prefix, uri]);
    extraDecls.push([prefix, uri]);
  };
  // element name
  const [uri, local] = splitClark(el.tag);
  let name: string;
  if (uri === null) {
    const dflt = scope.lookup(null);
    if (dflt) declare(null, "");
    name = local;
  } else {
    let p: string | null | undefined = el._prefix;
    if (scope.lookup(p ?? null) !== uri) p = scope.prefixFor(uri, true);
    if (p === undefined) {
      p = el._prefix ?? preferredPrefixes[uri] ?? null;
      if (p !== null && scope.lookup(p) !== undefined) p = genPrefix(scope);
      declare(p, uri);
    }
    name = p === null ? local : `${p}:${local}`;
  }
  // attributes
  let attrStr = "";
  for (const [an, av] of el._attrs) {
    if (an.charCodeAt(0) === 123) {
      const [au, al] = splitClark(an);
      let p = scope.prefixFor(au!, false);
      if (p === undefined) {
        p = preferredPrefixes[au!] ?? null;
        if (p === null || scope.lookup(p) !== undefined) p = genPrefix(scope);
        declare(p, au!);
      }
      attrStr += ` ${p}:${al}="${escAttr(av)}"`;
    } else attrStr += ` ${an}="${escAttr(av)}"`;
  }
  let head = "<" + name;
  for (const [p, u] of [...decls, ...extraDecls]) head += p === null ? ` xmlns="${escAttr(u)}"` : ` xmlns:${p}="${escAttr(u)}"`;
  head += attrStr;
  const ch = el._children;
  const text = el._text;
  if (ch.length === 0 && text === null) {
    out.push(head + "/>");
    return;
  }
  out.push(head + ">");
  if (text !== null) out.push(escText(text));
  // libxml2 only indents children when the element has no text content
  const format = level >= 0 && text === null && ch.every((c) => c.tail === null || c.tail === "");
  for (const c of ch) {
    if (format) out.push("\n" + "  ".repeat(level + 1));
    if (c instanceof XmlElement) writeElement(c, c._nsDecls ? c._nsDecls.slice() : [], scope, out, format ? level + 1 : -1);
    else if (c instanceof XmlComment) out.push(`<!--${c.text}-->`);
    else out.push(`<?${c.target}${c.text ? " " + c.text : ""}?>`);
    if (c.tail) out.push(escText(c.tail));
  }
  if (format && ch.length) out.push("\n" + "  ".repeat(level));
  out.push(`</${name}>`);
}

function genPrefix(scope: Scope): string {
  for (let i = 0; ; i++) if (scope.lookup(`ns${i}`) === undefined) return `ns${i}`;
}

/** Encode `s` as UTF-8 bytes. */
export function encodeUtf8(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
