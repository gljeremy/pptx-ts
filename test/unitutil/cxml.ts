// Port of python-pptx tests/unitutil/cxml.py: a parser for the Compact XML Expression
// Language (CXEL), e.g. 'p:txBody/(a:bodyPr{wrap=none},a:p/a:r/a:t"foo")', which produces
// XML element trees for unit tests.
import { parseXml } from "../../src/oxml/parser.js";
import { _nsmap as nsmap } from "../../src/oxml/ns.js";

/** An oxml element parsed from the XML generated from `cxelStr`. */
export function element(cxelStr: string): any {
  return parseXml(xml(cxelStr));
}

/** The XML generated from `cxelStr`. */
export function xml(cxelStr: string): string {
  const p = new CxelParser(cxelStr);
  const root = p.parseRoot();
  return root.xml;
}

function nsdecls(...nspfxs: string[]): string {
  return nspfxs.map((p) => ` xmlns:${p}="${nsmap[p]}"`).join("");
}

class Element {
  _children: Element[] = [];
  isRoot = false;
  _indentStr = "";
  constructor(
    public _tagname: string,
    public _attrs: [string, string][],
    public _text: string,
  ) {}

  get localNspfxs(): string[] {
    const nspfx = (name: string, isElement = false): string | null => {
      const i = name.indexOf(":");
      if (i === -1) return isElement ? "" : null;
      return name.slice(0, i);
    };
    const out = [nspfx(this._tagname, true)!];
    for (const [name] of this._attrs) {
      const p = nspfx(name);
      if (p === null || out.includes(p)) continue;
      out.push(p);
    }
    return out;
  }

  get nspfxs(): string[] {
    const out = this.localNspfxs;
    for (const c of this._children) for (const p of c.nspfxs) if (!out.includes(p)) out.push(p);
    return out;
  }

  get xml(): string {
    return this._xml(0);
  }

  _xml(indent: number): string {
    this._indentStr = " ".repeat(indent);
    let s = this._startTag;
    for (const c of this._children) s += c._xml(indent + 2);
    return s + this._endTag;
  }

  get _startTag(): string {
    const decls = this.isRoot ? nsdecls(...this.nspfxs) : "";
    let tag = `${this._indentStr}<${this._tagname}${decls}`;
    for (const [n, v] of this._attrs) tag += ` ${n}="${v}"`;
    if (this._text) tag += `>${this._text}`;
    else if (this._children.length) tag += ">\n";
    else tag += "/>\n";
    return tag;
  }

  get _endTag(): string {
    if (this._text) return `</${this._tagname}>\n`;
    if (this._children.length) return `${this._indentStr}</${this._tagname}>\n`;
    return "";
  }
}

const ALPHA = /[A-Za-z]/;
const ALNUM = /[A-Za-z0-9]/;
const ATTR_NAME = /[A-Za-z:]/;
const ATTR_VAL = /[A-Za-z0-9 %\-./:_]/;

class CxelParser {
  i = 0;
  constructor(private s: string) {}

  fail(what: string): never {
    throw new Error(`ParseException: expected ${what} at char ${this.i} in ${JSON.stringify(this.s)}`);
  }

  word(re: RegExp, what: string): string {
    const start = this.i;
    while (this.i < this.s.length && re.test(this.s[this.i])) this.i++;
    if (this.i === start) this.fail(what);
    return this.s.slice(start, this.i);
  }

  peek(): string {
    return this.s[this.i];
  }

  eat(ch: string): boolean {
    if (this.s[this.i] === ch) {
      this.i++;
      return true;
    }
    return false;
  }

  element(): Element {
    const pfx = this.word(ALPHA, "namespace prefix");
    if (!this.eat(":")) this.fail("':'");
    const local = this.word(ALNUM, "local name");
    const attrs: [string, string][] = [];
    if (this.eat("{")) {
      do {
        const n = this.word(ATTR_NAME, "attribute name");
        if (!this.eat("=")) this.fail("'='");
        const v = this.word(ATTR_VAL, "attribute value");
        attrs.push([n, v]);
      } while (this.eat(","));
      if (!this.eat("}")) this.fail("'}'");
    }
    let text = "";
    if (this.peek() === '"') {
      this.i++;
      const start = this.i;
      while (this.i < this.s.length && this.s[this.i] !== '"') {
        if (this.s[this.i] === "\\") this.i++;
        this.i++;
      }
      text = this.s.slice(start, this.i);
      if (!this.eat('"')) this.fail("closing quote");
    }
    return new Element(`${pfx}:${local}`, attrs, text);
  }

  node(): Element {
    const el = this.element();
    if (this.eat("/")) for (const c of this.childNodeList()) el._children.push(c);
    return el;
  }

  childNodeList(): Element[] {
    if (this.eat("(")) {
      const out = [this.node()];
      while (this.eat(",")) out.push(this.node());
      if (!this.eat(")")) this.fail("')'");
      return out;
    }
    return [this.node()];
  }

  parseRoot(): Element {
    const root = this.node();
    if (this.i !== this.s.length) this.fail("end of text");
    root.isRoot = true;
    return root;
  }
}
