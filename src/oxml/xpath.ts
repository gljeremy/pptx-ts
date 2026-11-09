// The XPath 1.0 subset python-pptx uses: location paths with the child,
// descendant(-or-self), self, parent, ancestor(-or-self), following-sibling,
// preceding-sibling and attribute axes; name tests with prefixes, "*",
// node() and text(); predicates with positions, existence tests,
// comparisons, and/or/not(), count(), last() and position(); unions.
// Compiled expressions are cached.
import { XmlElement, XmlNode, splitClark } from "./xml.js";

export type XPathResult = Array<XmlElement | string> | number | boolean | string;

type Axis =
  | "child"
  | "descendant"
  | "descendant-or-self"
  | "self"
  | "parent"
  | "ancestor"
  | "ancestor-or-self"
  | "following-sibling"
  | "preceding-sibling"
  | "attribute";

interface Step {
  axis: Axis;
  /** "*" | "node()" | "text()" | prefix:local | local */
  test: string;
  preds: Expr[];
}

type Expr =
  | { k: "path"; abs: boolean; steps: Step[] }
  | { k: "union"; parts: Expr[] }
  /** a filter expression: `(expr)[pred]/step/...` or `fn()[pred]//step` */
  | { k: "filter"; base: Expr; preds: Expr[]; steps: Step[] }
  | { k: "lit"; v: string }
  | { k: "num"; v: number }
  | { k: "bin"; op: string; l: Expr; r: Expr }
  | { k: "fn"; name: string; args: Expr[] }
  | { k: "neg"; e: Expr };

// -- tokenizer -----------------------------------------------------------------------------------

type Tok = { t: string; v: string };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '"' || c === "'") {
      const j = src.indexOf(c, i + 1);
      if (j < 0) throw new Error(`XPathEvalError: unterminated literal in ${src}`);
      out.push({ t: "lit", v: src.slice(i + 1, j) });
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(src[i + 1] ?? ""))) {
      const m = /^[0-9]*\.?[0-9]+|^[0-9]+/.exec(src.slice(i))!;
      out.push({ t: "num", v: m[0] });
      i += m[0].length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (["//", "::", "!=", "<=", ">=", ".."].includes(two)) {
      out.push({ t: two, v: two });
      i += 2;
      continue;
    }
    if ("/[]()@|=<>,*.+-".includes(c)) {
      out.push({ t: c, v: c });
      i++;
      continue;
    }
    const m = /^[A-Za-z_][\w.-]*(?::(?:[A-Za-z_][\w.-]*|\*))?/.exec(src.slice(i));
    if (!m) throw new Error(`XPathEvalError: Invalid expression ${src}`);
    out.push({ t: "name", v: m[0] });
    i += m[0].length;
  }
  return out;
}

// -- parser --------------------------------------------------------------------------------------

class XParser {
  i = 0;
  constructor(
    private toks: Tok[],
    private src: string,
  ) {}
  peek(o = 0): Tok | undefined {
    return this.toks[this.i + o];
  }
  next(): Tok {
    const t = this.toks[this.i++];
    if (!t) throw new Error(`XPathEvalError: Invalid expression ${this.src}`);
    return t;
  }
  expect(t: string): void {
    const n = this.next();
    if (n.t !== t) throw new Error(`XPathEvalError: Invalid expression ${this.src}`);
  }
  parse(): Expr {
    const e = this.or();
    if (this.i < this.toks.length) throw new Error(`XPathEvalError: Invalid expression ${this.src}`);
    return e;
  }
  or(): Expr {
    let l = this.and();
    while (this.peek()?.t === "name" && this.peek()!.v === "or") {
      this.next();
      l = { k: "bin", op: "or", l, r: this.and() };
    }
    return l;
  }
  and(): Expr {
    let l = this.eq();
    while (this.peek()?.t === "name" && this.peek()!.v === "and") {
      this.next();
      l = { k: "bin", op: "and", l, r: this.eq() };
    }
    return l;
  }
  eq(): Expr {
    let l = this.rel();
    while (this.peek() && ["=", "!="].includes(this.peek()!.t)) {
      const op = this.next().t;
      l = { k: "bin", op, l, r: this.rel() };
    }
    return l;
  }
  rel(): Expr {
    let l = this.add();
    while (this.peek() && ["<", ">", "<=", ">="].includes(this.peek()!.t)) {
      const op = this.next().t;
      l = { k: "bin", op, l, r: this.add() };
    }
    return l;
  }
  add(): Expr {
    let l = this.unary();
    while (this.peek() && ["+", "-"].includes(this.peek()!.t)) {
      const op = this.next().t;
      l = { k: "bin", op, l, r: this.unary() };
    }
    return l;
  }
  unary(): Expr {
    if (this.peek()?.t === "-") {
      this.next();
      return { k: "neg", e: this.unary() };
    }
    return this.union();
  }
  union(): Expr {
    const first = this.primaryOrPath();
    if (this.peek()?.t !== "|") return first;
    const parts = [first];
    while (this.peek()?.t === "|") {
      this.next();
      parts.push(this.primaryOrPath());
    }
    return { k: "union", parts };
  }
  primaryOrPath(): Expr {
    const t = this.peek();
    if (!t) throw new Error(`XPathEvalError: Invalid expression ${this.src}`);
    if (t.t === "lit") {
      this.next();
      return { k: "lit", v: t.v };
    }
    if (t.t === "num") {
      this.next();
      return { k: "num", v: parseFloat(t.v) };
    }
    if (t.t === "(") {
      this.next();
      const e = this.or();
      this.expect(")");
      return this.filterTail(e);
    }
    if (t.t === "name" && this.peek(1)?.t === "(" && !["text", "node"].includes(t.v)) {
      this.next();
      this.next();
      const args: Expr[] = [];
      if (this.peek()?.t !== ")") {
        args.push(this.or());
        while (this.peek()?.t === ",") {
          this.next();
          args.push(this.or());
        }
      }
      this.expect(")");
      return this.filterTail({ k: "fn", name: t.v, args });
    }
    return this.path();
  }
  /** Predicates and a relative location path after a primary expression, if any. */
  filterTail(base: Expr): Expr {
    const preds: Expr[] = [];
    while (this.peek()?.t === "[") {
      this.next();
      preds.push(this.or());
      this.expect("]");
    }
    const steps: Step[] = [];
    for (;;) {
      const n = this.peek();
      if (n?.t === "/") {
        this.next();
        steps.push(this.step());
      } else if (n?.t === "//") {
        this.next();
        steps.push({ axis: "descendant-or-self", test: "node()", preds: [] });
        steps.push(this.step());
      } else break;
    }
    return preds.length || steps.length ? { k: "filter", base, preds, steps } : base;
  }
  path(): Expr {
    const steps: Step[] = [];
    let abs = false;
    const t = this.peek();
    if (t?.t === "/") {
      this.next();
      abs = true;
      if (!this.startsStep()) return { k: "path", abs, steps };
    } else if (t?.t === "//") {
      this.next();
      abs = true;
      steps.push({ axis: "descendant-or-self", test: "node()", preds: [] });
    }
    steps.push(this.step());
    for (;;) {
      const n = this.peek();
      if (n?.t === "/") {
        this.next();
        steps.push(this.step());
      } else if (n?.t === "//") {
        this.next();
        steps.push({ axis: "descendant-or-self", test: "node()", preds: [] });
        steps.push(this.step());
      } else break;
    }
    return { k: "path", abs, steps };
  }
  startsStep(): boolean {
    const t = this.peek();
    return !!t && ["name", "*", "@", ".", ".."].includes(t.t);
  }
  step(): Step {
    const t = this.next();
    if (t.t === ".") return { axis: "self", test: "node()", preds: [] };
    if (t.t === "..") return { axis: "parent", test: "node()", preds: [] };
    let axis: Axis = "child";
    let test: string;
    if (t.t === "@") {
      axis = "attribute";
      const n = this.next();
      test = n.t === "*" ? "*" : n.v;
    } else if (t.t === "*") test = "*";
    else if (t.t === "name") {
      if (this.peek()?.t === "::") {
        this.next();
        axis = t.v as Axis;
        const n = this.next();
        test = n.t === "*" ? "*" : n.v;
      } else test = t.v;
      if ((test === "text" || test === "node") && this.peek()?.t === "(") {
        this.next();
        this.expect(")");
        test += "()";
      }
    } else throw new Error(`XPathEvalError: Invalid expression ${this.src}`);
    const preds: Expr[] = [];
    while (this.peek()?.t === "[") {
      this.next();
      preds.push(this.or());
      this.expect("]");
    }
    return { axis, test, preds };
  }
}

const cache = new Map<string, Expr>();

function compile(src: string): Expr {
  let e = cache.get(src);
  if (!e) {
    e = new XParser(tokenize(src), src).parse();
    cache.set(src, e);
  }
  return e;
}

// -- evaluation ----------------------------------------------------------------------------------

/** A text() or attribute result remembers where it came from (like lxml smart strings). */
type Item = XmlElement | AttrItem | TextItem;
class AttrItem {
  constructor(
    public owner: XmlElement,
    public name: string,
    public value: string,
  ) {}
}
class TextItem {
  constructor(
    public owner: XmlNode,
    public value: string,
    public isTail: boolean,
  ) {}
}

interface Ctx {
  ns: Record<string, string>;
}

function nameMatches(test: string, tag: string, ctx: Ctx, isAttr: boolean): boolean {
  if (test === "*" || test === "node()") return true;
  const colon = test.indexOf(":");
  if (colon < 0) return isAttr ? tag === test : tag === test; // no-namespace name
  const pfx = test.slice(0, colon);
  const local = test.slice(colon + 1);
  const uri = ctx.ns[pfx];
  if (uri === undefined) throw new Error(`XPathEvalError: Undefined namespace prefix ${pfx}`);
  const [tu, tl] = splitClark(tag);
  return tu === uri && (local === "*" || tl === local);
}

function textItems(e: XmlElement): TextItem[] {
  const out: TextItem[] = [];
  if (e.text !== null && e.text !== "") out.push(new TextItem(e, e.text, false));
  for (const c of e._children) if (c.tail) out.push(new TextItem(c, c.tail, true));
  return out;
}

function axisNodes(n: Item, axis: Axis): Item[] {
  if (!(n instanceof XmlElement)) {
    if (axis === "self") return [n];
    if (axis === "parent") return n instanceof AttrItem ? [n.owner] : n.isTail ? [n.owner.getparent()!].filter(Boolean) : [n.owner as XmlElement];
    if (axis === "ancestor" || axis === "ancestor-or-self") {
      const p = axisNodes(n, "parent")[0] as XmlElement | undefined;
      const out: Item[] = axis === "ancestor-or-self" ? [n] : [];
      for (let e = p ?? null; e; e = e._parent) out.push(e);
      return out;
    }
    return [];
  }
  switch (axis) {
    case "child":
      return n._children.filter((c): c is XmlElement => c instanceof XmlElement);
    case "self":
      return [n];
    case "parent":
      return n._parent ? [n._parent] : [];
    case "attribute":
      return [...n._attrs].map(([k, v]) => new AttrItem(n, k, v));
    case "descendant":
    case "descendant-or-self": {
      const out: Item[] = [];
      const walk = (e: XmlElement) => {
        out.push(e);
        for (const c of e._children) if (c instanceof XmlElement) walk(c);
      };
      if (axis === "descendant-or-self") walk(n);
      else for (const c of n._children) if (c instanceof XmlElement) walk(c);
      return out;
    }
    case "ancestor":
    case "ancestor-or-self": {
      const out: Item[] = axis === "ancestor-or-self" ? [n] : [];
      for (let e = n._parent; e; e = e._parent) out.push(e);
      return out;
    }
    case "following-sibling":
    case "preceding-sibling": {
      const p = n._parent;
      if (!p) return [];
      const sibs = p._children.filter((c): c is XmlElement => c instanceof XmlElement);
      const i = sibs.indexOf(n);
      return axis === "following-sibling" ? sibs.slice(i + 1) : sibs.slice(0, i).reverse();
    }
  }
}

const REVERSE_AXES = new Set<Axis>(["ancestor", "ancestor-or-self", "preceding-sibling"]);

function evalStep(nodes: Item[], step: Step, ctx: Ctx): Item[] {
  const out: Item[] = [];
  const seen = new Set<unknown>();
  for (const n of nodes) {
    let cand: Item[];
    if (step.test === "text()") {
      if (step.axis === "child") cand = n instanceof XmlElement ? textItems(n) : [];
      else if (step.axis === "descendant-or-self" || step.axis === "descendant")
        cand = axisNodes(n, step.axis).flatMap((e) => (e instanceof XmlElement ? textItems(e) : []));
      else cand = [];
    } else {
      cand = axisNodes(n, step.axis);
      if (step.test !== "node()") {
        const isAttr = step.axis === "attribute";
        cand = cand.filter((c) =>
          isAttr ? c instanceof AttrItem && nameMatches(step.test, c.name, ctx, true) : c instanceof XmlElement && nameMatches(step.test, c.tag, ctx, false),
        );
      }
    }
    for (const p of step.preds) {
      const size = cand.length;
      cand = cand.filter((c, i) => {
        const v = evalExpr(p, c, ctx, i + 1, size);
        if (typeof v === "number") return v === i + 1;
        return toBool(v);
      });
    }
    for (const c of cand) {
      const key = c instanceof AttrItem ? `${idOf(c.owner)}@${c.name}` : c instanceof TextItem ? c : c;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(c);
    }
  }
  return out;
}

const ids = new WeakMap<object, number>();
let nextId = 0;
function idOf(o: object): number {
  let v = ids.get(o);
  if (v === undefined) ids.set(o, (v = nextId++));
  return v;
}

type Val = Item[] | number | boolean | string;

function evalExpr(e: Expr, node: Item, ctx: Ctx, pos: number, size: number): Val {
  switch (e.k) {
    case "lit":
      return e.v;
    case "num":
      return e.v;
    case "neg":
      return -toNum(evalExpr(e.e, node, ctx, pos, size));
    case "path": {
      let cur: Item[];
      if (e.abs) {
        const root = node instanceof XmlElement ? node.getroottree() : (axisNodes(node, "ancestor").pop() as XmlElement);
        // the document node: "/x" selects the root element if it matches x
        cur = [];
        if (e.steps.length === 0) return [root];
        const [first, ...rest] = e.steps;
        if (first.axis === "child") {
          cur = evalStep([root], { ...first, axis: "self" }, ctx);
        } else if (first.axis === "descendant-or-self" && first.test === "node()") {
          cur = axisNodes(root, "descendant-or-self");
        } else cur = evalStep([root], first, ctx);
        for (const st of rest) cur = evalStep(cur, st, ctx);
        return cur;
      }
      cur = [node];
      for (const st of e.steps) cur = evalStep(cur, st, ctx);
      return cur;
    }
    case "union": {
      const all: Item[] = [];
      for (const p of e.parts) {
        const v = evalExpr(p, node, ctx, pos, size);
        if (!Array.isArray(v)) throw new Error("XPathEvalError: union of non-node-sets");
        for (const x of v) if (!all.includes(x)) all.push(x);
      }
      return sortDocOrder(all);
    }
    case "fn":
      return callFn(e.name, e.args, node, ctx, pos, size);
    case "filter": {
      const v = evalExpr(e.base, node, ctx, pos, size);
      if (!Array.isArray(v)) throw new Error("XPathEvalError: filter of a non-node-set");
      // predicates count positions in document order (the base is a node-set)
      let cur = sortDocOrder(v);
      for (const p of e.preds) {
        const n = cur.length;
        cur = cur.filter((c, i) => {
          const pv = evalExpr(p, c, ctx, i + 1, n);
          return typeof pv === "number" ? pv === i + 1 : toBool(pv);
        });
      }
      for (const st of e.steps) cur = evalStep(cur, st, ctx);
      return e.steps.length ? sortDocOrder(cur) : cur;
    }
    case "bin": {
      if (e.op === "or") return toBool(evalExpr(e.l, node, ctx, pos, size)) || toBool(evalExpr(e.r, node, ctx, pos, size));
      if (e.op === "and") return toBool(evalExpr(e.l, node, ctx, pos, size)) && toBool(evalExpr(e.r, node, ctx, pos, size));
      const l = evalExpr(e.l, node, ctx, pos, size);
      const r = evalExpr(e.r, node, ctx, pos, size);
      if (e.op === "+") return toNum(l) + toNum(r);
      if (e.op === "-") return toNum(l) - toNum(r);
      return compare(e.op, l, r);
    }
  }
}

function strVal(i: Item): string {
  if (i instanceof AttrItem || i instanceof TextItem) return i.value;
  // string-value of an element: all descendant text in document order
  const parts: string[] = [];
  const walk = (e: XmlElement) => {
    if (e.text) parts.push(e.text);
    for (const c of e._children) {
      if (c instanceof XmlElement) walk(c);
      if (c.tail) parts.push(c.tail);
    }
  };
  walk(i);
  return parts.join("");
}

function toBool(v: Val): boolean {
  if (Array.isArray(v)) return v.length > 0;
  if (typeof v === "number") return v !== 0 && !Number.isNaN(v);
  if (typeof v === "string") return v.length > 0;
  return v;
}

function toNum(v: Val): number {
  if (Array.isArray(v)) return v.length ? Number(strVal(v[0]).trim() || NaN) : NaN;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") return v.trim() === "" ? NaN : Number(v);
  return v;
}

function toStr(v: Val): string {
  if (Array.isArray(v)) return v.length ? strVal(v[0]) : "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  if (typeof v === "boolean") return v ? "true" : "false";
  return v;
}

function cmpAtom(op: string, a: string | number | boolean, b: string | number | boolean): boolean {
  if (op === "=" || op === "!=") {
    let eq: boolean;
    if (typeof a === "boolean" || typeof b === "boolean") eq = toBool(a as Val) === toBool(b as Val);
    else if (typeof a === "number" || typeof b === "number") eq = toNum(a as Val) === toNum(b as Val);
    else eq = a === b;
    return op === "=" ? eq : !eq;
  }
  const x = toNum(a as Val);
  const y = toNum(b as Val);
  return op === "<" ? x < y : op === ">" ? x > y : op === "<=" ? x <= y : x >= y;
}

function compare(op: string, l: Val, r: Val): boolean {
  if (Array.isArray(l) && Array.isArray(r)) return l.some((a) => r.some((b) => cmpAtom(op, strVal(a), strVal(b))));
  if (Array.isArray(l)) return l.some((a) => cmpAtom(op, typeof r === "number" ? toNum(strVal(a)) : strVal(a), r as any));
  if (Array.isArray(r)) return r.some((b) => cmpAtom(op, l as any, typeof l === "number" ? toNum(strVal(b)) : strVal(b)));
  return cmpAtom(op, l, r);
}

function callFn(name: string, args: Expr[], node: Item, ctx: Ctx, pos: number, size: number): Val {
  const a = (i: number) => evalExpr(args[i], node, ctx, pos, size);
  switch (name) {
    case "count": {
      const v = a(0);
      if (!Array.isArray(v)) throw new Error("XPathEvalError: count() of non-node-set");
      return v.length;
    }
    case "not":
      return !toBool(a(0));
    case "true":
      return true;
    case "false":
      return false;
    case "last":
      return size;
    case "position":
      return pos;
    case "string":
      return args.length ? toStr(a(0)) : strVal(node);
    case "number":
      return args.length ? toNum(a(0)) : toNum([node]);
    case "boolean":
      return toBool(a(0));
    case "concat":
      return args.map((_, i) => toStr(a(i))).join("");
    case "contains":
      return toStr(a(0)).includes(toStr(a(1)));
    case "starts-with":
      return toStr(a(0)).startsWith(toStr(a(1)));
    case "string-length":
      return (args.length ? toStr(a(0)) : strVal(node)).length;
    case "normalize-space":
      return (args.length ? toStr(a(0)) : strVal(node)).trim().replace(/\s+/g, " ");
    case "local-name": {
      const v = args.length ? (a(0) as Item[])[0] : node;
      if (!v) return "";
      return v instanceof XmlElement ? splitClark(v.tag)[1] : v instanceof AttrItem ? splitClark(v.name)[1] : "";
    }
    case "name": {
      const v = args.length ? (a(0) as Item[])[0] : node;
      return v instanceof XmlElement ? (v.prefix ? `${v.prefix}:${splitClark(v.tag)[1]}` : splitClark(v.tag)[1]) : "";
    }
    case "sum": {
      const v = a(0) as Item[];
      return v.reduce((s, x) => s + toNum(strVal(x)), 0);
    }
    default:
      throw new Error(`XPathEvalError: Unregistered function ${name}`);
  }
}

function docPath(n: Item): number[] {
  const el = n instanceof XmlElement ? n : n instanceof AttrItem ? n.owner : (n.owner as XmlElement);
  const path: number[] = [];
  for (let e: XmlNode = el; e.getparent(); e = e.getparent()!) path.unshift(e.getparent()!._children.indexOf(e));
  if (n instanceof AttrItem) path.push(-1, [...n.owner._attrs.keys()].indexOf(n.name));
  return path;
}

function sortDocOrder(items: Item[]): Item[] {
  const keyed = items.map((i) => [docPath(i), i] as const);
  keyed.sort((x, y) => {
    const a = x[0];
    const b = y[0];
    for (let i = 0; i < Math.min(a.length, b.length); i++) if (a[i] !== b[i]) return a[i] - b[i];
    return a.length - b.length;
  });
  return keyed.map((k) => k[1]);
}

/**
 * Evaluate `expr` with `node` as the context node (lxml `_Element.xpath`).
 * Node-set results are arrays of elements and strings (attribute values, text nodes).
 */
export function xpath(node: XmlElement, expr: string, namespaces: Record<string, string>): XPathResult {
  const v = evalExpr(compile(expr), node, { ns: namespaces }, 1, 1);
  if (!Array.isArray(v)) return v;
  return v.map((i) => (i instanceof XmlElement ? i : i.value));
}
