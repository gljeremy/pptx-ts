// A small evaluator for the python expressions behave steps pass to `eval()`: literals
// (numbers, strings with escapes, True/False/None), tuples/lists, names from the step module,
// attribute access, calls, unary minus and + - * /. Strings come back as JS strings, python
// tuples and lists as arrays.
import { $call } from "../pyshim.js";

type Tok = { k: "num" | "str" | "name" | "op" | "end"; v: any };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    const num = /^(\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+(?:[eE][-+]?\d+)?)/.exec(src.slice(i));
    if (num) {
      out.push({ k: "num", v: Number(num[1]) });
      i += num[1].length;
      continue;
    }
    const str = /^([uUbBrR]?)(['"])/.exec(src.slice(i));
    if (str) {
      const raw = str[1].toLowerCase() === "r";
      const q = str[2];
      i += str[0].length;
      let s = "";
      while (src[i] !== q) {
        if (i >= src.length) throw new SyntaxError(`unterminated string in ${src}`);
        if (src[i] === "\\" && !raw) {
          const e = src[i + 1];
          const simple: Record<string, string> = { n: "\n", t: "\t", r: "\r", v: "\v", "\\": "\\", "'": "'", '"': '"', "0": "\0" };
          if (e in simple) {
            s += simple[e];
            i += 2;
          } else if (e === "x" || e === "u" || e === "U") {
            const n = { x: 2, u: 4, U: 8 }[e]!;
            s += String.fromCodePoint(parseInt(src.slice(i + 2, i + 2 + n), 16));
            i += 2 + n;
          } else {
            s += "\\";
            i += 1;
          }
        } else s += src[i++];
      }
      i++;
      out.push({ k: "str", v: s });
      continue;
    }
    const name = /^[A-Za-z_]\w*/.exec(src.slice(i));
    if (name) {
      out.push({ k: "name", v: name[0] });
      i += name[0].length;
      continue;
    }
    if ("()[],.+-*/".includes(c)) {
      out.push({ k: "op", v: c });
      i++;
      continue;
    }
    throw new SyntaxError(`unexpected ${JSON.stringify(c)} in ${src}`);
  }
  out.push({ k: "end", v: null });
  return out;
}

const camel = (s: string) => s.replace(/_+([a-z0-9])/g, (_m, c) => c.toUpperCase());

/** python `eval(src)` with the module-level names in `scope`. */
export function pyEval(src: string, scope: Record<string, any>): any {
  const toks = tokenize(String(src));
  let p = 0;
  const peek = () => toks[p];
  const isOp = (v: string) => toks[p].k === "op" && toks[p].v === v;
  const expect = (v: string) => {
    if (!isOp(v)) throw new SyntaxError(`expected ${v} in ${src}`);
    p++;
  };

  function seq(close: string): any[] {
    const items: any[] = [];
    while (!isOp(close)) {
      items.push(sum());
      if (!isOp(",")) break;
      p++;
    }
    expect(close);
    return items;
  }

  function atom(): any {
    const t = toks[p++];
    if (t.k === "num" || t.k === "str") {
      // adjacent string literals concatenate
      if (t.k === "str") {
        let s = t.v;
        while (peek().k === "str") s += toks[p++].v;
        return s;
      }
      return t.v;
    }
    if (t.k === "name") {
      if (t.v === "True") return true;
      if (t.v === "False") return false;
      if (t.v === "None") return null;
      if (!(t.v in scope)) throw new ReferenceError(`NameError: name '${t.v}' is not defined`);
      return scope[t.v];
    }
    if (t.k === "op" && t.v === "(") {
      if (isOp(")")) {
        p++;
        return [];
      }
      const first = sum();
      if (isOp(")")) {
        p++;
        return first;
      }
      expect(",");
      return [first, ...seq(")")];
    }
    if (t.k === "op" && t.v === "[") return seq("]");
    throw new SyntaxError(`unexpected token in ${src}`);
  }

  function postfix(): any {
    let v = atom();
    for (;;) {
      if (isOp(".")) {
        p++;
        const name = toks[p++].v;
        v = name in Object(v) ? v[name] : v[camel(name)];
      } else if (isOp("(")) {
        p++;
        v = $call(v, ...seq(")"));
      } else return v;
    }
  }

  function unary(): any {
    if (isOp("-")) {
      p++;
      return -unary();
    }
    if (isOp("+")) {
      p++;
      return +unary();
    }
    return postfix();
  }

  function product(): any {
    let v = unary();
    while (isOp("*") || isOp("/")) {
      const op = toks[p++].v;
      const r = unary();
      v = op === "*" ? v * r : v / r;
    }
    return v;
  }

  function sum(): any {
    let v = product();
    while (isOp("+") || isOp("-")) {
      const op = toks[p++].v;
      const r = product();
      v = op === "+" ? (typeof v === "string" ? v + r : Number(v) + Number(r)) : v - r;
    }
    return v;
  }

  // a bare top-level tuple: `1, 2`
  const items = [sum()];
  let tuple = false;
  while (isOp(",")) {
    p++;
    tuple = true;
    if (peek().k !== "end") items.push(sum());
  }
  if (peek().k !== "end") throw new SyntaxError(`unexpected trailing input in ${src}`);
  return tuple ? items : items[0];
}
