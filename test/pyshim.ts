// Runtime helpers giving Python semantics to tests machine-translated from the python-pptx
// pytest suite (scripts/py2ts_tests.py): indexing, `in`, `==`, `len`, truthiness, builtins,
// and a small pytest runtime (fixtures with params/finalizers, parametrize, raises).
import { describe, expect, it } from "vitest";
import { ElementProxy } from "../src/shared.js";
import { XmlElement } from "../src/oxml/xml.js";
import { Length, pyFloatStr, pyReprStr } from "../src/util.js";
import { createHash } from "node:crypto";
import nodePath from "node:path";
import { unzipSync } from "fflate";
import { BytesIO } from "../src/fileio.js";
import * as mockMod from "./mock.js";
import { ANY, _Call, camel, CallList, isMock, KW, setModules, setPyEq, setShow, signatureOf, type FixtureRequest } from "./mock.js";

import * as utilMod from "../src/util.js";
export * from "./mock.js";
export { BytesIO };
export { pyEval as $pyEval } from "./features/pyeval.js";

const GETITEM = Symbol.for("py.getitem");
const LEN = Symbol.for("py.len");
const CONTAINS = Symbol.for("py.contains");

setModules(import.meta.glob("../src/**/*.ts", { eager: true }));

const isTyped = (v: any) => ArrayBuffer.isView(v) && !(v instanceof DataView);

/**
 * A python dict literal: its items are plain own properties (so `d[k]`, `k in d` and
 * Object.entries work) and it also answers the Map-like API the port's source uses for
 * python mappings (`get` raising KeyError, `has`, `set`, `keys`, `values`, `items`).
 */
export class PyDict {
  get(k: any): any {
    const key = String(k);
    if (!Object.prototype.hasOwnProperty.call(this, key)) throw new KeyErrorShim(k);
    return (this as any)[key];
  }
  has(k: any): boolean {
    return (typeof k === "string" || k instanceof String || typeof k === "number") && Object.prototype.hasOwnProperty.call(this, String(k));
  }
  set(k: any, v: any): void {
    (this as any)[String(k)] = v;
  }
  delete(k: any): void {
    delete (this as any)[String(k)];
  }
  keys(): string[] {
    return Object.keys(this);
  }
  values(): any[] {
    return Object.values(this);
  }
  items(): [string, any][] {
    return Object.entries(this);
  }
  get length(): number {
    return Object.keys(this).length;
  }
  [Symbol.iterator](): Iterator<string> {
    return Object.keys(this)[Symbol.iterator]();
  }
}
for (const k of Object.getOwnPropertyNames(PyDict.prototype)) Object.defineProperty(PyDict.prototype, k, { enumerable: false });

/** dict literal */
export function $d(o: Record<string, any>): any {
  return Object.assign(new PyDict(), o);
}

const isPlain = (v: any) =>
  v !== null && typeof v === "object" && (Object.getPrototypeOf(v) === Object.prototype || Object.getPrototypeOf(v) === null || v instanceof PyDict);

// ------------------------------------------------------------------------------------------------
// builtins

/** obj[key] */
export function $get(o: any, k: any): any {
  if (o === null || o === undefined) throw new TypeError("'NoneType' object is not subscriptable");
  if (typeof o[GETITEM] === "function") return o[GETITEM](k);
  if (Array.isArray(o) || typeof o === "string" || o instanceof String || isTyped(o)) {
    const n = o.length;
    if (typeof k !== "number") throw new TypeError(`list indices must be integers or slices, not ${typeof k}`);
    const j = k < 0 ? n + k : k;
    if (j < 0 || j >= n) throw new RangeError("IndexError: list index out of range");
    return typeof o === "string" || o instanceof String ? String(o)[j] : o[j];
  }
  if (o instanceof Map) {
    if (!o.has(k)) throw new KeyErrorShim(k);
    return o.get(k);
  }
  if (typeof o.get === "function" && !(o instanceof XmlElement)) return o.get(k);
  const key = typeof k === "object" && k !== null ? String(k) : k;
  if (!(key in o)) throw new KeyErrorShim(k);
  return o[key];
}

class KeyErrorShim extends Error {
  constructor(k: any) {
    super(`KeyError: ${String(k)}`);
    this.name = "KeyError";
  }
}

/** obj[key] = v */
export function $set(o: any, k: any, v: any): void {
  if (o instanceof Map) o.set(k, v);
  else if (Array.isArray(o)) o[k < 0 ? o.length + k : k] = v;
  else if (typeof o[Symbol.for("py.setitem")] === "function") o[Symbol.for("py.setitem")](k, v);
  else if (typeof o.set === "function" && !(o instanceof XmlElement) && !isPlain(o)) o.set(k, v);
  else o[typeof k === "object" && k !== null ? String(k) : k] = v;
}

/** del obj[key] */
export function $del(o: any, k: any): void {
  if (o instanceof Map) o.delete(k);
  else if (Array.isArray(o)) o.splice(k < 0 ? o.length + k : k, 1);
  else if (typeof o[Symbol.for("py.delitem")] === "function") o[Symbol.for("py.delitem")](k);
  else if (o instanceof XmlElement && typeof k === "number") o.remove(o.child(k));
  else delete o[String(k)];
}

/** key in obj */
export function $in(k: any, o: any): boolean {
  if (o === null || o === undefined) throw new TypeError("argument of type 'NoneType' is not iterable");
  if (typeof o[CONTAINS] === "function") return o[CONTAINS](k);
  if (typeof o === "string" || o instanceof String) return String(o).includes(String(k));
  if (o instanceof Set || o instanceof Map) {
    if (o.has(k)) return true;
    if (k instanceof String) return o.has(String(k));
    return [...o.keys()].some((x) => $eq(x, k));
  }
  if (Array.isArray(o) || isTyped(o)) return Array.from(o as any).some((e) => $eq(e, k));
  if (typeof o.has === "function") return o.has(k);
  if (typeof o[Symbol.iterator] === "function") return [...o].some((e) => $eq(e, k));
  return Object.prototype.hasOwnProperty.call(o, String(k));
}

export function $len(o: any): number {
  if (o === null || o === undefined) throw new TypeError("object of type 'NoneType' has no len()");
  if (typeof o[LEN] === "function") return o[LEN]();
  if (typeof o === "string" || o instanceof String) return [...String(o)].length;
  if (typeof o.length === "number") return o.length;
  if (typeof o.size === "number") return o.size;
  if (isPlain(o)) return Object.keys(o).length;
  if (typeof o[Symbol.iterator] === "function") return [...o].length;
  throw new TypeError(`object of type '${o?.constructor?.name}' has no len()`);
}

const unwrap = (v: any) => (v instanceof String || v instanceof Number || v instanceof Boolean ? v.valueOf() : v);

/** Python == */
export function $eq(a: any, b: any): boolean {
  if (a === b) return true;
  if (a === ANY || b === ANY) return true;
  if (a === undefined) a = null;
  if (b === undefined) b = null;
  if (a === null || b === null) return a === b;
  if (isMock(a) || isMock(b)) return a === b;
  if (typeof a === "object" && typeof a.equals === "function") return a.equals(b);
  if (typeof b === "object" && typeof b.equals === "function") return b.equals(a);
  a = unwrap(a);
  b = unwrap(b);
  if (typeof a !== "object" || typeof b !== "object") {
    if (typeof a === "boolean" && typeof b === "number") return (a ? 1 : 0) === b;
    if (typeof b === "boolean" && typeof a === "number") return (b ? 1 : 0) === a;
    return a === b;
  }
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (a instanceof Uint8Array && b instanceof Uint8Array) return a.length === b.length && a.every((x, i) => x === b[i]);
  if ((Array.isArray(a) || isTyped(a)) && (Array.isArray(b) || isTyped(b))) {
    const aa = Array.from(a as any);
    const bb = Array.from(b as any);
    return aa.length === bb.length && aa.every((e, i) => $eq(e, bb[i]));
  }
  if (a instanceof Set && b instanceof Set) return a.size === b.size && [...a].every((e) => [...b].some((x) => $eq(e, x)));
  if (a instanceof Map && b instanceof Map) return a.size === b.size && [...a].every(([k, v]) => b.has(k) && $eq(v, b.get(k)));
  if ((isPlain(a) || a instanceof Map) && (isPlain(b) || b instanceof Map)) {
    const ea = a instanceof Map ? [...a.entries()] : Object.entries(a);
    const eb = b instanceof Map ? new Map(b) : new Map(Object.entries(b));
    return ea.length === eb.size && ea.every(([k, v]) => eb.has(k) && $eq(v, eb.get(k)));
  }
  if (isPlain(a) && typeof b === "object" && typeof b.items === "function") return $eq(a, Object.fromEntries(b.items()));
  if (isPlain(b) && typeof a === "object" && typeof a.items === "function") return $eq(Object.fromEntries(a.items()), b);
  return false;
}
setPyEq($eq);

/** Python `is` (None and undefined are the same thing here) */
export function $is(a: any, b: any): boolean {
  if ((a === null || a === undefined) && (b === null || b === undefined)) return true;
  return a === b;
}

/** Python truthiness */
export function $truthy(v: any): boolean {
  if (v === null || v === undefined) return false;
  if (isMock(v)) return true;
  if (typeof v === "string" || v instanceof String) return String(v).length > 0;
  if (typeof v === "number" || v instanceof Number) return Number(v) !== 0;
  if (typeof v === "boolean" || v instanceof Boolean) return Boolean(v.valueOf());
  if (typeof v[LEN] === "function" && !(v instanceof XmlElement)) return v[LEN]() > 0;
  if (Array.isArray(v) || isTyped(v)) return v.length > 0;
  if (v instanceof Map || v instanceof Set) return v.size > 0;
  if (isPlain(v)) return Object.keys(v).length > 0;
  return true;
}

export const show = (v: any): string => {
  try {
    if (v instanceof XmlElement) return String(v);
    if (typeof v === "string" || v instanceof String) return pyReprStr(String(v));
    return JSON.stringify(v, (_k, x) => (x instanceof Uint8Array ? `b'${new TextDecoder().decode(x).slice(0, 200)}'` : isMock(x) ? `<Mock ${x._mockName}>` : x instanceof _Call ? String(x) : x)) ?? String(v);
  } catch {
    return String(v);
  }
};
setShow(show);

export function $assertEqual(a: any, b: any, msg?: string): void {
  if (!$eq(a, b)) {
    if ((typeof a === "string" || a instanceof String) && (typeof b === "string" || b instanceof String)) expect(String(a), msg).toBe(String(b));
    if (a instanceof Uint8Array && b instanceof Uint8Array) expect(new TextDecoder().decode(a), msg).toBe(new TextDecoder().decode(b));
    if (a instanceof CallList || b instanceof CallList) expect.fail(msg ?? `${String(a)}\n    !=\n${String(b)}`);
    expect.fail(msg ?? `${show(a)} != ${show(b)}`);
  }
}

export function $assert(v: any, msg?: any): void {
  if (!$truthy(v)) expect.fail(msg !== undefined ? String(msg) : `assertion failed (got ${show(v)})`);
}

export interface ExcInfo {
  value: any;
  type: any;
}

/** pytest.raises(cls, match=...); `cls` may be an array of classes or null for any exception */
export function $raises(cls: any, fn: () => unknown, match?: string | RegExp): ExcInfo {
  let err: any = null;
  try {
    fn();
  } catch (e) {
    err = e;
  }
  if (err === null) expect.fail(`DID NOT RAISE ${Array.isArray(cls) ? cls.map((c: any) => c?.name).join(", ") : cls?.name ?? "Exception"}`);
  const classes = Array.isArray(cls) ? cls : cls ? [cls] : [];
  if (classes.length && !classes.some((c: any) => c && err instanceof c)) {
    expect.fail(`expected ${classes.map((c: any) => c?.name).join(" or ")}, got ${err?.name}: ${err?.message}\n${err?.stack}`);
  }
  if (match !== undefined) expect(String(err.message)).toMatch(typeof match === "string" ? new RegExp(match) : match);
  return { value: err, type: err?.constructor };
}

/** str(exception) / str(x) */
export function $str(v: any): string {
  if (v === null || v === undefined) return "None";
  if (typeof v === "boolean") return v ? "True" : "False";
  if (v instanceof Error) return pyErrStr(v);
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : pyFloatStr(v);
  if (Array.isArray(v)) return "[" + v.map($repr).join(", ") + "]";
  return String(v);
}

function pyErrStr(e: Error): string {
  // python str(KeyError(x)) is repr(x)
  if (e.name === "KeyError") return pyReprStr(e.message);
  return e.message;
}

export function $repr(v: any): string {
  if (v !== null && typeof v === "object" && typeof v.__repr__ === "function") return v.__repr__();
  if (typeof v === "string" || v instanceof String) return pyReprStr(String(v));
  if (Array.isArray(v)) return "[" + v.map($repr).join(", ") + "]";
  return $str(v);
}

export function $list(v: any): any[] {
  if (v === null || v === undefined) throw new TypeError("'NoneType' object is not iterable");
  if (typeof v === "string" || v instanceof String) return [...String(v)];
  if (typeof v[Symbol.iterator] === "function") return [...v];
  if (isPlain(v)) return Object.keys(v);
  return Array.from(v);
}

export function $range(a: number, b?: number, step = 1): number[] {
  if (b === undefined) {
    b = a;
    a = 0;
  }
  const out: number[] = [];
  for (let i = a; step > 0 ? i < b : i > b; i += step) out.push(i);
  return out;
}

/** dict(a, **kw) */
export function $dict(...parts: any[]): any {
  const out: any = new PyDict();
  for (const p of parts) {
    if (p === null || p === undefined) continue;
    if (Array.isArray(p) || (typeof p[Symbol.iterator] === "function" && !isPlain(p))) {
      for (const [k, v] of p as any) out[String(k)] = v;
    } else if (typeof p.items === "function") for (const [k, v] of p.items()) out[k] = v;
    else Object.assign(out, p[KW] ? stripKw(p) : p);
  }
  return out;
}
const stripKw = (o: any) => Object.fromEntries(Object.entries(o));

export function $items(o: any): [any, any][] {
  if (o instanceof Map) return [...o.entries()];
  if (typeof o.items === "function") return o.items();
  return Object.entries(o);
}
export function $keys(o: any): any[] {
  if (o instanceof Map) return [...o.keys()];
  if (typeof o.keys === "function" && !isPlain(o)) return [...o.keys()];
  return Object.keys(o);
}
export function $values(o: any): any[] {
  if (o instanceof Map) return [...o.values()];
  if (typeof o.values === "function" && !isPlain(o)) return [...o.values()];
  return Object.values(o);
}

export function $add(a: any, b: any): any {
  if (Array.isArray(a) || Array.isArray(b)) return [...(a ?? []), ...(b ?? [])];
  if (a instanceof Uint8Array && b instanceof Uint8Array) {
    const o = new Uint8Array(a.length + b.length);
    o.set(a);
    o.set(b, a.length);
    return o;
  }
  if (a instanceof String || b instanceof String) return String(a) + String(b);
  return a + b;
}

export function $mul(a: any, b: any): any {
  if (Array.isArray(a) && typeof b === "number") return Array.from({ length: b }, () => a).flat();
  if (typeof a === "string" && typeof b === "number") return a.repeat(b);
  return a * b;
}

export function $mod(a: any, b: any): any {
  if (typeof a === "string" || a instanceof String) return $pyformat(String(a), b);
  const r = a % b;
  return r !== 0 && r < 0 !== b < 0 ? r + b : r;
}

export function $zip(...arrs: any[]): any[][] {
  const lists = arrs.map($list);
  const n = Math.min(...lists.map((l) => l.length));
  return Array.from({ length: n }, (_, i) => lists.map((l) => l[i]));
}

export function $enumerate(a: any, start = 0): [number, any][] {
  return $list(a).map((v, i) => [i + start, v]);
}

const cmp = (x: any, y: any): number => {
  if (Array.isArray(x) && Array.isArray(y)) {
    for (let i = 0; i < Math.min(x.length, y.length); i++) {
      const c = cmp(x[i], y[i]);
      if (c) return c;
    }
    return x.length - y.length;
  }
  x = unwrap(x);
  y = unwrap(y);
  return x < y ? -1 : x > y ? 1 : 0;
};

export function $sorted(a: any, opts: any = {}): any[] {
  const arr = [...$list(a)];
  const key = opts.key ?? ((v: any) => v);
  arr.sort((x, y) => cmp(key(x), key(y)));
  return opts.reverse ? arr.reverse() : arr;
}

export const $any = (a: any) => $list(a).some($truthy);
export const $all = (a: any) => $list(a).every($truthy);
export const $sum = (a: any, start = 0) => $list(a).reduce((x, y) => x + y, start);
export const $set_ = (a: any = []) => new Set($list(a));
export const $min = (...a: any[]) => (a.length === 1 ? $list(a[0]) : a).reduce((x, y) => (cmp(y, x) < 0 ? y : x));
export const $max = (...a: any[]) => (a.length === 1 ? $list(a[0]) : a).reduce((x, y) => (cmp(y, x) > 0 ? y : x));

const ids = new WeakMap<object, number>();
let nextId = 1;
export function $id(o: any): number {
  if (!ids.has(o)) ids.set(o, nextId++);
  return ids.get(o)!;
}

const LENGTH_TYPES = new Set<any>([utilMod.Length, utilMod.Emu, utilMod.Inches, utilMod.Pt, utilMod.Cm, utilMod.Mm, utilMod.Centipoints]);

/** isinstance; str/int/float/bool/list/tuple/dict/bytes map onto JS types */
export function $isinstance(v: any, cls: any): boolean {
  if (Array.isArray(cls)) return cls.some((c) => $isinstance(v, c));
  switch (cls) {
    case "tuple":
    case "list":
    case Array:
      return Array.isArray(v);
    case "dict":
    case Object:
      return isPlain(v) || v instanceof Map;
    case "str":
    case String:
      return typeof v === "string" || v instanceof String;
    case "int":
      return typeof v === "number" && Number.isInteger(v);
    case "float":
    case Number:
      return typeof v === "number";
    case "bool":
    case Boolean:
      return typeof v === "boolean";
    case "bytes":
      return v instanceof Uint8Array;
  }
  // lengths are plain integer EMU numbers in pptx-ts (python: Length, an int subclass)
  if (LENGTH_TYPES.has(cls)) return v instanceof utilMod.Length || (typeof v === "number" && Number.isInteger(v));
  if (typeof cls !== "function") return false;
  try {
    return v instanceof cls;
  } catch {
    return false;
  }
}

/** "%s" % args, "{}".format(...) */
export function $pyformat(fmt: string, args: any): string {
  if (fmt.includes("%")) {
    const list = Array.isArray(args) ? [...args] : [args];
    return fmt.replace(/%(\d*)([sdrfxX%])/g, (_m, w, c) => {
      if (c === "%") return "%";
      const v = list.shift();
      let s = c === "r" ? $repr(v) : c === "d" ? String(Math.trunc(Number(v))) : c === "x" ? Number(v).toString(16) : c === "X" ? Number(v).toString(16).toUpperCase() : c === "f" ? Number(v).toFixed(6) : $str(v);
      if (w) s = s.padStart(Number(w), w.startsWith("0") ? "0" : " ");
      return s;
    });
  }
  let i = 0;
  return fmt.replace(/\{(\w*)\}/g, (_m, k) => (k === "" ? $str(args[i++]) : /^\d+$/.test(k) ? $str(args[+k]) : $str(args[k])));
}


export function $getattr(o: any, name: string, ...dflt: any[]): any {
  if ((typeof o === "number" || o instanceof Number) && ["inches", "cm", "mm", "pt", "emu", "centipoints"].includes(name)) return (new Length(Number(o)) as any)[name];
  const c = camel(name);
  for (const n of [c, name]) {
    if (o !== null && o !== undefined && n in Object(o)) {
      const v = o[n];
      return typeof v === "function" && !isMock(v) && !/^class[\s{]/.test(Function.prototype.toString.call(v)) ? v.bind(o) : v;
    }
  }
  if (dflt.length) return dflt[0];
  throw new TypeError(`AttributeError: '${o?.constructor?.name}' object has no attribute '${name}'`);
}

export function $setattr(o: any, name: string, v: any): void {
  const c = camel(name);
  o[c in Object(o) || !(name in Object(o)) ? c : name] = v;
}

export const $hasattr = (o: any, name: string) => o !== null && o !== undefined && (camel(name) in Object(o) || name in Object(o));
export const $callable = (o: any) => typeof o === "function";

export function $slice(v: any, lo?: number | null, hi?: number | null, step?: number | null): any {
  const isStr = typeof v === "string" || v instanceof String;
  const arr: any[] = isStr ? [...String(v)] : v instanceof Uint8Array ? Array.from(v) : $list(v);
  const n = arr.length;
  const norm = (i: number | null | undefined, d: number) => (i === null || i === undefined ? d : i < 0 ? Math.max(0, n + i) : Math.min(i, n));
  let out: any[];
  if (step && step < 0) {
    const start = lo === null || lo === undefined ? n - 1 : lo < 0 ? n + lo : Math.min(lo, n - 1);
    const stop = hi === null || hi === undefined ? -1 : hi < 0 ? n + hi : hi;
    out = [];
    for (let i = start; i > stop; i += step) out.push(arr[i]);
  } else {
    out = arr.slice(norm(lo, 0), norm(hi, n));
    if (step && step > 1) out = out.filter((_, i) => i % step === 0);
  }
  if (isStr) return out.join("");
  if (v instanceof Uint8Array) return new Uint8Array(out);
  return out;
}

export const $int = (v: any, base?: number) => {
  if (typeof v === "string") {
    const n = base ? parseInt(v, base) : Number(v.trim());
    if (Number.isNaN(n) || (!base && !/^\s*[+-]?\d+\s*$/.test(v))) throw new Error(`ValueError: invalid literal for int() with base ${base ?? 10}: '${v}'`);
    return n;
  }
  return Math.trunc(Number(v));
};
export const $float = (v: any) => Number(v);
export const $abs = (v: any) => Math.abs(v);
export const $round = (v: number, nd = 0) => {
  const m = 10 ** nd;
  const y = v * m;
  const r = Math.abs(y % 1) === 0.5 ? 2 * Math.round(y / 2) : Math.round(y);
  return r / m;
};

/** b"..." literal: latin-1 bytes */
export function $b(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

/** bytes(x) */
export function $bytes(v: any = 0): Uint8Array {
  if (typeof v === "number") return new Uint8Array(v);
  if (v instanceof Uint8Array) return v.slice();
  return new Uint8Array($list(v));
}

export const $encode = (s: string, encoding = "utf-8") => {
  const enc = encoding.toLowerCase().replace(/_/g, "-");
  if (enc === "utf-16-be" || enc === "utf-16-le") {
    const out = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      out[i * 2 + (enc === "utf-16-be" ? 0 : 1)] = c >> 8;
      out[i * 2 + (enc === "utf-16-be" ? 1 : 0)] = c & 0xff;
    }
    return out;
  }
  if (enc === "latin-1" || enc === "iso-8859-1") return Uint8Array.from(s, (ch) => ch.charCodeAt(0));
  return new TextEncoder().encode(s);
};
export const $decode = (b: Uint8Array) => new TextDecoder().decode(b);

export function $iter(v: any): Iterator<any> {
  return $list(v)[Symbol.iterator]();
}
export function $next(it: Iterator<any>, ...dflt: any[]): any {
  const r = it.next();
  if (r.done) {
    if (dflt.length) return dflt[0];
    throw new Error("StopIteration");
  }
  return r.value;
}

/** python `type(v).__name__` for the builtin types whose JS names differ */
export function $typename(v: any): string {
  if (typeof v === "string") return "str";
  if (typeof v === "boolean") return "bool";
  if (v === null || v === undefined) return "NoneType";
  return $type(v).name;
}
export const $type = (v: any) => (v === null || v === undefined ? null : Object.getPrototypeOf(Object(v)).constructor);

/** itertools.product */
export function $product(...lists: any[]): any[][] {
  return lists.map($list).reduce<any[][]>((acc, l) => acc.flatMap((a) => l.map((x) => [...a, x])), [[]]);
}

// ------------------------------------------------------------------------------------------------
// pytest runtime

export class PytestParam {
  constructor(
    public values: any[],
    public id?: string,
    public marks?: any,
  ) {}
}
/** pytest.param(*values, id=...) */
export function pytestParam(...a: any[]): PytestParam {
  let id: string | undefined;
  let marks: any;
  const last = a[a.length - 1];
  if (last && typeof last === "object" && last[KW]) {
    a = a.slice(0, -1);
    id = last.id;
    marks = last.marks;
  }
  return new PytestParam(a, id, marks);
}

interface FixtureDef {
  args: string[];
  fn: (...a: any[]) => any;
  params?: any[];
  ids?: any;
  scope?: string;
  autouse?: boolean;
}

interface TestDef {
  name: string;
  args: string[];
  fn: (...a: any[]) => any;
  parametrize?: { names: string[]; values: any[]; ids?: any }[];
  skip?: string | boolean;
  xfail?: string | boolean;
}

interface ClassDef {
  fixtures?: Record<string, FixtureDef>;
  tests: TestDef[];
  helpers?: Record<string, (...a: any[]) => any>;
  base?: ClassDef;
  attrs?: Record<string, any>;
}

const moduleFixtures: Record<string, FixtureDef> = {};

/** Register module-level fixtures (and run module-level tests). */
export function pytestModule(def: { fixtures?: Record<string, FixtureDef>; tests?: TestDef[] }): void {
  Object.assign(moduleFixtures, def.fixtures ?? {});
  if (def.tests?.length) runTests(def.tests, {}, () => ({}), "");
}

function paramId(v: any, i: number): string {
  if (v instanceof PytestParam && v.id) return v.id;
  const vals = v instanceof PytestParam ? v.values : Array.isArray(v) ? v : [v];
  const simple = vals.map((x: any) => (typeof x === "string" && x.length < 40 ? x : typeof x === "number" || typeof x === "boolean" || x === null ? String(x) : null));
  return simple.every((s: any) => s !== null) ? simple.join("-") : String(i);
}

/** Generate vitest cases for a pytest-style class (`class DescribeFoo:`). */
export function pytestClass(name: string, def: ClassDef): void {
  describe(name, () => {
    const fixtures: Record<string, FixtureDef> = {};
    const chain: ClassDef[] = [];
    for (let d: ClassDef | undefined = def; d; d = d.base) chain.unshift(d);
    for (const d of chain) Object.assign(fixtures, d.fixtures ?? {});
    const classCache = new Map<string, any>();
    const makeSelf = () => {
      const self: any = {};
      for (const d of chain) {
        for (const [k, fn] of Object.entries(d.helpers ?? {})) self[k] = (...a: any[]) => fn(self, ...a);
        Object.assign(self, d.attrs ?? {});
      }
      return self;
    };
    const tests = chain.flatMap((d) => d.tests);
    runTests(tests, fixtures, makeSelf, name, classCache);
  });
}

function runTests(tests: TestDef[], classFixtures: Record<string, FixtureDef>, makeSelf: () => any, _cls: string, classCache = new Map<string, any>()): void {
  const allFixtures = () => ({ ...moduleFixtures, ...classFixtures });
  for (const t of tests) {
    const fx = allFixtures();
    // parametrize marks: product of their cases
    let cases: { values: Record<string, any>; id: string }[] = [{ values: {}, id: "" }];
    for (const p of t.parametrize ?? []) {
      const next: typeof cases = [];
      p.values.forEach((v: any, i: number) => {
        const vals = v instanceof PytestParam ? v.values : p.names.length === 1 ? [v] : $list(v);
        const id = Array.isArray(p.ids) ? String(p.ids[i]) : typeof p.ids === "function" ? String(p.ids(v)) : paramId(p.names.length === 1 && !(v instanceof PytestParam) ? [v] : v, i);
        for (const c of cases) next.push({ values: { ...c.values, ...Object.fromEntries(p.names.map((n, j) => [n, vals[j]])) }, id: c.id ? `${c.id}-${id}` : id });
      });
      cases = next;
    }
    // parametrized fixtures in the closure
    const closure: string[] = [];
    const visit = (n: string, seen: Set<string>) => {
      if (seen.has(n) || n === "request" || n === "self") return;
      seen.add(n);
      const f = fx[n];
      if (!f) return;
      for (const a of f.args) if (!(a in cases[0].values)) visit(a, seen);
      if (f.params && !closure.includes(n)) closure.push(n);
    };
    const seen = new Set<string>();
    for (const a of t.args) if (!(a in cases[0].values)) visit(a, seen);
    for (const [n, f] of Object.entries(fx)) if (f.autouse) visit(n, seen);
    let fxCases: { idx: Record<string, number>; id: string }[] = [{ idx: {}, id: "" }];
    for (const n of closure) {
      const f = fx[n];
      const next: typeof fxCases = [];
      f.params!.forEach((v, i) => {
        const id = Array.isArray(f.ids) ? String(f.ids[i]) : paramId(v, i);
        for (const c of fxCases) next.push({ idx: { ...c.idx, [n]: i }, id: c.id ? `${c.id}-${id}` : id });
      });
      fxCases = next;
    }
    const runner = t.skip ? it.skip : it;
    for (const c of cases) {
      for (const fc of fxCases) {
        const id = [c.id, fc.id].filter(Boolean).join("-");
        const title = id ? `${t.name}[${id}]` : t.name;
        runner(title, () => {
          const finalizers: (() => void)[] = [];
          const cache = new Map<string, any>();
          const self = makeSelf();
          const resolve = (n: string, requester: string): any => {
            if (n in c.values) return c.values[n];
            if (n === "self") return self;
            if (n === "request") return makeRequest(requester);
            if (cache.has(n)) return cache.get(n);
            const f = fx[n];
            if (!f) throw new Error(`fixture '${n}' not found`);
            const ck = `${n}#${fc.idx[n] ?? ""}`;
            if (f.scope === "class" && classCache.has(ck)) {
              cache.set(n, classCache.get(ck));
              return classCache.get(ck);
            }
            const args = f.args.map((a) => resolve(a, n));
            let v = f.fn(self, ...args);
            if (v && typeof v === "object" && typeof v.next === "function" && typeof v[Symbol.iterator] === "function" && f.fn.constructor.name === "GeneratorFunction") {
              const gen = v;
              v = gen.next().value;
              finalizers.push(() => void gen.next());
            }
            cache.set(n, v);
            if (f.scope === "class") classCache.set(ck, v);
            return v;
          };
          const makeRequest = (fixturename: string): FixtureRequest => {
            const f = fx[fixturename];
            return {
              fixturename,
              param: f?.params ? unwrapParam(f.params[fc.idx[fixturename]]) : undefined,
              addfinalizer: (fn: () => void) => finalizers.push(fn),
              getfixturevalue: (n: string) => resolve(n, fixturename),
            };
          };
          try {
            for (const [n, f] of Object.entries(fx)) if (f.autouse) resolve(n, t.name);
            const args = t.args.map((a) => resolve(a, t.name));
            if (t.xfail) {
              let failed = false;
              try {
                t.fn(self, ...args);
              } catch {
                failed = true;
              }
              expect(failed, "XPASS (strict)").toBe(true);
            } else t.fn(self, ...args);
          } finally {
            for (const f of finalizers.reverse()) f();
          }
        });
      }
    }
  }
}

const unwrapParam = (v: any) => (v instanceof PytestParam ? (v.values.length === 1 ? v.values[0] : v.values) : v);

/** `with pytest.warns(...)` / `warnings.catch_warnings()` stand-in. */
export function $warns(_cls: any, fn: () => unknown): void {
  fn();
}

// ------------------------------------------------------------------------------------------------
// calls with keyword arguments, operators, builtin-method dispatch

export const $path = {
  join: (...p: string[]) => nodePath.join(...p),
  abspath: (p: string) => nodePath.resolve(p),
  dirname: (p: string) => nodePath.dirname(p),
  split: (p: string) => [nodePath.dirname(p), nodePath.basename(p)],
  splitext: (p: string) => {
    const e = nodePath.extname(p);
    return [p.slice(0, p.length - e.length), e];
  },
  exists: (p: string) => statOrNull(p) !== null,
  isfile: (p: string) => statOrNull(p)?.isFile() === true,
  isdir: (p: string) => statOrNull(p)?.isDirectory() === true,
  getsize: (p: string) => nodeFs().statSync(p).size,
};

const nodeFs = () => (globalThis as any).process.getBuiltinModule("node:fs");
function statOrNull(p: string): any {
  try {
    return nodeFs().statSync(p);
  } catch {
    return null;
  }
}

/** the `os` functions the tests use (os.path is $path) */
export const $os = { remove: (p: string) => nodeFs().unlinkSync(p), path: $path };

export const $sys = { maxsize: Number.MAX_SAFE_INTEGER, version_info: [3, 12, 0] };

/** typing.cast */
export const $cast = (_t: any, v: any) => v;

/** open(path, "rb") -> a BytesIO over the file's bytes */
export function $open(file: string, mode = "r"): any {
  const fs = (globalThis as any).process.getBuiltinModule("node:fs");
  if (mode.includes("w")) {
    const b = new BytesIO();
    // like a python file, the bytes reach the disk on close(), which leaving `with` calls
    (b as any).close = () => fs.writeFileSync(file, b.getvalue());
    (b as any).__exit__ = () => (b as any).close();
    return b;
  }
  const bytes = new Uint8Array(fs.readFileSync(file));
  if (mode.includes("b")) return Object.assign(new BytesIO(bytes), { close() {} });
  const text = new TextDecoder().decode(bytes);
  return { read: () => text, close() {}, [Symbol.iterator]: () => text.split(/(?<=\n)/)[Symbol.iterator]() };
}

/** itertools */
export const $itertools = {
  count: function* (start = 0, step = 1) {
    for (let n = start; ; n += step) yield n;
  },
  cycle: function* (it: any) {
    const l = $list(it);
    for (;;) yield* l;
  },
  chain: (...its: any[]) => its.flatMap($list),
  product: (...lists: any[]) => $product(...lists),
  /** `islice(it, n)`: the next `n` items, taken lazily so a shared or infinite iterator advances by `n` */
  islice: (it: any, n: number) => {
    const out: any[] = [];
    if (n <= 0) return out;
    const iter = it[Symbol.iterator]();
    for (let r = iter.next(); !r.done; r = iter.next()) {
      out.push(r.value);
      if (out.length >= n) break;
    }
    return out;
  },
  repeat: function* (v: any, n?: number) {
    for (let i = 0; n === undefined || i < n; i++) yield v;
  },
};

/** collections */
export const $collections = {
  OrderedDict: (init?: any) => $dict(init),
  defaultdict: (factory: any, init?: any) => {
    const d: any = $dict(init);
    const make = () => (factory === Array ? [] : factory === Object ? {} : $call(factory));
    Object.defineProperty(d, "get", { enumerable: false, value: (k: any) => (Object.prototype.hasOwnProperty.call(d, String(k)) ? d[String(k)] : (d[String(k)] = make())) });
    return d;
  },
  namedtuple: (_n: string, fields: string | string[]) => {
    const fs = typeof fields === "string" ? fields.split(/[ ,]+/) : fields;
    return (...a: any[]) => Object.assign([...a], Object.fromEntries(fs.map((f, i) => [f, a[i]])));
  },
};

/** hashlib.sha1(b).hexdigest() */
export const $hashlib = {
  sha1: (b: Uint8Array) => ({ hexdigest: () => createHash("sha1").update(b).digest("hex") }),
  md5: (b: Uint8Array) => ({ hexdigest: () => createHash("md5").update(b).digest("hex") }),
};

/** zipfile.ZipFile for reading archives in tests */
export const $zipfile = {
  ZIP_DEFLATED: 8,
  ZipFile: (file: any, _mode = "r") => {
    const bytes = file instanceof Uint8Array ? file : typeof file === "string" ? new Uint8Array((globalThis as any).process.getBuiltinModule("node:fs").readFileSync(file)) : file.getvalue ? file.getvalue() : file.read();
    const files = unzipSync(bytes);
    return {
      namelist: () => Object.keys(files),
      read: (name: string) => files[name],
      close: () => undefined,
    };
  },
};

export const pytest = {
  fail(msg = "") {
    expect.fail(msg);
  },
  skip(msg = "") {
    throw new Error(`skip: ${msg}`);
  },
};

/** Bind trailing `$kw({...})` keyword args to `fn`'s parameters (python call binding). */
function bindKw(fn: Function, args: any[]): any[] {
  const last = args[args.length - 1];
  if (!(last !== null && typeof last === "object" && last[KW])) return args;
  const pos = args.slice(0, -1);
  const sig = signatureOf(fn) ?? [];
  // a src parameter is the camelCased python name; a test-local helper keeps the python name
  const kw: Record<string, any> = Object.fromEntries(
    Object.entries(last).map(([k, v]) => [sig.some((p) => p.name === k && !p.rest && !p.keys) ? k : (mockMod.KW_ALIASES[k] ?? camel(k)), v]),
  );
  const out = pos.slice();
  const leftover: Record<string, any> = {};
  for (const [k, v] of Object.entries(kw)) {
    const i = sig.findIndex((p) => p.name === k && !p.rest && !p.keys);
    if (i >= 0) {
      while (out.length < i) out.push(undefined);
      out[i] = v;
    } else leftover[k] = v;
  }
  if (Object.keys(leftover).length) {
    // an options-object parameter (destructured, or named opts/options/kwargs) takes the rest
    let i = sig.findIndex((p) => p.keys || /^(opts|options|kwargs|kw|attrs|props)$/.test(p.name));
    if (i < 0) i = Math.max(pos.length, 0);
    while (out.length < i) out.push(undefined);
    out[i] = Object.assign({}, out[i] ?? {}, leftover);
  }
  return out;
}

/** python call `f(*args, **kwargs)`: classes are constructed, keyword args bound by signature */
export function $call(f: any, ...args: any[]): any {
  if (isMock(f)) return f(...args);
  if (typeof f !== "function") throw new TypeError(`'${f === null ? "NoneType" : typeof f}' object is not callable`);
  if (f.__pyPassKw) return f(...args);
  const bound = bindKw(f, args);
  if (/^class[\s{]/.test(Function.prototype.toString.call(f))) return new f(...bound);
  return f(...bound);
}
export const $newm = $call;

/** python method call `obj.name(*args, **kwargs)` */
export function $m(obj: any, name: string, ...args: any[]): any {
  if (obj === null || obj === undefined) throw new TypeError(`'NoneType' object has no attribute '${name}'`);
  const fn = obj[name];
  if (isMock(fn)) return fn(...args);
  if (typeof fn !== "function") throw new TypeError(`AttributeError: '${obj?.constructor?.name}' object has no attribute '${name}'`);
  if (fn.__pyPassKw) return fn.apply(obj, args);
  return fn.apply(obj, bindKw(fn, args));
}

for (const k of ["classMock", "clsAttrMock", "functionMock", "initializerMock", "instanceMock", "looseMock", "methodMock", "openMock", "propertyMock", "varMock", "Mock", "MagicMock", "PropertyMock", "NonCallableMagicMock", "createAutospec", "$kw"]) {
  const f = (mockMod as any)[k];
  if (typeof f === "function") (f as any).__pyPassKw = true;
}
(pytestParam as any).__pyPassKw = true;

/** python `a and b and ...` (returns an operand) */
export function $and(...thunks: (() => any)[]): any {
  let v: any;
  for (const t of thunks) {
    v = t();
    if (!$truthy(v)) return v;
  }
  return v;
}

/** python `a or b or ...` (returns an operand) */
export function $orv(...thunks: (() => any)[]): any {
  let v: any;
  for (const t of thunks) {
    v = t();
    if ($truthy(v)) return v;
  }
  return v;
}

/** python `|`: int bitwise or, set union, dict merge */
export function $or(a: any, b: any): any {
  if (a instanceof Set) return new Set([...a, ...b]);
  if (isPlain(a)) return { ...a, ...b };
  return a | b;
}

export function $cmp(a: any, op: string, b: any): boolean {
  const c = cmp(a, b);
  return op === "<" ? c < 0 : op === "<=" ? c <= 0 : op === ">" ? c > 0 : c >= 0;
}

/** format(value, spec) for the specs the tests use */
export function $format(v: any, spec: string): string {
  const m = /^(0?)(\d*)(?:\.(\d+))?([dfxXse%]?)$/.exec(spec);
  if (!m) return $str(v);
  const [, zero, width, prec, type] = m;
  let s: string;
  if (type === "d") s = String(Math.trunc(v));
  else if (type === "f") s = Number(v).toFixed(prec ? Number(prec) : 6);
  else if (type === "x") s = Number(v).toString(16);
  else if (type === "X") s = Number(v).toString(16).toUpperCase();
  else if (type === "%") s = (Number(v) * 100).toFixed(prec ? Number(prec) : 6) + "%";
  else if (prec) s = Number(v).toFixed(Number(prec));
  else s = $str(v);
  return width ? s.padStart(Number(width), zero ? "0" : " ") : s;
}

const isArr = Array.isArray;
const isS = (o: any) => typeof o === "string" || o instanceof String;
const callOwn = (o: any, name: string, args: any[]) => {
  const fn = o?.[name];
  if (typeof fn !== "function") throw new TypeError(`AttributeError: '${o?.constructor?.name}' object has no attribute '${name}'`);
  return isMock(fn) ? fn(...args) : fn.apply(o, bindKw(fn, args));
};

export const $append = (o: any, x: any) => (isArr(o) ? void o.push(x) : callOwn(o, "append", [x]));
export const $extend = (o: any, x: any) => (isArr(o) ? void o.push(...$list(x)) : callOwn(o, "extend", [x]));
export const $insert = (o: any, i: number, x: any) => (isArr(o) ? void o.splice(i < 0 ? Math.max(0, o.length + i) : i, 0, x) : callOwn(o, "insert", [i, x]));
export function $remove(o: any, x: any): any {
  if (isArr(o)) {
    const i = o.findIndex((e) => $eq(e, x));
    if (i < 0) throw new Error("ValueError: list.remove(x): x not in list");
    o.splice(i, 1);
    return;
  }
  if (o instanceof Set) return void o.delete(x);
  return callOwn(o, "remove", [x]);
}
export function $pop(o: any, ...a: any[]): any {
  if (isArr(o)) {
    const i = a.length ? (a[0] < 0 ? o.length + a[0] : a[0]) : o.length - 1;
    return o.splice(i, 1)[0];
  }
  if (isPlain(o)) {
    const k = String(a[0]);
    if (k in o) {
      const v = o[k];
      delete o[k];
      return v;
    }
    if (a.length > 1) return a[1];
    throw new Error(`KeyError: ${k}`);
  }
  if (o instanceof Map) {
    const v = o.get(a[0]);
    o.delete(a[0]);
    return v;
  }
  return callOwn(o, "pop", a);
}
export const $index = (o: any, x: any) => {
  if (isArr(o)) {
    const i = o.findIndex((e) => $eq(e, x));
    if (i < 0) throw new Error("ValueError: not in list");
    return i;
  }
  if (isS(o)) {
    const i = String(o).indexOf(x);
    if (i < 0) throw new Error("ValueError: substring not found");
    return i;
  }
  return callOwn(o, "index", [x]);
};
export const $count = (o: any, x: any) => (isArr(o) ? o.filter((e) => $eq(e, x)).length : isS(o) ? String(o).split(x).length - 1 : callOwn(o, "count", [x]));
export const $find = (o: any, ...a: any[]) => (isS(o) ? String(o).indexOf(a[0], a[1]) : callOwn(o, "find", a));
export const $replace = (o: any, ...a: any[]) => (isS(o) ? String(o).split(a[0]).join(a[1]) : callOwn(o, "replace", a));
export const $copy = (o: any) => (isArr(o) ? [...o] : isPlain(o) ? { ...o } : o instanceof Map ? new Map(o) : o instanceof Set ? new Set(o) : callOwn(o, "copy", []));
export const $update = (o: any, ...a: any[]) =>
  isPlain(o) ? void Object.assign(o, ...a.map((x) => (x?.[KW] ? { ...x } : $dict(x)))) : o instanceof Set ? void a.forEach((x) => $list(x).forEach((e) => o.add(e))) : callOwn(o, "update", a);
export const $dictget = (o: any, k: any, d: any = null) => {
  if (isPlain(o)) return Object.prototype.hasOwnProperty.call(o, String(k)) ? o[String(k)] : d;
  if (o instanceof Map) return o.has(k) ? o.get(k) : d;
  return callOwn(o, "get", d === null ? [k] : [k, d]);
};
export const $setdefault = (o: any, k: any, d: any = null) => {
  if (!(String(k) in o)) o[String(k)] = d;
  return o[String(k)];
};
export const $split = (o: any, sep?: string | null, max?: number) => {
  // a non-string with its own split() (e.g. a table cell) is a method call
  if (o !== null && typeof o === "object" && !(o instanceof String) && typeof o.split === "function") return o.split(...[sep, max].filter((x) => x !== undefined));
  const s = String(o);
  if (sep === undefined || sep === null) return s.trim() === "" ? [] : s.trim().split(/\s+/);
  const parts = s.split(sep);
  if (max !== undefined && max >= 0 && parts.length > max + 1) return [...parts.slice(0, max), parts.slice(max).join(sep)];
  return parts;
};
const stripChars = (s: string, chars: string | undefined, l: boolean, r: boolean) => {
  if (chars === undefined || chars === null) return l && r ? s.trim() : l ? s.trimStart() : s.trimEnd();
  let a = 0;
  let b = s.length;
  if (l) while (a < b && chars.includes(s[a])) a++;
  if (r) while (b > a && chars.includes(s[b - 1])) b--;
  return s.slice(a, b);
};
export const $strip = (o: any, c?: string) => (o instanceof Uint8Array ? $encode(stripChars($decode(o), c, true, true)) : stripChars(String(o), c, true, true));
export const $lstrip = (o: any, c?: string) => stripChars(String(o), c, true, false);
export const $rstrip = (o: any, c?: string) => stripChars(String(o), c, false, true);
export const $lower = (o: any) => String(o).toLowerCase();
export const $upper = (o: any) => String(o).toUpperCase();
export const $startswith = (o: any, p: any) => (isArr(p) ? p.some((x) => String(o).startsWith(x)) : String(o).startsWith(p));
export const $endswith = (o: any, p: any) => (isArr(p) ? p.some((x) => String(o).endsWith(x)) : String(o).endsWith(p));
export const $splitlines = (o: any) => {
  const l = String(o).split(/\r\n|\r|\n/);
  if (l.length && l[l.length - 1] === "") l.pop();
  return l;
};
export const $isdigit = (o: any) => /^\d+$/.test(String(o));
export const $read = (o: any, ...a: any[]) => callOwn(o, "read", a);
export const $clear = (o: any) => (isArr(o) ? void (o.length = 0) : isPlain(o) ? void Object.keys(o).forEach((k) => delete o[k]) : callOwn(o, "clear", []));
export const $sort = (o: any, opts: any = {}) => {
  const s = $sorted(o, opts?.[KW] ? { ...opts } : opts);
  o.splice(0, o.length, ...s);
};

/** The `pptx` package module: `__file__` places `templates/` (the bundled template sources) where python-pptx's are. */
export const $pptx = { __file__: new URL("../__init__.py", import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1") };

/**
 * python's `datetime` module, as pptx-ts represents it: a naive datetime is a `Date` holding that
 * wall-clock time in UTC (see src/oxml/coreprops.ts); a timedelta is a number of milliseconds.
 */
class PyDate extends Date {}
class PyDatetime extends PyDate {
  /** `dt.replace(tzinfo=None)`: datetimes here are already naive. */
  replace(_kw?: any): PyDatetime {
    return this;
  }
}
export const $datetime = {
  /** `date(y, m, d)`: midnight UTC (python's `datetime` subclasses `date`, as here). */
  date: Object.assign(
    function date(y: number, mo: number, d: number): PyDate {
      const t = new PyDate(Date.UTC(y, mo - 1, d));
      t.setUTCFullYear(y);
      return t;
    },
    { today: () => { const n = new Date(); return new PyDate(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())); } },
  ),
  datetime: Object.assign(
    function datetime(y: number, mo: number, d: number, h = 0, mi = 0, s = 0, us = 0): PyDatetime {
      const t = new PyDatetime(Date.UTC(y, mo - 1, d, h, mi, s, Math.floor(us / 1000)));
      t.setUTCFullYear(y);
      return t;
    },
    { now: (_tz?: any) => new PyDatetime(Date.now()), utcnow: () => new PyDatetime(Date.now()) },
  ),
  timedelta: (kw: any = {}) => {
    const o = kw?.kwargs ?? kw;
    return ((o.days ?? 0) * 86400 + (o.hours ?? 0) * 3600 + (o.minutes ?? 0) * 60 + (o.seconds ?? 0)) * 1000 + (o.milliseconds ?? 0);
  },
  timezone: { utc: null },
};
(($datetime.datetime as any).__pyPassKw = true);

/** `with cm as x:` entry: `cm.__enter__()` when defined, else `cm` itself. */
export function $enter(cm: any): any {
  return cm !== null && cm !== undefined && typeof cm.__enter__ === "function" ? cm.__enter__() : cm;
}

/** `with` exit: `cm.__exit__(None, None, None)` when defined. */
export function $exit(cm: any): void {
  if (cm !== null && cm !== undefined && typeof cm.__exit__ === "function") cm.__exit__(null, null, null);
}
