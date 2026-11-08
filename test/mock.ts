// A port of the parts of python's `unittest.mock` the python-pptx suite uses: Mock /
// MagicMock / PropertyMock, autospec'd instance and class mocks, `call`, `ANY`, and the
// patching helpers of tests/unitutil/mock.py (class_mock, method_mock, ...).
//
// Mock API names are camelCase like the rest of the port: returnValue, sideEffect,
// callArgs, callArgsList, callCount, mockCalls, assertCalledOnceWith, ...
// Keyword arguments in expected calls are marked with `$kw({...})`; signatures of the real
// functions (parsed from their source) are used to normalize positional vs keyword
// arguments, as autospec does in python.
import { expect } from "vitest";
import { AttributeError } from "../src/errors.js";

// ------------------------------------------------------------------------------------------------
// keyword arguments and signatures

export const KW = Symbol("kwargs");
/** Marks keyword arguments in a translated call: f(a, b=1) -> f(a, $kw({b: 1})). */
export function $kw(o: Record<string, any>): any {
  return Object.assign(Object.create(null), o, { [KW]: true });
}
const isKw = (v: any) => v !== null && typeof v === "object" && v[KW] === true;

/** python keyword names that can't be JS parameter names, and the names the port uses */
export const KW_ALIASES: Record<string, string> = { package: "pkg", default: "dflt", class: "cls", new: "new_", type: "type" };

export const camel = (name: string): string => {
  const m = /^_*/.exec(name)![0];
  return m + name.slice(m.length).replace(/_+([A-Za-z0-9])/g, (_s, c: string) => c.toUpperCase());
};

interface Param {
  name: string;
  rest?: boolean;
  /** names inside a destructured `{a, b}` options parameter */
  keys?: string[];
}

const sigCache = new WeakMap<Function, Param[] | null>();

function splitTop(s: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let cur = "";
  let q: string | null = null;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) {
      cur += c;
      if (c === "\\") cur += s[++i] ?? "";
      else if (c === q) q = null;
      continue;
    }
    if (c === '"' || c === "'" || c === "`") q = c;
    if ("([{".includes(c)) depth++;
    if (")]}".includes(c)) depth--;
    if (c === "," && depth === 0) {
      out.push(cur);
      cur = "";
    } else cur += c;
  }
  if (cur.trim()) out.push(cur);
  return out;
}

/** Parameter list of a function, method or class constructor, from its source text. */
export function signatureOf(fn: Function | undefined | null): Param[] | null {
  if (typeof fn !== "function") return null;
  if (sigCache.has(fn)) return sigCache.get(fn)!;
  let src = Function.prototype.toString.call(fn);
  let params: Param[] | null = null;
  try {
    if (/^class[\s{]/.test(src)) {
      const m = /\bconstructor\s*\(/.exec(src);
      if (!m) {
        const parent = Object.getPrototypeOf(fn);
        params = parent && parent !== Function.prototype ? signatureOf(parent) : [];
        sigCache.set(fn, params);
        return params;
      }
      src = src.slice(m.index + m[0].length - 1);
    } else {
      const open = src.indexOf("(");
      const arrow = src.indexOf("=>");
      if (open < 0 || (arrow >= 0 && arrow < open)) {
        // single-parameter arrow function `x => ...`
        const name = src.slice(0, arrow).trim().replace(/^async\s+/, "");
        params = [{ name }];
        sigCache.set(fn, params);
        return params;
      }
      src = src.slice(open);
    }
    let depth = 0;
    let end = 0;
    for (let i = 0; i < src.length; i++) {
      if (src[i] === "(") depth++;
      else if (src[i] === ")" && --depth === 0) {
        end = i;
        break;
      }
    }
    const inner = src.slice(1, end).replace(/\/\*[\s\S]*?\*\//g, "");
    params = splitTop(inner).map((p) => {
      p = p.trim();
      if (p.startsWith("...")) return { name: p.slice(3).trim(), rest: true };
      if (p.startsWith("{")) {
        const body = p.slice(1, p.lastIndexOf("}"));
        const keys = splitTop(body).map((k) => k.trim().split(/[=:]/)[0].trim()).filter(Boolean);
        return { name: "$opts", keys };
      }
      return { name: p.split("=")[0].trim() };
    });
  } catch {
    params = null;
  }
  sigCache.set(fn, params);
  return params;
}

/** A call as {pos, kw}. With a signature, positional args are bound to names (python binding). */
function normalize(args: any[], sig: Param[] | null): { pos: any[]; kw: Record<string, any> } {
  let pos = args.slice();
  let kw: Record<string, any> = {};
  if (pos.length && isKw(pos[pos.length - 1])) {
    const k = pos.pop();
    for (const key of Object.keys(k)) kw[KW_ALIASES[key] ?? camel(key)] = k[key];
  }
  while (pos.length && pos[pos.length - 1] === undefined) pos.pop();
  if (!sig) return { pos, kw };
  const bound: Record<string, any> = {};
  const extra: any[] = [];
  const restIdx = sig.findIndex((p) => p.rest);
  pos.forEach((v, i) => {
    const p = sig[i];
    if (restIdx >= 0 && i >= restIdx) extra.push(v);
    else if (!p) extra.push(v);
    else if (p.keys) {
      if (v && typeof v === "object" && !Array.isArray(v)) for (const k of Object.keys(v)) bound[camel(k)] = v[k];
      else bound[p.name] = v;
    } else bound[p.name] = v;
  });
  for (const [k, v] of Object.entries(kw)) bound[k] = v;
  for (const k of Object.keys(bound)) if (bound[k] === undefined) delete bound[k];
  return { pos: extra, kw: bound };
}

// ------------------------------------------------------------------------------------------------
// call objects

/** A recorded or expected call, like python's `mock.call(...)` / `mock.call.foo(...)`. */
export class _Call {
  constructor(
    public name: string,
    public args: any[],
    public sig: Param[] | null = null,
  ) {}

  /** positional arguments (python `call_args.args` / `call_args[0]`) */
  get pyArgs(): any[] {
    const { pos } = normalize(this.args, null);
    return pos;
  }

  /** keyword arguments (python `call_args.kwargs` / `call_args[1]`) */
  get kwargs(): Record<string, any> {
    return normalize(this.args, null).kw;
  }

  [Symbol.for("py.getitem")](i: number): any {
    if (i === 0 || i === -2) return this.pyArgs;
    if (i === 1 || i === -1) return this.kwargs;
    throw new RangeError("tuple index out of range");
  }

  [Symbol.iterator](): Iterator<any> {
    return [this.pyArgs, this.kwargs][Symbol.iterator]();
  }

  get length(): number {
    return 2;
  }

  equals(other: unknown): boolean {
    if (!(other instanceof _Call)) {
      // (args, kwargs) tuple or just an args tuple
      if (Array.isArray(other)) {
        if (other.length === 2 && Array.isArray(other[0]) && other[1] && typeof other[1] === "object" && !Array.isArray(other[1]))
          return pyEq(this.pyArgs, other[0]) && pyEq(this.kwargs, other[1]);
        return pyEq(this.pyArgs, other) && Object.keys(this.kwargs).length === 0;
      }
      return false;
    }
    if (this.name !== other.name && this.name !== "" && other.name !== "") return false;
    if (this.name !== other.name) return false;
    const sig = this.sig ?? other.sig;
    const a = normalize(this.args, sig);
    const b = normalize(other.args, sig);
    return pyEq(a.pos, b.pos) && pyEq(a.kw, b.kw);
  }

  toString(): string {
    const parts = this.args.map((a) => (isKw(a) ? Object.entries(a).map(([k, v]) => `${k}=${show(v)}`).join(", ") : show(a)));
    return `call${this.name ? "." + this.name : ""}(${parts.join(", ")})`;
  }
}

function makeCallBuilder(name: string): any {
  const f = (...args: any[]) => new _Call(name, args);
  return new Proxy(f, {
    get(t, prop) {
      if (typeof prop === "symbol" || prop === "prototype" || prop === "length" || prop === "name" || prop === "call" || prop === "apply" || prop === "bind")
        return (t as any)[prop];
      return makeCallBuilder(name ? `${name}.${prop}` : prop);
    },
  });
}

/** `mock.call`: call(1, 2), call.foo(1), call.foo.bar(). */
export const call: any = makeCallBuilder("");

/** `mock.ANY`: equal to anything. */
export const ANY = new (class ANY {
  equals() {
    return true;
  }
  toString() {
    return "<ANY>";
  }
})();

// equality that understands ANY, _Call, arrays, plain objects and `.equals`
let pyEqImpl: (a: any, b: any) => boolean = (a, b) => a === b;
export function setPyEq(fn: (a: any, b: any) => boolean): void {
  pyEqImpl = fn;
}
function pyEq(a: any, b: any): boolean {
  if (a === ANY || b === ANY) return true;
  return pyEqImpl(a, b);
}

let showImpl: (v: any) => string = (v) => String(v);
export function setShow(fn: (v: any) => string): void {
  showImpl = fn;
}
function show(v: any): string {
  try {
    return showImpl(v);
  } catch {
    return "<?>";
  }
}

// ------------------------------------------------------------------------------------------------
// mocks

const MOCK = Symbol("mock");

interface MockState {
  name: string;
  parent: MockState | null;
  callable: boolean;
  /** attribute names allowed (spec) or null for "anything" */
  spec: Set<string> | null;
  specProto: object | null;
  /** functions (from the spec) whose signatures normalize calls on children */
  specMembers: Map<string, { fn?: Function; kind: "method" | "prop" | "static" | "class" }> | null;
  sig: Param[] | null;
  /** class to autospec the return value from (class mocks) */
  instanceSpecOf: Function | null;
  children: Map<string, any>;
  attrs: Map<string | symbol, any>;
  returnValue: any;
  hasReturnValue: boolean;
  sideEffect: any;
  sideEffectIter: Iterator<any> | null;
  calls: _Call[];
  mockCalls: _Call[];
  isProperty: boolean;
}

const mockIds = new WeakMap<MockState, number>();
let nextMockId = 1;
function mockId(st: MockState): number {
  let v = mockIds.get(st);
  if (v === undefined) mockIds.set(st, (v = nextMockId++));
  return v;
}

/** True if `v` is a mock created here. */
export const isMock = (v: any): boolean => v !== null && (typeof v === "object" || typeof v === "function") && !!v[MOCK];

function newState(name: string, parent: MockState | null, callable: boolean): MockState {
  return {
    name,
    parent,
    callable,
    spec: null,
    specProto: null,
    specMembers: null,
    sig: null,
    instanceSpecOf: null,
    children: new Map(),
    attrs: new Map(),
    returnValue: undefined,
    hasReturnValue: false,
    sideEffect: undefined,
    sideEffectIter: null,
    calls: [],
    mockCalls: [],
    isProperty: false,
  };
}

/** Record a call on `st` (and in its ancestors' mockCalls) and produce its result. */
function invoke(st: MockState, proxy: any, args: any[]): any {
  const c = new _Call("", args, st.sig);
  st.calls.push(c);
  st.mockCalls.push(c);
  let path = st.name;
  for (let p = st.parent, child: MockState = st; p; child = p, p = p.parent) {
    p.mockCalls.push(new _Call(path, args, st.sig));
    path = `${p.name}.${path}`;
  }
  const se = st.sideEffect;
  if (se !== undefined && se !== null) {
    if (se instanceof Error) throw se;
    if (typeof se === "function" && /^class[\s{]/.test(Function.prototype.toString.call(se)) && se.prototype instanceof Error) throw new se();
    if (typeof se === "function" && !isMock(se)) {
      const r = se(...args.map((a) => (isKw(a) ? { ...a } : a)));
      if (r !== DEFAULT) return r;
    } else if (isMock(se)) {
      const r = se(...args);
      if (r !== DEFAULT) return r;
    } else {
      st.sideEffectIter ??= (se as Iterable<any>)[Symbol.iterator]();
      const n = st.sideEffectIter.next();
      if (n.done) throw new Error("StopIteration");
      if (n.value instanceof Error) throw n.value;
      return n.value;
    }
  }
  return getReturnValue(st, proxy);
}

/** `mock.DEFAULT`: a side_effect function returning it falls through to return_value. */
export const DEFAULT = Symbol("DEFAULT");

function getReturnValue(st: MockState, _proxy: any): any {
  if (!st.hasReturnValue) {
    st.returnValue = st.instanceSpecOf ? makeMock({ name: "()", parent: st, specClass: st.instanceSpecOf, instance: true }) : makeMock({ name: "()", parent: st });
    st.hasReturnValue = true;
  }
  return st.returnValue;
}

/** Names of members of class `cls` (instance side, or static side). */
function classMembers(cls: Function, staticSide: boolean): Map<string, { fn?: Function; kind: "method" | "prop" | "static" | "class" }> {
  const out = new Map<string, { fn?: Function; kind: "method" | "prop" | "static" | "class" }>();
  if (staticSide) {
    for (let c: any = cls; c && c !== Function.prototype; c = Object.getPrototypeOf(c)) {
      for (const k of Object.getOwnPropertyNames(c)) {
        if (["prototype", "length", "name", "caller", "arguments", "__pyBindings"].includes(k) || out.has(k)) continue;
        const d = Object.getOwnPropertyDescriptor(c, k)!;
        if (typeof d.value === "function") out.set(k, { fn: d.value, kind: /^class[\s{]/.test(Function.prototype.toString.call(d.value)) ? "class" : "static" });
        else out.set(k, { kind: "prop" });
      }
    }
  }
  for (let p: any = cls.prototype; p && p !== Object.prototype; p = Object.getPrototypeOf(p)) {
    for (const k of Object.getOwnPropertyNames(p)) {
      if (k === "constructor" || out.has(k)) continue;
      const d = Object.getOwnPropertyDescriptor(p, k)!;
      if (d.get || d.set) out.set(k, { kind: "prop", fn: d.get });
      else if (typeof d.value === "function") out.set(k, { fn: d.value, kind: "method" });
      else out.set(k, { kind: "prop" });
    }
  }
  return out;
}

const MOCK_API = new Set([
  "returnValue",
  "sideEffect",
  "called",
  "callCount",
  "callArgs",
  "callArgsList",
  "mockCalls",
  "methodCalls",
  "assertCalledOnceWith",
  "assertCalledWith",
  "assertCalledOnce",
  "assertCalled",
  "assertNotCalled",
  "assertHasCalls",
  "assertAnyCall",
  "resetMock",
  "configureMock",
  "_mockName",
  "_mockState",
  "attachMockSig",
]);

// names the JS runtime / vitest probe that must not be auto-created
const PASSTHROUGH = new Set(["then", "toJSON", "asymmetricMatch", "$$typeof", "nodeType", "tagName", "constructor", "_isMockFunction", "__esModule", "@@__IMMUTABLE_ITERABLE__@@", "@@__IMMUTABLE_RECORD__@@", "toString", "valueOf", "inspect"]);

// python protocol dunders and the JS spellings the source uses for them
const PROTOCOL_ALIAS: Record<string, string> = { get: "__getitem__", at: "__getitem__", has: "__contains__" };

interface MakeOpts {
  name?: string;
  parent?: MockState | null;
  callable?: boolean;
  specClass?: Function | null;
  /** autospec the instance side of specClass (else the class itself) */
  instance?: boolean;
  specSet?: boolean;
  sigFn?: Function | null;
  props?: Record<string, any>;
}

function makeMock(o: MakeOpts = {}): any {
  const callable = o.callable ?? !o.instance;
  const st = newState(o.name ?? "mock", o.parent ?? null, callable);
  if (o.specClass) {
    if (o.instance) {
      st.specMembers = classMembers(o.specClass, false);
      st.specProto = o.specClass.prototype;
    } else {
      st.specMembers = classMembers(o.specClass, true);
      st.instanceSpecOf = o.specClass;
      st.sig = signatureOf(o.specClass);
      st.specProto = Function.prototype;
    }
    st.spec = new Set(st.specMembers.keys());
  } else if (o.sigFn) st.sig = signatureOf(o.sigFn);
  const target: any = callable ? function mock() {} : {};
  let proxy: any;
  const handler: ProxyHandler<any> = {
    apply(_t, _this, args) {
      return invoke(st, proxy, args);
    },
    construct(_t, args) {
      const r = invoke(st, proxy, args);
      return r !== null && (typeof r === "object" || typeof r === "function") ? r : {};
    },
    getPrototypeOf() {
      return st.specProto ?? (callable ? Function.prototype : Object.prototype);
    },
    has(_t, prop) {
      if (st.attrs.has(prop)) return true;
      if (typeof prop === "symbol") return false;
      if (st.spec) return st.spec.has(prop) || MOCK_API.has(prop);
      return true;
    },
    set(_t, prop, value) {
      if (prop === "returnValue") {
        st.returnValue = value;
        st.hasReturnValue = true;
      } else if (prop === "sideEffect") {
        st.sideEffect = value;
        st.sideEffectIter = null;
      } else {
        st.attrs.set(prop, value);
        if (typeof prop === "string") st.children.delete(prop);
      }
      return true;
    },
    deleteProperty(_t, prop) {
      st.attrs.delete(prop);
      if (typeof prop === "string") st.children.delete(prop);
      return true;
    },
    get(t, prop) {
      if (prop === MOCK) return true;
      if (st.attrs.has(prop)) return st.attrs.get(prop);
      if (typeof prop === "symbol") {
        if (prop === Symbol.iterator) {
          if (st.children.has("__iter__") || !st.spec || st.spec.has(String(Symbol.iterator)) || specHasIter(st)) {
            const it = child("__iter__");
            return () => {
              const r = it();
              return isMock(r) ? [][Symbol.iterator]() : r[Symbol.iterator] ? r[Symbol.iterator]() : r;
            };
          }
          return undefined;
        }
        if (prop === Symbol.for("py.getitem")) return (k: any) => child("__getitem__")(k);
        // like MagicMock, an unconfigured mock is truthy; len() only when __len__ is set up
        if (prop === Symbol.for("py.len")) return st.children.has("__len__") || st.attrs.has("__len__") ? () => child("__len__")() : undefined;
        if (prop === Symbol.for("py.contains")) return (k: any) => child("__contains__")(k);
        // mocks are hashable: a distinct string per mock lets them key python dicts
        if (prop === Symbol.toPrimitive) return () => `<Mock name='${st.name}' id='${mockId(st)}'>`;
        return (t as any)[prop];
      }
      switch (prop) {
        case "returnValue":
          return getReturnValue(st, proxy);
        case "sideEffect":
          return st.sideEffect;
        case "called":
          return st.calls.length > 0;
        case "callCount":
          return st.calls.length;
        case "callArgs":
          return st.calls.length ? st.calls[st.calls.length - 1] : null;
        case "callArgsList":
          return new CallList(st.calls);
        case "mockCalls":
          return new CallList(st.mockCalls);
        case "methodCalls":
          return new CallList(st.mockCalls.filter((c) => c.name !== "" && !c.name.includes("()")));
        case "_mockName":
          return st.name;
        case "_mockState":
          return st;
        case "resetMock":
          return () => {
            st.calls.length = 0;
            st.mockCalls.length = 0;
            for (const c of st.children.values()) c.resetMock();
          };
        case "configureMock":
          return (kw: Record<string, any>) => {
            for (const [k, v] of Object.entries(kw)) {
              const parts = camel(k).split(".");
              let o: any = proxy;
              for (const p of parts.slice(0, -1)) o = o[p];
              o[parts[parts.length - 1]] = v;
            }
          };
        case "assertCalledOnceWith":
          return (...expected: any[]) => {
            expect(st.calls.length, `Expected '${st.name}' to be called once. Called ${st.calls.length} times.${callsHint(st)}`).toBe(1);
            assertCallEq(st, st.calls[0], expected);
          };
        case "assertCalledWith":
          return (...expected: any[]) => {
            expect(st.calls.length, `expected call not found. Expected: ${new _Call("", expected)} Actual: not called.`).toBeGreaterThan(0);
            assertCallEq(st, st.calls[st.calls.length - 1], expected);
          };
        case "assertCalledOnce":
          return () => expect(st.calls.length, `Expected '${st.name}' to have been called once. Called ${st.calls.length} times.`).toBe(1);
        case "assertCalled":
          return () => expect(st.calls.length, `Expected '${st.name}' to have been called.`).toBeGreaterThan(0);
        case "assertNotCalled":
          return () => expect(st.calls.length, `Expected '${st.name}' to not have been called. Called ${st.calls.length} times.${callsHint(st)}`).toBe(0);
        case "assertAnyCall":
          return (...expected: any[]) => {
            const e = new _Call("", expected, st.sig);
            expect(st.calls.some((c) => c.equals(e)), `${e} call not found${callsHint(st)}`).toBe(true);
          };
        case "assertHasCalls":
          return (calls: any, anyOrder?: any) => {
            const exp: _Call[] = [...(calls ?? [])].map((c: any) => (c instanceof _Call ? new _Call(c.name, c.args, st.sig) : c));
            const ao = isKw(anyOrder) ? anyOrder.any_order ?? anyOrder.anyOrder : anyOrder;
            const actual = st.mockCalls;
            if (ao) {
              for (const e of exp) expect(actual.some((c) => c.equals(e)), `${e} not all found in ${new CallList(actual)}`).toBe(true);
            } else {
              let ok = false;
              for (let i = 0; i + exp.length <= actual.length && !ok; i++) ok = exp.every((e, j) => actual[i + j].equals(e));
              expect(ok || exp.length === 0, `Calls not found.\nExpected: ${new CallList(exp)}\nActual: ${new CallList(actual)}`).toBe(true);
            }
          };
        case "attachMockSig":
          return (fn: Function) => {
            st.sig = signatureOf(fn);
          };
      }
      if (PROTOCOL_ALIAS[prop] && (st.children.has(PROTOCOL_ALIAS[prop]) || !st.spec || !st.spec.has(prop))) {
        if (!st.spec || st.spec.has(prop) || st.children.has(PROTOCOL_ALIAS[prop]) || st.spec.has(PROTOCOL_ALIAS[prop]) || specHasProtocol(st, prop)) return child(PROTOCOL_ALIAS[prop]);
      }
      if (prop === "length") {
        if (st.children.has("__len__") || (st.spec && (st.spec.has("length") || specHasLen(st)))) {
          const r = child("__len__")();
          return isMock(r) ? 0 : r;
        }
        if (!st.spec) return isMock(st.children.get("__len__")) ? child("__len__")() : 0;
      }
      if (PASSTHROUGH.has(prop)) {
        if (prop === "toString") return () => `<Mock name='${st.name}'>`;
        if (prop === "valueOf") return () => proxy;
        if (prop === "constructor") return st.specProto ? (st.specProto as any).constructor : Object;
        return undefined;
      }
      if (prop === "prototype" && callable) return (t as any).prototype;
      if (st.spec && !st.spec.has(prop) && !prop.startsWith("__")) {
        // python raises AttributeError; JS code probes attributes, so answer `undefined`
        return undefined;
      }
      return child(prop);
    },
  };
  function child(name: string): any {
    let c = st.children.get(name);
    if (!c) {
      const member = st.specMembers?.get(name);
      if (member?.kind === "method" || member?.kind === "static") {
        c = makeMock({ name, parent: st, sigFn: member.fn });
        if (member.kind === "method") (c._mockState as MockState).sig = signatureOf(member.fn!);
      } else if (member?.kind === "class") c = makeMock({ name, parent: st, specClass: member.fn as Function });
      else if (member?.kind === "prop" && st.spec) c = makeMock({ name, parent: st, callable: true });
      else c = makeMock({ name, parent: st });
      st.children.set(name, c);
    }
    return c;
  }
  proxy = new Proxy(target, handler);
  if (o.props) for (const [k, v] of Object.entries(o.props)) proxy[k] = v;
  return proxy;
}

function specHasIter(st: MockState): boolean {
  const p: any = st.specProto;
  return !!p && typeof p[Symbol.iterator] === "function";
}
function specHasLen(st: MockState): boolean {
  const p: any = st.specProto;
  return !!p && (typeof p[Symbol.for("py.len")] === "function" || "length" in p);
}
function specHasProtocol(st: MockState, prop: string): boolean {
  const p: any = st.specProto;
  if (!p) return false;
  if (prop === "get") return typeof p[Symbol.for("py.getitem")] === "function";
  if (prop === "has") return typeof p[Symbol.for("py.contains")] === "function";
  return false;
}

function callsHint(st: MockState): string {
  return st.calls.length ? `\nCalls: ${new CallList(st.calls)}` : "";
}

function assertCallEq(st: MockState, actual: _Call, expected: any[]): void {
  const e = new _Call("", expected, st.sig);
  const a = new _Call("", actual.args, st.sig);
  if (!a.equals(e)) {
    expect.fail(`expected call not found.\nExpected: ${e}\nActual: ${a}`);
  }
}

/** A list of calls that compares equal to a list of `call(...)` objects. */
export class CallList extends Array<_Call> {
  constructor(calls: _Call[] | number = []) {
    super();
    if (Array.isArray(calls)) this.push(...calls);
  }
  static get [Symbol.species]() {
    return Array;
  }
  equals(other: unknown): boolean {
    const o = [...((other as any) ?? [])];
    return o.length === this.length && o.every((c, i) => this[i].equals(c));
  }
  toString(): string {
    return `[${this.map(String).join(", ")}]`;
  }
}

// ------------------------------------------------------------------------------------------------
// public constructors (unittest.mock)

/** `Mock(name=..., **attrs)` / `MagicMock(...)` */
export function Mock(opts: any = {}): any {
  const o = isKw(opts) || (opts && typeof opts === "object" && !isMock(opts)) ? { ...opts } : {};
  const name = o.name ?? "mock";
  delete o.name;
  const spec = o.spec ?? o.spec_set ?? o.specSet;
  delete o.spec;
  delete o.spec_set;
  delete o.specSet;
  const m = makeMock({ name, specClass: typeof spec === "function" ? spec : null, instance: typeof spec === "function" });
  for (const [k, v] of Object.entries(o)) {
    const parts = camel(k).split(".");
    let t: any = m;
    for (const p of parts.slice(0, -1)) t = t[p];
    t[parts[parts.length - 1]] = v;
  }
  return m;
}
export const MagicMock = Mock;
export const NonCallableMagicMock = (opts: any = {}) => Mock(opts);

/** `PropertyMock(return_value=...)`: reading the property calls it, assigning calls it with the value. */
export function PropertyMock(opts: any = {}): any {
  return Mock(opts);
}

/** `create_autospec(cls, instance=True, spec_set=True, **kw)` */
export function createAutospec(cls: Function, opts: any = {}): any {
  const o = { ...opts };
  const instance = o.instance ?? false;
  const name = o._name ?? o.name ?? "mock";
  for (const k of ["instance", "_name", "name", "spec_set", "specSet"]) delete o[k];
  const m = makeMock({ name, specClass: cls, instance });
  for (const [k, v] of Object.entries(o)) m[camel(k)] = v;
  return m;
}

// ------------------------------------------------------------------------------------------------
// patching

export interface FixtureRequest {
  fixturename: string;
  param?: any;
  addfinalizer(fn: () => void): void;
  getfixturevalue(name: string): any;
}

type Bindings = { names: string[]; get(n: string): any; set(n: string, v: any): boolean };
type ModuleMap = Record<string, any>;

let modules: ModuleMap = {};
/** Install the map of src modules (path -> module namespace) used to resolve dotted names. */
export function setModules(m: ModuleMap): void {
  modules = m;
}

function moduleFor(dotted: string): { mod: any; key: string } {
  // "pptx.opc.package.PartFactory" -> module "opc/package", name "PartFactory"
  const parts = dotted.split(".");
  if (parts[0] === "pptx") parts.shift();
  const name = parts.pop()!;
  const path = parts.join("/");
  const candidates = [`/src/${path}.ts`, `/src/${path}/index.ts`];
  for (const [k, m] of Object.entries(modules)) {
    if (candidates.some((c) => k.endsWith(c))) return { mod: m, key: name };
  }
  throw new Error(`patch: no module for '${dotted}'`);
}

function moduleById(id: string): any {
  for (const m of Object.values(modules)) if ((m as any).__pyBindings?.id === id) return m;
  return null;
}

/** Follow `import { name } from "./x.js"` chains to the module that defines `name`. */
function followImport(mod: any, name: string, depth = 0): { bindings: Bindings; name: string } | null {
  const b = mod?.__pyBindings;
  if (!b || depth > 8) return null;
  if (b.names.includes(name)) return { bindings: b, name };
  const imp = b.imports?.[name];
  if (!imp) return null;
  const [spec, orig] = imp;
  const base = b.id.slice(0, b.id.lastIndexOf("/"));
  const parts = `${base}/${spec}`.split("/");
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== ".") out.push(p);
  }
  const target = out.join("/").replace(/\.js$/, ".ts");
  const m = moduleById(target) ?? moduleById(target.replace(/\.ts$/, "/index.ts"));
  return m ? followImport(m, orig, depth + 1) : null;
}

/** Resolve `dotted` to the module binding that defines it (following imports by name). */
function resolveBinding(dotted: string): { bindings: Bindings; name: string } {
  const { mod, key } = moduleFor(dotted);
  for (const name of [key, camel(key)]) {
    const r = followImport(mod, name);
    if (r) return r;
  }
  throw new Error(`patch: cannot resolve '${dotted}'`);
}

function camelIfNeeded(mod: any, key: string): string {
  const b: Bindings | undefined = mod.__pyBindings;
  if (b?.names.includes(key) || key in mod) return key;
  const c = camel(key);
  return c;
}

function patchBinding(request: FixtureRequest, dotted: string, value: any): void {
  const { bindings, name } = resolveBinding(dotted);
  const orig = bindings.get(name);
  bindings.set(name, value);
  request.addfinalizer(() => bindings.set(name, orig));
}

function patchAttr(request: FixtureRequest, obj: any, attr: string, desc: PropertyDescriptor): void {
  const own = Object.getOwnPropertyDescriptor(obj, attr);
  Object.defineProperty(obj, attr, { configurable: true, ...desc });
  request.addfinalizer(() => {
    if (own) Object.defineProperty(obj, attr, own);
    else delete obj[attr];
  });
}

/** Where `name` lives on `cls`: a static member, or the prototype (instance member). */
function memberHost(cls: any, name: string): { host: any; isStatic: boolean } {
  for (let c = cls; c && c !== Function.prototype; c = Object.getPrototypeOf(c)) {
    if (Object.prototype.hasOwnProperty.call(c, name)) return { host: cls, isStatic: true };
  }
  return { host: cls.prototype, isStatic: false };
}

function lookupMember(cls: any, name: string): { fn: any; desc?: PropertyDescriptor } {
  for (let c = cls; c && c !== Function.prototype; c = Object.getPrototypeOf(c)) {
    const d = Object.getOwnPropertyDescriptor(c, name);
    if (d) return { fn: d.value, desc: d };
  }
  for (let p = cls.prototype; p && p !== Object.prototype; p = Object.getPrototypeOf(p)) {
    const d = Object.getOwnPropertyDescriptor(p, name);
    if (d) return { fn: d.value ?? d.get, desc: d };
  }
  return { fn: undefined };
}

const optsOf = (kwargs: any) => (kwargs && typeof kwargs === "object" ? { ...kwargs } : {});

function applyKwargs(m: any, kwargs: any): void {
  for (const [k, v] of Object.entries(optsOf(kwargs))) {
    if (k === "autospec" || k === "spec_set" || k === "specSet" || k === "name") continue;
    m[camel(k)] = v;
  }
}

/** Mock patching the class with qualified name `qClassName` (autospec'd). */
export function classMock(request: FixtureRequest, qClassName: string, kwargs?: any): any {
  const { bindings, name } = resolveBinding(qClassName);
  const cls = bindings.get(name);
  const opts = optsOf(kwargs);
  const m =
    opts.autospec === false || typeof cls !== "function"
      ? makeMock({ name, sigFn: typeof cls === "function" ? cls : null })
      : makeMock({ name, specClass: /^class[\s{]/.test(Function.prototype.toString.call(cls)) ? cls : null, sigFn: cls });
  if (typeof cls === "function" && !/^class[\s{]/.test(Function.prototype.toString.call(cls))) {
    // a function standing in for a python class (e.g. PartFactory): spec its static members
    (m._mockState as MockState).sig = signatureOf(cls);
  }
  applyKwargs(m, kwargs);
  patchBinding(request, qClassName, m);
  return m;
}

/** Mock for a class attribute `attrName` on `cls`. */
export function clsAttrMock(request: FixtureRequest, cls: any, attrName: string, name?: string | null, kwargs?: any): any {
  const attr = camelMember(cls, attrName);
  const m = makeMock({ name: name ?? request.fixturename });
  applyKwargs(m, kwargs);
  patchAttr(request, cls, attr, { value: m, writable: true });
  return m;
}

/** Mock patching the function with qualified name `qFunctionName`. */
export function functionMock(request: FixtureRequest, qFunctionName: string, kwargs?: any): any {
  const { bindings, name } = resolveBinding(qFunctionName);
  const fn = bindings.get(name);
  const m = makeMock({ name, sigFn: typeof fn === "function" ? fn : null });
  applyKwargs(m, kwargs);
  patchBinding(request, qFunctionName, m);
  return m;
}

/**
 * Mock for `cls.__init__`: constructing `cls` records the call (with the new instance first,
 * like python's autospec'd `__init__`) and leaves the instance uninitialized.
 */
export function initializerMock(request: FixtureRequest, cls: any, kwargs?: any): any {
  const m = makeMock({ name: "__init__", sigFn: cls });
  const st = m._mockState as MockState;
  st.sig = [{ name: "self" }, ...(signatureOf(cls) ?? [])];
  applyKwargs(m, kwargs);
  const proxyCls = new Proxy(cls, {
    construct(target, args, newTarget) {
      const self = Object.create((newTarget === proxyCls ? target : newTarget).prototype);
      m(self, ...args);
      return self;
    },
    apply(target, _this, args) {
      const self = Object.create(target.prototype);
      m(self, ...args);
      return self;
    },
  });
  // patch every module binding holding the class
  let patched = false;
  for (const mod of Object.values(modules)) {
    const b: Bindings | undefined = mod.__pyBindings;
    if (!b) continue;
    for (const n of b.names) {
      if (b.get(n) === cls) {
        b.set(n, proxyCls);
        request.addfinalizer(() => b.set(n, cls));
        patched = true;
      }
    }
  }
  if (!patched) throw new Error(`initializer_mock: ${cls.name} is not a module binding`);
  // subclasses defined with `extends cls` still construct through the real class; route
  // their super() calls by swapping the prototype link of direct subclasses
  return m;
}

/** Mock for an instance of `cls` drawing its spec from that class. */
export function instanceMock(request: FixtureRequest, cls: any, kwargs?: any): any {
  const opts = optsOf(kwargs);
  const m = makeMock({ name: opts.name ?? request.fixturename, specClass: cls, instance: true });
  applyKwargs(m, kwargs);
  return m;
}

/** A "loose" mock with no spec. */
export function looseMock(request: FixtureRequest, kwargs?: any): any {
  const opts = optsOf(kwargs);
  const m = makeMock({ name: opts.name ?? request.fixturename });
  applyKwargs(m, kwargs);
  return m;
}

/** JS spellings of python protocol methods mocked by name (`method_mock(request, X, "__iter__")`). */
const DUNDER_MEMBER: Record<string, string | symbol> = { __iter__: Symbol.iterator, __getitem__: Symbol.for("py.getitem"), __len__: Symbol.for("py.len") };

function camelMember(cls: any, name: string): any {
  if (lookupMember(cls, name).desc) return name;
  if (name in DUNDER_MEMBER) return DUNDER_MEMBER[name];
  const c = camel(name);
  return c;
}

/** Mock for method `methodName` on `cls` (an instance method receives `self` first). */
export function methodMock(request: FixtureRequest, cls: any, methodName: string, kwargs?: any): any {
  const name = camelMember(cls, methodName);
  const { fn } = lookupMember(cls, name);
  const { host, isStatic } = memberHost(cls, name);
  const opts = optsOf(kwargs);
  const autospec = opts.autospec !== false;
  const m = makeMock({ name, sigFn: fn });
  const st = m._mockState as MockState;
  if (!isStatic && autospec) st.sig = [{ name: "self" }, ...(signatureOf(fn) ?? [])];
  applyKwargs(m, kwargs);
  // an autospec'd instance method receives `self`; read through an instance it still exposes the
  // mock's API (`obj.method.assert_called_once_with(obj, ...)`), like python's autospec function
  const bound = function (this: any, ...args: any[]) {
    return m(this, ...args);
  };
  const value = isStatic || !autospec ? m : new Proxy(bound, { get: (t, p, r) => (p === "call" || p === "apply" || p === "bind" ? Reflect.get(t, p, r) : m[p]) });
  patchAttr(request, host, name, { value, writable: true });
  return m;
}

/** Mock for the builtin `open()`; not applicable in JS (files are read through src/fileio). */
export function openMock(_request: FixtureRequest, _moduleName: string): any {
  throw new Error("open_mock is not supported in the TS port");
}

/** Mock for property `propName` on class `cls`. */
export function propertyMock(request: FixtureRequest, cls: any, propName: string, kwargs?: any): any {
  const name = camelMember(cls, propName);
  const m = makeMock({ name });
  applyKwargs(m, kwargs);
  const { host } = memberHost(cls, name);
  patchAttr(request, host, name, {
    get() {
      return m();
    },
    set(v: any) {
      m(v);
    },
  });
  // lazy properties cache themselves on the instance; make sure no cached value shadows
  return m;
}

/** Mock patching the variable with qualified name `qVarName`. */
export function varMock(request: FixtureRequest, qVarName: string, kwargs?: any): any {
  const opts = optsOf(kwargs);
  if ("new" in opts) {
    patchBinding(request, qVarName, opts.new);
    return opts.new;
  }
  const m = makeMock({ name: qVarName.split(".").pop() });
  applyKwargs(m, kwargs);
  patchBinding(request, qVarName, m);
  return m;
}

/** `patch.object(obj, attr, new)` for use inside a test body; restored after the test. */
export function patchObject(request: FixtureRequest, obj: any, attr: string, value?: any): any {
  const name = camelMember(obj, attr);
  const v = value === undefined ? makeMock({ name }) : value;
  patchAttr(request, obj, name, { value: v, writable: true });
  return v;
}

export { AttributeError };
