// Utility functions and classes (`pptx.util`), plus the small pieces of Python
// numeric semantics the port needs to produce identical XML.
import { ValueError } from "./errors.js";

const EMUS_PER_INCH = 914400;
const EMUS_PER_CENTIPOINT = 127;
const EMUS_PER_CM = 360000;
const EMUS_PER_MM = 36000;
const EMUS_PER_PT = 12700;

/**
 * A length in English Metric Units with unit-conversion properties.
 *
 * Lengths are plain numbers (EMU) everywhere in the pptx-ts API; wrap one in `Length` (or
 * `new Emu(x)`) to read it in other units: `new Length(shape.width).inches`.
 */
export class Length extends Number {
  constructor(emu: number) {
    super(Math.trunc(Number(emu)));
  }
  /** Floating point length in inches. */
  get inches(): number {
    return this.valueOf() / EMUS_PER_INCH;
  }
  /** Integer length in hundredths of a point (1/7200 inch). */
  get centipoints(): number {
    return Math.floor(this.valueOf() / EMUS_PER_CENTIPOINT);
  }
  /** Floating point length in centimeters. */
  get cm(): number {
    return this.valueOf() / EMUS_PER_CM;
  }
  /** Integer length in English Metric Units. */
  get emu(): number {
    return this.valueOf();
  }
  /** Floating point length in millimeters. */
  get mm(): number {
    return this.valueOf() / EMUS_PER_MM;
  }
  /** Floating point length in points. */
  get pt(): number {
    return this.valueOf() / EMUS_PER_PT;
  }
}

/** Length in EMU from inches. */
export const Inches = (inches: number): number => Math.trunc(inches * EMUS_PER_INCH);
/** Length in EMU from hundredths of a point. */
export const Centipoints = (centipoints: number): number => Math.trunc(centipoints * EMUS_PER_CENTIPOINT);
/** Length in EMU from centimeters. */
export const Cm = (cm: number): number => Math.trunc(cm * EMUS_PER_CM);
/** Length in EMU (truncated to an integer). */
export const Emu = (emu: number | string): number => Math.trunc(Number(emu));
/** Length in EMU from millimeters. */
export const Mm = (mm: number): number => Math.trunc(mm * EMUS_PER_MM);
/** Length in EMU from points. */
export const Pt = (points: number): number => Math.trunc(points * EMUS_PER_PT);

/**
 * Like a getter, but evaluated only on first access; the value is cached on the instance.
 * Use inside a getter: `get foo() { return lazy(this, "foo", () => compute()); }`.
 */
export function lazy<T>(obj: object, name: string, fget: () => T): T {
  const v = fget();
  Object.defineProperty(obj, name, { value: v, writable: false, configurable: true, enumerable: false });
  return v;
}

/** Python truthiness for collections: empty (len 0) is false; any other object is true. */
export function pyBool(v: unknown): boolean {
  if (v === null || v === undefined) return false;
  const len = (v as any)[Symbol.for("py.len")];
  if (typeof len === "function") return len.call(v) > 0;
  if (Array.isArray(v) || typeof v === "string") return v.length > 0;
  return Boolean(v);
}

// -- Python numeric semantics --------------------------------------------------------------------

/** Python `round()` (round half to even) to `ndigits` decimals. */
export function pyRound(x: number, ndigits = 0): number {
  if (!Number.isFinite(x)) return x;
  const m = 10 ** ndigits;
  const y = ndigits ? Number((x * m).toPrecision(15)) : x;
  const r = Math.round(y);
  const res = Math.abs(y % 1) === 0.5 ? 2 * Math.round(y / 2) : r;
  return ndigits ? res / m : res;
}

/** Python `%` (result takes the sign of the divisor). */
export function pyMod(a: number, b: number): number {
  const r = a % b;
  return r !== 0 && r < 0 !== b < 0 ? r + b : r;
}

/** Python `int(str)`: strict base-10 integer parse, ValueError otherwise. */
export function pyInt(s: string | number): number {
  if (typeof s === "number") return Math.trunc(s);
  if (typeof s !== "string" && !((s as unknown) instanceof String))
    throw new TypeError(`int() argument must be a string, a bytes-like object or a real number, not '${s === null || s === undefined ? "NoneType" : typeof s}'`);
  if (!/^\s*[+-]?\d+(_\d+)*\s*$/.test(s)) throw new ValueError(`invalid literal for int() with base 10: '${s}'`);
  return parseInt(s.replace(/_/g, ""), 10);
}

/** Python `float(str)`. */
export function pyFloat(s: string | number): number {
  if (typeof s === "number") return s;
  const t = s.trim().toLowerCase();
  if (/^[+-]?(inf|infinity)$/.test(t)) return t.startsWith("-") ? -Infinity : Infinity;
  if (/^[+-]?nan$/.test(t)) return NaN;
  if (!/^[+-]?((\d+(_\d+)*)?\.?\d*(_\d+)*)([eE][+-]?\d+)?$/.test(t) || !/\d/.test(t))
    throw new ValueError(`could not convert string to float: '${s}'`);
  return parseFloat(t.replace(/_/g, ""));
}

/** Python `str(float)`: "1.0", "1e+16", "1e-05", "0.1". */
export function pyFloatStr(x: number): string {
  if (Number.isNaN(x)) return "nan";
  if (x === Infinity) return "inf";
  if (x === -Infinity) return "-inf";
  if (Object.is(x, -0)) return "-0.0";
  const abs = Math.abs(x);
  if (abs !== 0 && (abs >= 1e16 || abs < 1e-4)) {
    let s = x.toExponential();
    // JS "1e+16" / "1.5e-7" -> python "1e+16" / "1.5e-07"
    s = s.replace(/e([+-])(\d)$/, "e$10$2");
    return s;
  }
  const s = String(x);
  return Number.isInteger(x) ? s + ".0" : s;
}

/** Python `str()` of a value (None/True/False/float formatting). */
export function pyStr(v: unknown): string {
  if (v === null || v === undefined) return "None";
  if (v === true) return "True";
  if (v === false) return "False";
  return String(v);
}

/** Python `repr()` of a str. */
export function pyReprStr(s: string): string {
  const q = s.includes("'") && !s.includes('"') ? '"' : "'";
  let out = "";
  for (const ch of s) {
    if (ch === "\\") out += "\\\\";
    else if (ch === q) out += "\\" + q;
    else if (ch === "\n") out += "\\n";
    else if (ch === "\r") out += "\\r";
    else if (ch === "\t") out += "\\t";
    else if (ch.charCodeAt(0) < 0x20 || ch.charCodeAt(0) === 0x7f) out += "\\x" + ch.charCodeAt(0).toString(16).padStart(2, "0");
    else out += ch;
  }
  return q + out + q;
}

/** Python `repr()` for simple values. */
export function pyRepr(v: unknown): string {
  if (typeof v === "string") return pyReprStr(v);
  if (Array.isArray(v)) return "[" + v.map(pyRepr).join(", ") + "]";
  return pyStr(v);
}

/** Python type name of a JS value, for error messages (`<class 'str'>`). */
export function pyTypeName(v: unknown): string {
  if (v === null || v === undefined) return "<class 'NoneType'>";
  if (typeof v === "string") return "<class 'str'>";
  if (typeof v === "boolean") return "<class 'bool'>";
  if (typeof v === "number") return Number.isInteger(v) ? "<class 'int'>" : "<class 'float'>";
  if (Array.isArray(v)) return "<class 'list'>";
  const name = (v as object).constructor?.name ?? "object";
  return `<class '${name}'>`;
}
