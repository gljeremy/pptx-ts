// Objects shared by modules in the opc sub-package (`pptx.opc.shared`).
import { KeyError } from "../errors.js";
import { PY_CONTAINS, PY_GETITEM, PY_LEN } from "../oxml/xml.js";

/** A str-keyed mapping that matches keys without respect to case: D['A'] == D['a']. */
export class CaseInsensitiveDict<V = string> implements Iterable<string> {
  _d = new Map<string, V>();

  constructor(init?: Iterable<[string, V]> | Record<string, V>) {
    if (!init) return;
    const entries = Symbol.iterator in Object(init) ? (init as Iterable<[string, V]>) : Object.entries(init as Record<string, V>);
    for (const [k, v] of entries) this._d.set(String(k).toLowerCase(), v);
  }

  has(key: string | String): boolean {
    return this._d.has(String(key).toLowerCase());
  }

  get(key: string | String): V {
    const k = String(key).toLowerCase();
    if (!this._d.has(k)) throw new KeyError(k);
    return this._d.get(k)!;
  }

  set(key: string | String, value: V): void {
    this._d.set(String(key).toLowerCase(), value);
  }

  keys(): string[] {
    return [...this._d.keys()];
  }

  values(): V[] {
    return [...this._d.values()];
  }

  items(): [string, V][] {
    return [...this._d.entries()];
  }

  get length(): number {
    return this._d.size;
  }

  [Symbol.iterator](): Iterator<string> {
    return this._d.keys();
  }

  [PY_GETITEM](k: string): V {
    return this.get(k);
  }
  [PY_LEN](): number {
    return this._d.size;
  }
  [PY_CONTAINS](k: string): boolean {
    return this.has(k);
  }

  equals(other: unknown): boolean {
    const o = other instanceof CaseInsensitiveDict ? Object.fromEntries(other._d) : (other as Record<string, V>);
    if (!o || typeof o !== "object") return false;
    const keys = Object.keys(o);
    return keys.length === this._d.size && keys.every((k) => this._d.has(k) && this._d.get(k) === o[k]);
  }
}
