// Objects shared by pptx modules (`pptx.shared`).
import type { BaseOxmlElement } from "./oxml/xmlchemy.js";
import { PY_GETITEM, PY_LEN } from "./oxml/xml.js";

/** Anything that can supply the package part it belongs to. */
export interface ProvidesPart {
  readonly part: any;
}

/**
 * Base class for element proxy classes: objects whose responsibilities are fulfilled by
 * manipulating the attributes and child elements of an XML element.
 */
export class ElementProxy {
  _element: any;

  constructor(element: BaseOxmlElement) {
    this._element = element;
  }

  /** True if this proxy refers to the same XML element as `other` (python `==`). */
  equals(other: unknown): boolean {
    if (!(other instanceof ElementProxy)) return false;
    return this._element === other._element;
  }

  /** The XML element proxied by this object. */
  get element(): any {
    return this._element;
  }
}

/** Provides access to ancestor objects and part. */
export class ParentedElementProxy extends ElementProxy {
  _parent: any;

  constructor(element: BaseOxmlElement, parent: ProvidesPart) {
    super(element);
    this._parent = parent;
  }

  /** The ancestor proxy object to this one. */
  get parent(): any {
    return this._parent;
  }

  /** The package part containing this object. */
  get part(): any {
    return this._parent.part;
  }
}

/** Common members for proxy-objects that wrap a part's root element, e.g. `p:sld`. */
export class PartElementProxy extends ElementProxy {
  _part: any;

  constructor(element: BaseOxmlElement, part: any) {
    super(element);
    this._part = part;
  }

  /** The package part containing this object. */
  get part(): any {
    return this._part;
  }
}

/**
 * python `a == b` for element proxies: identity, or `a.equals(b)` returning true (two proxies of
 * the same element). Anything else, including null, compares unequal.
 */
export function proxyEquals(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  return a instanceof ElementProxy && a.equals(b) === true;
}

/** Python sequence equality: unwrap String labels by value, then use `equals()` or `===`. */
function itemEquals(a: any, b: unknown): boolean {
  if (a instanceof String) a = a.valueOf();
  if (b instanceof String) b = b.valueOf();
  return a !== null && typeof a === "object" && typeof a.equals === "function" ? a.equals(b) : a === b;
}

/**
 * python's `collections.abc.Sequence`: subclasses supply `at(idx)` and `length`, and get
 * iteration, `[PY_GETITEM]`/`[PY_LEN]` (python `seq[i]`/`len(seq)`) and `includes()` (`in`)
 * from them, as python derives them from `__getitem__`/`__len__`.
 */
export abstract class PySequence<T> implements Iterable<T> {
  abstract at(idx: number): T;
  abstract get length(): number;

  [PY_GETITEM](idx: number): T {
    return this.at(idx);
  }

  [PY_LEN](): number {
    return this.length;
  }

  *[Symbol.iterator](): Generator<T> {
    for (let i = 0; i < this.length; i++) yield this.at(i);
  }

  /** python `value in seq`. */
  includes(value: unknown): boolean {
    for (const v of this) if (itemEquals(v, value)) return true;
    return false;
  }
}
