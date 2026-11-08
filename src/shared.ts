// Objects shared by pptx modules (`pptx.shared`).
import type { BaseOxmlElement } from "./oxml/xmlchemy.js";

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
