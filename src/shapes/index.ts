// Objects used across sub-package. (`pptx.shapes.__init__`)
import type { XmlPart } from "../opc/package.js";
import type { ProvidesPart } from "../types.js";

/**
 * Provides access to the containing part for drawing elements that occur below a shape.
 *
 * Access to the part is required for example to add or drop a relationship. Provides
 * `self._parent` attribute to subclasses.
 */
export class Subshape {
  declare _parent: any;

  constructor(parent: ProvidesPart) {
    this._parent = parent;
  }

  /** The package part containing this object. */
  get part(): XmlPart {
    return this._parent.part;
  }
}
