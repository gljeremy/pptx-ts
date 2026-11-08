// Visual effects on a shape such as shadow, glow, and reflection. (`pptx.dml.effect`)
import { pyBool } from "../util.js";

/** Provides access to shadow effect on a shape. */
export class ShadowFormat {
  declare _element: any;

  constructor(spPr: any) {
    this._element = spPr;
  }

  /**
   * True if shape inherits shadow settings.
   *
   * Read/write. An explicitly-defined shadow setting on a shape causes
   * this property to return |False|. A shape with no explicitly-defined
   * shadow setting inherits its shadow settings from the style hierarchy
   * (and so returns |True|).
   *
   * Assigning |True| causes any explicitly-defined shadow setting to be
   * removed and inheritance is restored. Note this has the side-effect of
   * removing **all** explicitly-defined effects, such as glow and
   * reflection, and restoring inheritance for all effects on the shape.
   * Assigning |False| causes the inheritance link to be broken and **no**
   * effects to appear on the shape.
   */
  get inherit() {
    if (this._element.effectLst === null) {
      return true;
    }
    return false;
  }

  set inherit(value: any) {
    const inherit = pyBool(value);
    if (inherit) {
      this._element._removeEffectLst();
    } else {
      this._element.getOrAddEffectLst();
    }
  }
}
