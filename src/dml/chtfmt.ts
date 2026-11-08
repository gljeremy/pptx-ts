// |ChartFormat| and related objects. (`pptx.dml.chtfmt`)
import { lazy } from "../util.js";
import { FillFormat } from "./fill.js";
import { LineFormat } from "./line.js";
import { ElementProxy } from "../shared.js";

/**
 * The |ChartFormat| object provides access to visual shape properties for
 * chart elements like |Axis|, |Series|, and |MajorGridlines|. It has two
 * properties, :attr:`fill` and :attr:`line`, which return a |FillFormat|
 * and |LineFormat| object respectively. The |ChartFormat| object is
 * provided by the :attr:`format` property on the target axis, series, etc.
 */
export class ChartFormat extends ElementProxy {

  /**
   * |FillFormat| instance for this object, providing access to fill
   * properties such as fill color.
   */
  get fill() {
    return lazy(this, "fill", () => {
      const spPr = this._element.getOrAddSpPr();
      return FillFormat.fromFillParent(spPr);
    });
  }

  /**
   * The |LineFormat| object providing access to the visual properties of
   * this object, such as line color and line style.
   */
  get line() {
    return lazy(this, "line", () => {
      const spPr = this._element.getOrAddSpPr();
      return new LineFormat(spPr);
    });
  }
}
