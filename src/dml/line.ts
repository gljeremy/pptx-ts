// DrawingML objects related to line formatting. (`pptx.dml.line`)
import { lazy } from "../util.js";
import { FillFormat } from "./fill.js";
import { MSO_FILL } from "../enum/dml.js";
import { Emu } from "../util.js";

/**
 * Provides access to line properties such as color, style, and width.
 *
 * A LineFormat object is typically accessed via the ``.line`` property of
 * a shape such as |Shape| or |Picture|.
 */
export class LineFormat {
  declare _parent: any;

  constructor(parent: any) {
    this._parent = parent;
  }

  /**
   * The |ColorFormat| instance that provides access to the color settings
   * for this line. Essentially a shortcut for ``line.fill.fore_color``.
   * As a side-effect, accessing this property causes the line fill type
   * to be set to ``MSO_FILL.SOLID``. If this sounds risky for your use
   * case, use ``line.fill.type`` to non-destructively discover the
   * existing fill type.
   */
  get color() {
    return lazy(this, "color", () => {
      if (this.fill.type !== MSO_FILL.SOLID) {
        this.fill.solid();
      }
      return this.fill.foreColor;
    });
  }

  /**
   * Return value indicating line style.
   *
   * Returns a member of :ref:`MsoLineDashStyle` indicating line style, or
   * |None| if no explicit value has been set. When no explicit value has
   * been set, the line dash style is inherited from the style hierarchy.
   *
   * Assigning |None| removes any existing explicitly-defined dash style.
   */
  get dashStyle() {
    const ln = this._ln;
    if (ln === null) {
      return null;
    }
    return ln.prstDashVal;
  }

  set dashStyle(dashStyle: any) {
    let ln;
    if (dashStyle === null) {
      ln = this._ln;
      if (ln === null) {
        return;
      }
      ln._removePrstDash();
      ln._removeCustDash();
      return;
    }
    ln = this._getOrAddLn();
    ln.prstDashVal = dashStyle;
  }

  /**
   * |FillFormat| instance for this line, providing access to fill
   * properties such as foreground color.
   */
  get fill() {
    return lazy(this, "fill", () => {
      const ln = this._getOrAddLn();
      return FillFormat.fromFillParent(ln);
    });
  }

  /**
   * The width of the line expressed as an integer number of :ref:`English
   * Metric Units <EMU>`. The returned value is an instance of |Length|,
   * a value class having properties such as `.inches`, `.cm`, and `.pt`
   * for converting the value into convenient units.
   */
  get width() {
    const ln = this._ln;
    if (ln === null) {
      return Emu(0);
    }
    return ln.w;
  }

  set width(emu: any) {
    if (emu === null) {
      emu = 0;
    }
    const ln = this._getOrAddLn();
    ln.w = emu;
  }

  /**
   * Return the ``<a:ln>`` element containing the line format properties
   * in the XML.
   */
  _getOrAddLn() {
    return this._parent.getOrAddLn();
  }

  get _ln() {
    return this._parent.ln;
  }
}
