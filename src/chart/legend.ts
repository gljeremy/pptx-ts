// Legend of a chart. (`pptx.chart.legend`)
import { lazy, pyBool } from "../util.js";
import { XL_LEGEND_POSITION } from "../enum/chart.js";
import { Font } from "../text/text.js";

/** Represents the legend in a chart. A chart can have at most one legend. */
export class Legend {
  declare _element: any;

  constructor(legendElm: any) {
    this._element = legendElm;
  }

  /**
   * The |Font| object that provides access to the text properties for
   * this legend, such as bold, italic, etc.
   */
  get font() {
    return lazy(this, "font", () => {
      const defRPr = this._element.defRPr;
      const font = new Font(defRPr);
      return font;
    });
  }

  /**
   * Adjustment of the x position of the legend from its default.
   * Expressed as a float between -1.0 and 1.0 representing a fraction of
   * the chart width. Negative values move the legend left, positive
   * values move it to the right. |None| if no setting is specified.
   */
  get horzOffset() {
    return this._element.horzOffset;
  }

  set horzOffset(value: any) {
    this._element.horzOffset = value;
  }

  /**
   * |True| if legend should be located inside plot area.
   *
   * Read/write boolean specifying whether legend should be placed inside
   * the plot area. In many cases this will cause it to be superimposed on
   * the chart itself. Assigning |None| to this property causes any
   * `c:overlay` element to be removed, which is interpreted the same as
   * |True|. This use case should rarely be required and assigning
   * a boolean value is recommended.
   */
  get includeInLayout() {
    const overlay = this._element.overlay;
    if (overlay === null) {
      return true;
    }
    return overlay.val;
  }

  set includeInLayout(value: any) {
    if (value === null) {
      this._element._removeOverlay();
      return;
    }
    this._element.getOrAddOverlay().val = pyBool(value);
  }

  /**
   * Read/write :ref:`XlLegendPosition` enumeration value specifying the
   * general region of the chart in which to place the legend.
   */
  get position() {
    const legendPos = this._element.legendPos;
    if (legendPos === null) {
      return XL_LEGEND_POSITION.RIGHT;
    }
    return legendPos.val;
  }

  set position(position: any) {
    this._element.getOrAddLegendPos().val = position;
  }
}
