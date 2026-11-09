// Marker-related objects. (`pptx.chart.marker`)
import { lazy } from "../util.js";
import { ChartFormat } from "../dml/chtfmt.js";
import { ElementProxy } from "../shared.js";

/**
 * Represents a data point marker, such as a diamond or circle, on
 * a line-type chart.
 */
export class Marker extends ElementProxy {

  /**
   * The |ChartFormat| instance for this marker, providing access to shape
   * properties such as fill and line.
   */
  get format() {
    return lazy(this, "format", () => {
      const marker = this._element.getOrAddMarker();
      return new ChartFormat(marker);
    });
  }

  /**
   * An integer between 2 and 72 inclusive indicating the size of this
   * marker in points. A value of |None| indicates no explicit value is
   * set and the size is inherited from a higher-level setting or the
   * PowerPoint default (which may be 9). Assigning |None| removes any
   * explicitly assigned size, causing this value to be inherited.
   */
  get size() {
    const marker = this._element.marker;
    if (marker === null) {
      return null;
    }
    return marker.sizeVal;
  }

  set size(value: any) {
    const marker = this._element.getOrAddMarker();
    marker._removeSize();
    if (value === null) {
      return;
    }
    const size = marker._addSize();
    size.val = value;
  }

  /**
   * A member of the :ref:`XlMarkerStyle` enumeration indicating the shape
   * of this marker. Returns |None| if no explicit style has been set,
   * which corresponds to the "Automatic" option in the PowerPoint UI.
   */
  get style() {
    const marker = this._element.marker;
    if (marker === null) {
      return null;
    }
    return marker.symbolVal;
  }

  set style(value: any) {
    const marker = this._element.getOrAddMarker();
    marker._removeSymbol();
    if (value === null) {
      return;
    }
    const symbol = marker._addSymbol();
    symbol.val = value;
  }
}
