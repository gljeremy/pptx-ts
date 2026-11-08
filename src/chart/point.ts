// Data point-related objects. (`pptx.chart.point`)
import { IndexError } from "../errors.js";
import { PySequence } from "../shared.js";
import { lazy } from "../util.js";
import { DataLabel } from "./datalabel.js";
import { Marker } from "./marker.js";
import { ChartFormat } from "../dml/chtfmt.js";

/** Sequence providing access to the individual data points in a series. */
export abstract class _BasePoints extends PySequence<Point> {
  declare _element: any;
  declare _ser: any;

  constructor(ser: any) {
    super();
    this._element = ser;
    this._ser = ser;
  }

  at(idx: number): Point {
    if ((idx < 0 || idx >= this.length)) {
      throw new IndexError("point index out of range");
    }
    return new Point(this._ser, idx);
  }
}

/**
 * Sequence providing access to the individual data points in
 * a |BubbleSeries| object.
 */
export class BubblePoints extends _BasePoints {

  get length(): number {
    return Math.min(this._ser.xValPtCountVal, this._ser.yValPtCountVal, this._ser.bubbleSizePtCountVal);
  }
}

/**
 * Sequence providing access to individual |Point| objects, each
 * representing the visual properties of a data point in the specified
 * category series.
 */
export class CategoryPoints extends _BasePoints {

  get length(): number {
    return this._ser.catPtCountVal;
  }
}

/**
 * Provides access to the properties of an individual data point in
 * a series, such as the visual properties of its marker and the text and
 * font of its data label.
 */
export class Point {
  declare _element: any;
  declare _ser: any;
  declare _idx: any;

  constructor(ser: any, idx: any) {
    this._element = ser;
    this._ser = ser;
    this._idx = idx;
  }

  /** The |DataLabel| object representing the label on this data point. */
  get dataLabel() {
    return lazy(this, "dataLabel", () => {
      return new DataLabel(this._ser, this._idx);
    });
  }

  /**
   * The |ChartFormat| object providing access to the shape formatting
   * properties of this data point, such as line and fill.
   */
  get format() {
    return lazy(this, "format", () => {
      const dPt = this._ser.getOrAddDPtForPoint(this._idx);
      return new ChartFormat(dPt);
    });
  }

  /**
   * The |Marker| instance for this point, providing access to the visual
   * properties of the data point marker, such as fill and line. Setting
   * these properties overrides any value set at the series level.
   */
  get marker() {
    return lazy(this, "marker", () => {
      const dPt = this._ser.getOrAddDPtForPoint(this._idx);
      return new Marker(dPt);
    });
  }
}

/**
 * Sequence providing access to the individual data points in an |XySeries|
 * object.
 */
export class XyPoints extends _BasePoints {

  get length(): number {
    return Math.min(this._ser.xValPtCountVal, this._ser.yValPtCountVal);
  }
}
