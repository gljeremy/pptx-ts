// Series-related objects. (`pptx.chart.series`)
import { KeyError, NotImplementedError } from "../errors.js";
import { PySequence } from "../shared.js";
import { lazy, pyGetItem, pyItem } from "../util.js";
import { DataLabels } from "./datalabel.js";
import { Marker } from "./marker.js";
import { BubblePoints, CategoryPoints, XyPoints } from "./point.js";
import { ChartFormat } from "../dml/chtfmt.js";
import { qn } from "../oxml/ns.js";

/** Base class for |BarSeries| and other series classes. */
export class _BaseSeries {
  declare _element: any;
  declare _ser: any;

  constructor(ser: any) {
    this._element = ser;
    this._ser = ser;
  }

  /**
   * The |ChartFormat| instance for this series, providing access to shape
   * properties such as fill and line.
   */
  get format() {
    return lazy(this, "format", () => {
      return new ChartFormat(this._ser);
    });
  }

  /**
   * The zero-based integer index of this series as reported in its
   * `c:ser/c:idx` element.
   */
  get index() {
    return this._element.idx.val;
  }

  /**
   * The string label given to this series, appears as the title of the
   * column for this series in the Excel worksheet. It also appears as the
   * label for this series in the legend.
   */
  get name() {
    const names = this._element.xpath("./c:tx//c:pt/c:v/text()");
    const name = (names.length ? names[0] : "");
    return name;
  }
}

/** Base class for |BarSeries| and other category chart series classes. */
export class _BaseCategorySeries extends _BaseSeries {

  /** |DataLabels| object controlling data labels for this series. */
  get dataLabels() {
    return lazy(this, "dataLabels", () => {
      return new DataLabels(this._ser.getOrAddDLbls());
    });
  }

  /**
   * The |CategoryPoints| object providing access to individual data
   * points in this series.
   */
  get points() {
    return lazy(this, "points", () => {
      return new CategoryPoints(this._ser);
    });
  }

  /**
   * Read-only. A sequence containing the float values for this series, in
   * the order they appear on the chart.
   */
  get values() {
    let idx;
    const iterValues = function* (this: any, ) {
      const val = this._element.val;
      if (val === null) {
        return;
      }
      for (let idx = 0; idx < val.ptCountVal; idx++) {
        yield val.ptV(idx);
      }
    }.bind(this);
    return [...iterValues()];
  }
}

/**
 * Mixin class providing `.marker` property for line-type chart series. The
 * line-type charts are Line, XY, and Radar.
 */
//
// python's second base class is a class-factory mixin here (as `_InheritsDimensions` is for
// placeholders); `x instanceof _MarkerMixin` still works.
const MARKER_MIXIN = Symbol("pptx.MarkerMixin");

type Ctor<T = object> = abstract new (...args: any[]) => T;

export function _MarkerMixin<TBase extends Ctor<{ _ser: any }>>(Base: TBase) {
  abstract class MarkerMixin extends Base {
    get [MARKER_MIXIN](): true {
      return true;
    }

    /**
     * The |Marker| instance for this series, providing access to data point
     * marker properties such as fill and line. Setting these properties
     * determines the appearance of markers for all points in this series
     * that are not overridden by settings at the point level.
     */
    get marker(): Marker {
      return lazy(this, "marker", () => {
        return new Marker(this._ser);
      });
    }
  }
  return MarkerMixin;
}
Object.defineProperty(_MarkerMixin, Symbol.hasInstance, {
  value: (v: unknown) => typeof v === "object" && v !== null && (v as any)[MARKER_MIXIN] === true,
});

/** A data point series belonging to an area plot. */
export class AreaSeries extends _BaseCategorySeries {
}

/** A data point series belonging to a bar plot. */
export class BarSeries extends _BaseCategorySeries {

  /**
   * |True| if a point having a value less than zero should appear with a
   * fill different than those with a positive value. |False| if the fill
   * should be the same regardless of the bar's value. When |True|, a bar
   * with a solid fill appears with white fill; in a bar with gradient
   * fill, the direction of the gradient is reversed, e.g. dark -> light
   * instead of light -> dark. The term "invert" here should be understood
   * to mean "invert the *direction* of the *fill gradient*".
   */
  get invertIfNegative() {
    const invertIfNegative = this._element.invertIfNegative;
    if (invertIfNegative === null) {
      return true;
    }
    return invertIfNegative.val;
  }

  set invertIfNegative(value: any) {
    const invertIfNegative = this._element.getOrAddInvertIfNegative();
    invertIfNegative.val = value;
  }
}

/** A data point series belonging to a line plot. */
export class LineSeries extends _MarkerMixin(_BaseCategorySeries) {

  /**
   * Read/write boolean specifying whether to use curve smoothing to
   * form the line connecting the data points in this series into
   * a continuous curve. If |False|, a series of straight line segments
   * are used to connect the points.
   */
  get smooth() {
    const smooth = this._element.smooth;
    if (smooth === null) {
      return true;
    }
    return smooth.val;
  }

  set smooth(value: any) {
    this._element.getOrAddSmooth().val = value;
  }
}

/** A data point series belonging to a pie plot. */
export class PieSeries extends _BaseCategorySeries {
}

/** A data point series belonging to a radar plot. */
export class RadarSeries extends _MarkerMixin(_BaseCategorySeries) {
}

/** A data point series belonging to an XY (scatter) plot. */
export class XySeries extends _MarkerMixin(_BaseSeries) {

  /**
   * Generate each float Y value in this series, in the order they appear
   * on the chart. A value of `None` represents a missing Y value
   * (corresponding to a blank Excel cell).
   */
  *iterValues() {
    let idx;
    const yVal = this._element.yVal;
    if (yVal === null) {
      return;
    }
    for (let idx = 0; idx < yVal.ptCountVal; idx++) {
      yield yVal.ptV(idx);
    }
  }

  /**
   * The |XyPoints| object providing access to individual data points in
   * this series.
   */
  get points() {
    return lazy(this, "points", () => {
      return new XyPoints(this._ser);
    });
  }

  /**
   * Read-only. A sequence containing the float values for this series, in
   * the order they appear on the chart.
   */
  get values() {
    return [...this.iterValues()];
  }
}

/** A data point series belonging to a bubble plot. */
export class BubbleSeries extends XySeries {

  /**
   * The |BubblePoints| object providing access to individual data point
   * objects used to discover and adjust the formatting and data labels of
   * a data point.
   */
  get points() {
    return lazy(this, "points", () => {
      return new BubblePoints(this._ser);
    });
  }
}

/** A sequence of |Series| objects. */
export class SeriesCollection extends PySequence<any> {
  declare _element: any;

  constructor(parentElm: any) {
    super();
    this._element = parentElm;
  }

  at(index: number): any {
    const ser = pyItem(this._element.sers, index);
    return _SeriesFactory(ser);
  }

  get length(): number {
    return this._element.sers.length;
  }
}

/**
 * Return an instance of the appropriate subclass of _BaseSeries based on the
 * xChart element *ser* appears in.
 */
export function _SeriesFactory(ser: any) {
  const xChartTag = ser.getparent().tag;
  let SeriesCls;
  try {
    SeriesCls = pyGetItem({ [qn("c:areaChart")]: AreaSeries, [qn("c:barChart")]: BarSeries, [qn("c:bubbleChart")]: BubbleSeries, [qn("c:doughnutChart")]: PieSeries, [qn("c:lineChart")]: LineSeries, [qn("c:pieChart")]: PieSeries, [qn("c:radarChart")]: RadarSeries, [qn("c:scatterChart")]: XySeries }, xChartTag);
  } catch (e) {
    if (!(e instanceof KeyError)) throw e;
    throw new NotImplementedError(`series class for ${xChartTag} not yet implemented`);
  }
  return new SeriesCls(ser);
}
