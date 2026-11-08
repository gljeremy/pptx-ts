// Plot-related objects. (`pptx.chart.plot`)
import { ValueError, KeyError, NotImplementedError } from "../errors.js";
import { lazy, pyBool, pyGetItem } from "../util.js";
import { Categories } from "./category.js";
import { DataLabels } from "./datalabel.js";
import { SeriesCollection } from "./series.js";
import { XL_CHART_TYPE as XL } from "../enum/chart.js";
import { qn } from "../oxml/ns.js";
import { ST_BarDir, ST_Grouping } from "../oxml/simpletypes.js";

/**
 * A distinct plot that appears in the plot area of a chart. A chart may
 * have more than one plot, in which case they appear as superimposed
 * layers, such as a line plot appearing on top of a bar chart.
 */
export class _BasePlot {
  declare _element: any;
  declare _chart: any;

  constructor(xChart: any, chart: any) {
    this._element = xChart;
    this._chart = chart;
  }

  /**
   * Returns a |category.Categories| sequence object containing
   * a |category.Category| object for each of the category labels
   * associated with this plot. The |category.Category| class derives from
   * ``str``, so the returned value can be treated as a simple sequence of
   * strings for the common case where all you need is the labels in the
   * order they appear on the chart. |category.Categories| provides
   * additional properties for dealing with hierarchical categories when
   * required.
   */
  get categories() {
    return lazy(this, "categories", () => {
      return new Categories(this._element);
    });
  }

  /** The |Chart| object containing this plot. */
  get chart() {
    return this._chart;
  }

  /**
   * |DataLabels| instance providing properties and methods on the
   * collection of data labels associated with this plot.
   */
  get dataLabels() {
    const dLbls = this._element.dLbls;
    if (dLbls === null) {
      throw new ValueError("plot has no data labels, set has_data_labels = True first");
    }
    return new DataLabels(dLbls);
  }

  /**
   * Read/write boolean, |True| if the series has data labels. Assigning
   * |True| causes data labels to be added to the plot. Assigning False
   * removes any existing data labels.
   */
  get hasDataLabels() {
    return this._element.dLbls !== null;
  }

  set hasDataLabels(value: any) {
    if (pyBool(value) === false) {
      this._element._removeDLbls();
    } else if (this._element.dLbls === null) {
      const dLbls = this._element._addDLbls();
      dLbls.showVal.val = true;
    }
  }

  /**
   * A sequence of |Series| objects representing the series in this plot,
   * in the order they appear in the plot.
   */
  get series() {
    return lazy(this, "series", () => {
      return new SeriesCollection(this._element);
    });
  }

  /**
   * Read/write boolean value specifying whether to use a different color
   * for each of the points in this plot. Only effective when there is
   * a single series; PowerPoint automatically varies color by series when
   * more than one series is present.
   */
  get varyByCategories() {
    const varyColors = this._element.varyColors;
    if (varyColors === null) {
      return true;
    }
    return varyColors.val;
  }

  set varyByCategories(value: any) {
    this._element.getOrAddVaryColors().val = pyBool(value);
  }
}

/** An area plot. */
export class AreaPlot extends _BasePlot {
}

/** A 3-dimensional area plot. */
export class Area3DPlot extends _BasePlot {
}

/** A bar chart-style plot. */
export class BarPlot extends _BasePlot {

  /**
   * Width of gap between bar(s) of each category, as an integer
   * percentage of the bar width. The default value for a new bar chart is
   * 150, representing 150% or 1.5 times the width of a single bar.
   */
  get gapWidth() {
    const gapWidth = this._element.gapWidth;
    if (gapWidth === null) {
      return 150;
    }
    return gapWidth.val;
  }

  set gapWidth(value: any) {
    const gapWidth = this._element.getOrAddGapWidth();
    gapWidth.val = value;
  }

  /**
   * Read/write int value in range -100..100 specifying a percentage of
   * the bar width by which to overlap adjacent bars in a multi-series bar
   * chart. Default is 0. A setting of -100 creates a gap of a full bar
   * width and a setting of 100 causes all the bars in a category to be
   * superimposed. A stacked bar plot has overlap of 100 by default.
   */
  get overlap() {
    const overlap = this._element.overlap;
    if (overlap === null) {
      return 0;
    }
    return overlap.val;
  }

  set overlap(value: any) {
    if (value === 0) {
      this._element._removeOverlap();
      return;
    }
    this._element.getOrAddOverlap().val = value;
  }
}

/** A bubble chart plot. */
export class BubblePlot extends _BasePlot {

  /**
   * An integer between 0 and 300 inclusive indicating the percentage of
   * the default size at which bubbles should be displayed. Assigning
   * |None| produces the same behavior as assigning `100`.
   */
  get bubbleScale() {
    const bubbleScale = this._element.bubbleScale;
    if (bubbleScale === null) {
      return 100;
    }
    return bubbleScale.val;
  }

  set bubbleScale(value: any) {
    const bubbleChart = this._element;
    bubbleChart._removeBubbleScale();
    if (value === null) {
      return;
    }
    const bubbleScale = bubbleChart._addBubbleScale();
    bubbleScale.val = value;
  }
}

/** An doughnut plot. */
export class DoughnutPlot extends _BasePlot {
}

/** A line chart-style plot. */
export class LinePlot extends _BasePlot {
}

/** A pie chart-style plot. */
export class PiePlot extends _BasePlot {
}

/** A radar-style plot. */
export class RadarPlot extends _BasePlot {
}

/** An XY (scatter) plot. */
export class XyPlot extends _BasePlot {
}

/**
 * Return an instance of the appropriate subclass of _BasePlot based on the
 * tagname of *xChart*.
 */
export function PlotFactory(xChart: any, chart: any) {
  let PlotCls;
  try {
    PlotCls = pyGetItem({ [qn("c:areaChart")]: AreaPlot, [qn("c:area3DChart")]: Area3DPlot, [qn("c:barChart")]: BarPlot, [qn("c:bubbleChart")]: BubblePlot, [qn("c:doughnutChart")]: DoughnutPlot, [qn("c:lineChart")]: LinePlot, [qn("c:pieChart")]: PiePlot, [qn("c:radarChart")]: RadarPlot, [qn("c:scatterChart")]: XyPlot }, xChart.tag);
  } catch (e) {
    if (!(e instanceof KeyError)) throw e;
    throw new ValueError(`unsupported plot type ${xChart.tag}`);
  }
  return new PlotCls(xChart, chart);
}

/**
 * "One-shot" service object that knows how to identify the type of a plot
 * as a member of the XL_CHART_TYPE enumeration.
 */
export class PlotTypeInspector {

  /**
   * Return the member of :ref:`XlChartType` that corresponds to the chart
   * type of *plot*.
   */
  static chartType(plot: any) {
    // python keys this on `plot.__class__.__name__`; class names don't survive minification,
    // so match the exact class instead (a subclass is not a match, as with the name lookup)
    const methods: [Function, (plot: any) => any][] = [
      [AreaPlot, this._differentiateAreaChartType],
      [Area3DPlot, this._differentiateArea3dChartType],
      [BarPlot, this._differentiateBarChartType],
      [BubblePlot, this._differentiateBubbleChartType],
      [DoughnutPlot, this._differentiateDoughnutChartType],
      [LinePlot, this._differentiateLineChartType],
      [PiePlot, this._differentiatePieChartType],
      [RadarPlot, this._differentiateRadarChartType],
      [XyPlot, this._differentiateXyChartType],
    ];
    const found = methods.find(([cls]) => plot?.constructor === cls);
    if (found === undefined) {
      throw new NotImplementedError(`chart_type() not implemented for ${plot?.constructor?.name}`);
    }
    return found[1].call(this, plot);
  }

  static _differentiateArea3dChartType(plot: any) {
    return pyGetItem({ [ST_Grouping.STANDARD]: XL.THREE_D_AREA, [ST_Grouping.STACKED]: XL.THREE_D_AREA_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.THREE_D_AREA_STACKED_100 }, plot._element.groupingVal);
  }

  static _differentiateAreaChartType(plot: any) {
    return pyGetItem({ [ST_Grouping.STANDARD]: XL.AREA, [ST_Grouping.STACKED]: XL.AREA_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.AREA_STACKED_100 }, plot._element.groupingVal);
  }

  static _differentiateBarChartType(plot: any) {
    const barChart = plot._element;
    if (barChart.barDir.val === ST_BarDir.BAR) {
      return pyGetItem({ [ST_Grouping.CLUSTERED]: XL.BAR_CLUSTERED, [ST_Grouping.STACKED]: XL.BAR_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.BAR_STACKED_100 }, barChart.groupingVal);
    }
    if (barChart.barDir.val === ST_BarDir.COL) {
      return pyGetItem({ [ST_Grouping.CLUSTERED]: XL.COLUMN_CLUSTERED, [ST_Grouping.STACKED]: XL.COLUMN_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.COLUMN_STACKED_100 }, barChart.groupingVal);
    }
    throw new ValueError(`invalid barChart.barDir value '${barChart.barDir.val}'`);
  }

  static _differentiateBubbleChartType(plot: any) {
    const firstBubble3D = (bubbleChart: any) => {
      const results = bubbleChart.xpath("c:ser/c:bubble3D");
      return (results.length ? results[0] : null);
    };
    const bubbleChart = plot._element;
    const bubble3D = firstBubble3D(bubbleChart);
    if (bubble3D === null) {
      return XL.BUBBLE;
    }
    if (bubble3D.val) {
      return XL.BUBBLE_THREE_D_EFFECT;
    }
    return XL.BUBBLE;
  }

  static _differentiateDoughnutChartType(plot: any) {
    const doughnutChart = plot._element;
    const explosion = doughnutChart.xpath("./c:ser/c:explosion");
    return (explosion.length ? XL.DOUGHNUT_EXPLODED : XL.DOUGHNUT);
  }

  static _differentiateLineChartType(plot: any) {
    const lineChart = plot._element;
    const hasLineMarkers = () => {
      const matches = lineChart.xpath("c:ser/c:marker/c:symbol[@val=\"none\"]");
      if (matches.length) {
        return false;
      }
      return true;
    };
    if (hasLineMarkers()) {
      return pyGetItem({ [ST_Grouping.STANDARD]: XL.LINE_MARKERS, [ST_Grouping.STACKED]: XL.LINE_MARKERS_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.LINE_MARKERS_STACKED_100 }, plot._element.groupingVal);
    } else {
      return pyGetItem({ [ST_Grouping.STANDARD]: XL.LINE, [ST_Grouping.STACKED]: XL.LINE_STACKED, [ST_Grouping.PERCENT_STACKED]: XL.LINE_STACKED_100 }, plot._element.groupingVal);
    }
  }

  static _differentiatePieChartType(plot: any) {
    const pieChart = plot._element;
    const explosion = pieChart.xpath("./c:ser/c:explosion");
    return (explosion.length ? XL.PIE_EXPLODED : XL.PIE);
  }

  static _differentiateRadarChartType(plot: any) {
    const radarChart = plot._element;
    const radarStyle = radarChart.xpath("c:radarStyle")[0].get("val");
    const noMarkers = () => {
      const matches = radarChart.xpath("c:ser/c:marker/c:symbol");
      if ((matches.length && matches[0].get("val") === "none")) {
        return true;
      }
      return false;
    };
    if (radarStyle === null) {
      return XL.RADAR;
    }
    if (radarStyle === "filled") {
      return XL.RADAR_FILLED;
    }
    if (noMarkers()) {
      return XL.RADAR;
    }
    return XL.RADAR_MARKERS;
  }

  static _differentiateXyChartType(plot: any) {
    const scatterChart = plot._element;
    const noLine = () => {
      return pyBool(scatterChart.xpath("c:ser/c:spPr/a:ln/a:noFill"));
    };
    const noMarkers = () => {
      const symbols = scatterChart.xpath("c:ser/c:marker/c:symbol");
      if ((symbols.length && symbols[0].get("val") === "none")) {
        return true;
      }
      return false;
    };
    const scatterStyle = scatterChart.xpath("c:scatterStyle")[0].get("val");
    if (scatterStyle === "lineMarker") {
      if (noLine()) {
        return XL.XY_SCATTER;
      }
      if (noMarkers()) {
        return XL.XY_SCATTER_LINES_NO_MARKERS;
      }
      return XL.XY_SCATTER_LINES;
    }
    if (scatterStyle === "smoothMarker") {
      if (noMarkers()) {
        return XL.XY_SCATTER_SMOOTH_NO_MARKERS;
      }
      return XL.XY_SCATTER_SMOOTH;
    }
    return XL.XY_SCATTER;
  }
}
