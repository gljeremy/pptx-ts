// Chart-related objects such as Chart and ChartTitle. (`pptx.chart.chart`)
import { ValueError } from "../errors.js";
import { lazy, pyBool, pyItem } from "../util.js";
import { CategoryAxis, DateAxis, ValueAxis } from "./axis.js";
import { Legend } from "./legend.js";
import { PlotFactory, PlotTypeInspector } from "./plot.js";
import { SeriesCollection } from "./series.js";
import { SeriesXmlRewriterFactory } from "./xmlwriter.js";
import { ChartFormat } from "../dml/chtfmt.js";
import { ElementProxy, PartElementProxy, PySequence } from "../shared.js";
import { Font, TextFrame } from "../text/text.js";

/** A chart object. */
export class Chart extends PartElementProxy {
  declare _chartSpace: any;

  constructor(chartSpace: any, chartPart: any) {
    super(chartSpace, chartPart);
    this._chartSpace = chartSpace;
  }

  /**
   * The category axis of this chart. In the case of an XY or Bubble
   * chart, this is the X axis. Raises |ValueError| if no category
   * axis is defined (as is the case for a pie chart, for example).
   */
  get categoryAxis() {
    const catAxLst = this._chartSpace.catAxLst;
    if (catAxLst.length) {
      return new CategoryAxis(catAxLst[0]);
    }
    const dateAxLst = this._chartSpace.dateAxLst;
    if (dateAxLst.length) {
      return new DateAxis(dateAxLst[0]);
    }
    const valAxLst = this._chartSpace.valAxLst;
    if (valAxLst.length) {
      return new ValueAxis(valAxLst[0]);
    }
    throw new ValueError("chart has no category axis");
  }

  /**
   * Read/write integer index of chart style used to format this chart.
   * Range is from 1 to 48. Value is |None| if no explicit style has been
   * assigned, in which case the default chart style is used. Assigning
   * |None| causes any explicit setting to be removed. The integer index
   * corresponds to the style's position in the chart style gallery in the
   * PowerPoint UI.
   */
  get chartStyle() {
    const style = this._chartSpace.style;
    if (style === null) {
      return null;
    }
    return style.val;
  }

  set chartStyle(value: any) {
    this._chartSpace._removeStyle();
    if (value === null) {
      return;
    }
    this._chartSpace._addStyle({ val: value });
  }

  /**
   * A |ChartTitle| object providing access to title properties.
   *
   * Calling this property is destructive in the sense it adds a chart
   * title element (`c:title`) to the chart XML if one is not already
   * present. Use :attr:`has_title` to test for presence of a chart title
   * non-destructively.
   */
  get chartTitle() {
    return new ChartTitle(this._element.getOrAddTitle());
  }

  /**
   * Member of :ref:`XlChartType` enumeration specifying type of this chart.
   *
   * If the chart has two plots, for example, a line plot overlayed on a bar plot,
   * the type reported is for the first (back-most) plot. Read-only.
   */
  get chartType() {
    const firstPlot = this.plots.at(0);
    return PlotTypeInspector.chartType(firstPlot);
  }

  /** Font object controlling text format defaults for this chart. */
  get font() {
    return lazy(this, "font", () => {
      const defRPr = this._chartSpace.getOrAddTxPr().pLst[0].getOrAddPPr().getOrAddDefRPr();
      return new Font(defRPr);
    });
  }

  /**
   * Read/write boolean, |True| if the chart has a legend. Assigning
   * |True| causes a legend to be added to the chart if it doesn't already
   * have one. Assigning False removes any existing legend definition
   * along with any existing legend settings.
   */
  get hasLegend() {
    return this._chartSpace.chart.hasLegend;
  }

  set hasLegend(value: any) {
    this._chartSpace.chart.hasLegend = pyBool(value);
  }

  /**
   * Read/write boolean, specifying whether this chart has a title.
   *
   * Assigning |True| causes a title to be added if not already present.
   * Assigning |False| removes any existing title along with its text and
   * settings.
   */
  get hasTitle() {
    const title = this._chartSpace.chart.title;
    if (title === null) {
      return false;
    }
    return true;
  }

  set hasTitle(value: any) {
    const chart = this._chartSpace.chart;
    if (pyBool(value) === false) {
      chart._removeTitle();
      const autoTitleDeleted = chart.getOrAddAutoTitleDeleted();
      autoTitleDeleted.val = true;
      return;
    }
    chart.getOrAddTitle();
  }

  /**
   * A |Legend| object providing access to the properties of the legend
   * for this chart.
   */
  get legend() {
    const legendElm = this._chartSpace.chart.legend;
    if (legendElm === null) {
      return null;
    }
    return new Legend(legendElm);
  }

  /**
   * The sequence of plots in this chart. A plot, called a *chart group*
   * in the Microsoft API, is a distinct sequence of one or more series
   * depicted in a particular charting type. For example, a chart having
   * a series plotted as a line overlaid on three series plotted as
   * columns would have two plots; the first corresponding to the three
   * column series and the second to the line series. Plots are sequenced
   * in the order drawn, i.e. back-most to front-most. Supports *len()*,
   * membership (e.g. ``p in plots``), iteration, slicing, and indexed
   * access (e.g. ``plot = plots[i]``).
   */
  get plots() {
    return lazy(this, "plots", () => {
      const plotArea = this._chartSpace.chart.plotArea;
      return new _Plots(plotArea, this);
    });
  }

  /**
   * Use the categories and series values in the |ChartData| object
   * *chart_data* to replace those in the XML and Excel worksheet for this
   * chart.
   */
  replaceData(chartData: any) {
    const rewriter = SeriesXmlRewriterFactory(this.chartType, chartData);
    rewriter.replaceSeriesData(this._chartSpace);
    this._workbook.updateFromXlsxBlob(chartData.xlsxBlob);
  }

  /**
   * A |SeriesCollection| object containing all the series in this
   * chart. When the chart has multiple plots, all the series for the
   * first plot appear before all those for the second, and so on. Series
   * within a plot have an explicit ordering and appear in that sequence.
   */
  get series() {
    return lazy(this, "series", () => {
      return new SeriesCollection(this._chartSpace.plotArea);
    });
  }

  /**
   * The |ValueAxis| object providing access to properties of the value
   * axis of this chart. Raises |ValueError| if the chart has no value
   * axis.
   */
  get valueAxis() {
    const valAxLst = this._chartSpace.valAxLst;
    if (valAxLst.length === 0) {
      throw new ValueError("chart has no value axis");
    }
    const idx = (valAxLst.length > 1 ? 1 : 0);
    return new ValueAxis(valAxLst[idx]);
  }

  /**
   * The |ChartWorkbook| object providing access to the Excel source data
   * for this chart.
   */
  get _workbook() {
    return this.part.chartWorkbook;
  }
}

/** Provides properties for manipulating a chart title. */
export class ChartTitle extends ElementProxy {
  declare _title: any;

  constructor(title: any) {
    super(title);
    this._title = title;
  }

  /**
   * |ChartFormat| object providing access to line and fill formatting.
   *
   * Return the |ChartFormat| object providing shape formatting properties
   * for this chart title, such as its line color and fill.
   */
  get format() {
    return lazy(this, "format", () => {
      return new ChartFormat(this._title);
    });
  }

  /**
   * Read/write Boolean specifying whether this title has a text frame.
   *
   * Return |True| if this chart title has a text frame, and |False|
   * otherwise. Assigning |True| causes a text frame to be added if not
   * already present. Assigning |False| causes any existing text frame to
   * be removed along with its text and formatting.
   */
  get hasTextFrame() {
    if (this._title.txRich === null) {
      return false;
    }
    return true;
  }

  set hasTextFrame(value: any) {
    if (pyBool(value) === false) {
      this._title._removeTx();
      return;
    }
    this._title.getOrAddTxRich();
  }

  /**
   * |TextFrame| instance for this chart title.
   *
   * Return a |TextFrame| instance allowing read/write access to the text
   * of this chart title and its text formatting properties. Accessing this
   * property is destructive in the sense it adds a text frame if one is
   * not present. Use :attr:`has_text_frame` to test for the presence of
   * a text frame non-destructively.
   */
  get textFrame() {
    const rich = this._title.getOrAddTxRich();
    return new TextFrame(rich, this as any);
  }
}

/**
 * The sequence of plots in a chart, such as a bar plot or a line plot. Most
 * charts have only a single plot. The concept is necessary when two chart
 * types are displayed in a single set of axes, like a bar plot with
 * a superimposed line plot.
 */
export class _Plots extends PySequence<any> {
  declare _plotArea: any;
  declare _chart: any;

  constructor(plotArea: any, chart: any) {
    super();
    this._plotArea = plotArea;
    this._chart = chart;
  }

  at(index: number): any {
    const xCharts = this._plotArea.xCharts;
    const xChart = pyItem(xCharts, index);
    return PlotFactory(xChart, this._chart);
  }

  /** python `plots[start:end]`: the plots in that range (negative indices count from the end). */
  slice(start?: number, end?: number): any[] {
    return [...this._plotArea.xCharts].slice(start, end).map((xChart: any) => PlotFactory(xChart, this._chart));
  }

  get length(): number {
    return this._plotArea.xCharts.length;
  }
}
