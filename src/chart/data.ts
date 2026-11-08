// ChartData and related objects. (`pptx.chart.data`)
import { ValueError, NotImplementedError } from "../errors.js";
import { PySequence } from "../shared.js";
import { encodeUtf8 } from "../oxml/xml.js";
import { lazy, pyFormatValue, pyItem } from "../util.js";
import { BubbleWorkbookWriter, CategoryWorkbookWriter, XyWorkbookWriter } from "./xlsx.js";
import { ChartXmlWriter } from "./xmlwriter.js";

/**
 * Base class providing common members for chart data objects.
 *
 * A chart data object serves as a proxy for the chart data table that will be written to an
 * Excel worksheet; operating as a sequence of series as well as providing access to chart-level
 * attributes. A chart data object is used as a parameter in :meth:`shapes.add_chart` and
 * :meth:`Chart.replace_data`. The data structure varies between major chart categories such as
 * category charts and XY charts.
 */
export class _BaseChartData extends PySequence<any> {
  declare _numberFormat: any;
  declare _series: any;

  constructor(numberFormat: any = "General") {
    super();
    this._numberFormat = numberFormat;
    this._series = [];
  }

  at(index: number): any {
    return pyItem(this._series, index);
  }

  get length(): number {
    return this._series.length;
  }

  append(series: any): void {
    this._series.push(series);
  }

  /**
   * The total integer number of data points appearing in the series of
   * this chart that are prior to *series* in this sequence.
   */
  dataPointOffset(series: any) {
    let count, thisSeries;
    count = 0;
    for (const thisSeries of this) {
      if (series === thisSeries) {
        return count;
      }
      count += thisSeries.length;
    }
    throw new ValueError("series not in chart data object");
  }

  /**
   * The formatting template string, e.g. '#,##0.0', that determines how
   * X and Y values are formatted in this chart and in the Excel
   * spreadsheet. A number format specified on a series will override this
   * value for that series. Likewise, a distinct number format can be
   * specified for a particular data point within a series.
   */
  get numberFormat() {
    return this._numberFormat;
  }

  /** Return the integer index of *series* in this sequence. */
  seriesIndex(series: any) {
    let idx, s;
    for (const [idx, s] of [...this].entries()) {
      if (series === s) {
        return idx;
      }
    }
    throw new ValueError("series not in chart data object");
  }

  /**
   * Return the Excel worksheet reference to the cell containing the name
   * for *series*.
   */
  seriesNameRef(series: any) {
    return this._workbookWriter.seriesNameRef(series);
  }

  /**
   * The Excel worksheet reference to the X values for *series* (not
   * including the column label).
   */
  xValuesRef(series: any) {
    return this._workbookWriter.xValuesRef(series);
  }

  /**
   * Return a blob containing an Excel workbook file populated with the
   * contents of this chart data object.
   */
  get xlsxBlob() {
    return this._workbookWriter.xlsxBlob;
  }

  /**
   * Return a blob containing the XML for a chart of *chart_type*
   * containing the series in this chart data object, as bytes suitable
   * for writing directly to a file.
   */
  xmlBytes(chartType: any): Uint8Array {
    return encodeUtf8(this._xml(chartType));
  }

  /**
   * The Excel worksheet reference to the Y values for *series* (not
   * including the column label).
   */
  yValuesRef(series: any) {
    return this._workbookWriter.yValuesRef(series);
  }

  /**
   * The worksheet writer object to which layout and writing of the Excel
   * worksheet for this chart will be delegated.
   */
  get _workbookWriter(): any {
    throw new NotImplementedError("must be implemented by all subclasses");
  }

  /**
   * Return (as unicode text) the XML for a chart of *chart_type*
   * populated with the values in this chart data object. The XML is
   * a complete XML document, including an XML declaration specifying
   * UTF-8 encoding.
   */
  _xml(chartType: any): string {
    return ChartXmlWriter(chartType, this).xml;
  }
}

/**
 * Base class providing common members for series data objects. A series
 * data object serves as proxy for a series data column in the Excel
 * worksheet. It operates as a sequence of data points, as well as providing
 * access to series-level attributes like the series label.
 */
export class _BaseSeriesData extends PySequence<any> {
  declare _chartData: any;
  declare _name: any;
  declare _numberFormat: any;
  declare _dataPoints: any;

  constructor(chartData: any, name: any, numberFormat: any) {
    super();
    this._chartData = chartData;
    this._name = name;
    this._numberFormat = numberFormat;
    this._dataPoints = [];
  }

  at(index: number): any {
    return pyItem(this._dataPoints, index);
  }

  get length(): number {
    return this._dataPoints.length;
  }

  append(dataPoint: any): void {
    this._dataPoints.push(dataPoint);
  }

  /**
   * The integer count of data points that appear in all chart series
   * prior to this one.
   */
  get dataPointOffset() {
    return this._chartData.dataPointOffset(this);
  }

  /**
   * Zero-based integer indicating the sequence position of this series in
   * its chart. For example, the second of three series would return `1`.
   */
  get index() {
    return this._chartData.seriesIndex(this);
  }

  /**
   * The name of this series, e.g. 'Series 1'. This name is used as the
   * column heading for the y-values of this series and may also appear in
   * the chart legend and perhaps other chart locations.
   */
  get name() {
    return (this._name !== null ? this._name : "");
  }

  /**
   * The Excel worksheet reference to the cell containing the name for
   * this series.
   */
  get nameRef() {
    return this._chartData.seriesNameRef(this);
  }

  /**
   * The formatting template string that determines how a number in this
   * series is formatted, both in the chart and in the Excel spreadsheet;
   * for example '#,##0.0'. If not specified for this series, it is
   * inherited from the parent chart data object.
   */
  get numberFormat() {
    const numberFormat = this._numberFormat;
    if (numberFormat === null) {
      return this._chartData.numberFormat;
    }
    return numberFormat;
  }

  /**
   * A sequence containing the X value of each datapoint in this series,
   * in data point order.
   */
  get xValues() {
    return [...this._dataPoints].map((dp) => dp.x);
  }

  /**
   * The Excel worksheet reference to the X values for this chart (not
   * including the column heading).
   */
  get xValuesRef() {
    return this._chartData.xValuesRef(this);
  }

  /**
   * A sequence containing the Y value of each datapoint in this series,
   * in data point order.
   */
  get yValues() {
    return [...this._dataPoints].map((dp) => dp.y);
  }

  /**
   * The Excel worksheet reference to the Y values for this chart (not
   * including the column heading).
   */
  get yValuesRef() {
    return this._chartData.yValuesRef(this);
  }
}

/** Base class providing common members for data point objects. */
export class _BaseDataPoint {
  declare _seriesData: any;
  declare _numberFormat: any;

  constructor(seriesData: any, numberFormat: any) {
    this._seriesData = seriesData;
    this._numberFormat = numberFormat;
  }

  /**
   * The formatting template string that determines how the value of this
   * data point is formatted, both in the chart and in the Excel
   * spreadsheet; for example '#,##0.0'. If not specified for this data
   * point, it is inherited from the parent series data object.
   */
  get numberFormat() {
    const numberFormat = this._numberFormat;
    if (numberFormat === null) {
      return this._seriesData.numberFormat;
    }
    return numberFormat;
  }
}

/**
 * Accumulates data specifying the categories and series values for a chart
 * and acts as a proxy for the chart data table that will be written to an
 * Excel worksheet. Used as a parameter in :meth:`shapes.add_chart` and
 * :meth:`Chart.replace_data`.
 *
 * This object is suitable for use with category charts, i.e. all those
 * having a discrete set of label values (categories) as the range of their
 * independent variable (X-axis) values. Unlike the ChartData types for
 * charts supporting a continuous range of independent variable values (such
 * as XyChartData), CategoryChartData has a single collection of category
 * (X) values and each data point in its series specifies only the Y value.
 * The corresponding X value is inferred by its position in the sequence.
 */
export class CategoryChartData extends _BaseChartData {
  declare _categories: any;

  /**
   * Return a newly created |data.Category| object having *label* and
   * appended to the end of the category collection for this chart.
   * *label* can be a string, a number, a datetime.date, or
   * datetime.datetime object. All category labels in a chart must be the
   * same type. All category labels in a chart having multi-level
   * categories must be strings.
   */
  addCategory(label: any) {
    return this.categories.addCategory(label);
  }

  /**
   * Add a series to this data set entitled *name* and having the data
   * points specified by *values*, an iterable of numeric values.
   * *number_format* specifies how the series values will be displayed,
   * and may be a string, e.g. '#,##0' corresponding to an Excel number
   * format.
   */
  addSeries(name: any, values: any = [], numberFormat: any = null) {
    let value;
    const seriesData = new CategorySeriesData(this, name, numberFormat);
    this.append(seriesData);
    for (const value of values) {
      seriesData.addDataPoint(value);
    }
    return seriesData;
  }

  /**
   * |data.Categories| object providing access to category-object hierarchy.
   *
   * Assigning an iterable of category labels (strings, numbers, or dates) replaces
   * the |data.Categories| object with a new one containing a category for each label
   * in the sequence.
   *
   * Creating a chart from chart data having date categories will cause the chart to
   * have a |DateAxis| for its category axis.
   */
  get categories() {
    // python: `if not getattr(self, "_categories", False)`, so an empty Categories is replaced too
    if (this._categories === undefined || this._categories.length === 0) {
      this._categories = new Categories();
    }
    return this._categories;
  }

  set categories(categoryLabels: any) {
    const categories = new Categories();
    for (const label of categoryLabels) {
      categories.addCategory(label);
    }
    this._categories = categories;
  }

  /**
   * The Excel worksheet reference to the categories for this chart (not
   * including the column heading).
   */
  get categoriesRef() {
    return this._workbookWriter.categoriesRef;
  }

  /**
   * The Excel worksheet reference to the values for *series* (not
   * including the column heading).
   */
  valuesRef(series: any) {
    return this._workbookWriter.valuesRef(series);
  }

  /**
   * The worksheet writer object to which layout and writing of the Excel
   * worksheet for this chart will be delegated.
   */
  override get _workbookWriter(): any {
    return lazy(this, "_workbookWriter", () => {
      return new CategoryWorkbookWriter(this);
    });
  }
}

/**
 * A sequence of |data.Category| objects, also having certain hierarchical
 * graph behaviors for support of multi-level (nested) categories.
 */
export class Categories extends PySequence<any> {
  declare _categories: any;
  declare _numberFormat: any;

  constructor() {
    super();
    this._categories = [];
    this._numberFormat = null;
  }

  at(idx: number): any {
    return pyItem(this._categories, idx);
  }

  /**
   * Return the count of the highest level of category in this sequence.
   * If it contains hierarchical (multi-level) categories, this number
   * will differ from :attr:`category_count`, which is the number of leaf
   * nodes.
   */
  get length(): number {
    return this._categories.length;
  }

  /**
   * Return a newly created |data.Category| object having *label* and
   * appended to the end of this category sequence. *label* can be
   * a string, a number, a datetime.date, or datetime.datetime object. All
   * category labels in a chart must be the same type. All category labels
   * in a chart having multi-level categories must be strings.
   *
   * Creating a chart from chart data having date categories will cause
   * the chart to have a |DateAxis| for its category axis.
   */
  addCategory(label: any) {
    const category = new Category(label, this);
    this._categories.push(category);
    return category;
  }

  /**
   * Return |True| if the first category in this collection has a date
   * label (as opposed to str or numeric). A date label is one of type
   * datetime.date or datetime.datetime. Returns |False| otherwise,
   * including when this category collection is empty. It also returns
   * False when this category collection is hierarchical, because
   * hierarchical categories can only be written as string labels.
   */
  get areDates() {
    if (this.depth !== 1) {
      return false;
    }
    const firstCatLabel = this.at(0).label;
    return firstCatLabel instanceof Date;
  }

  /**
   * Return |True| if the first category in this collection has a numeric
   * label (as opposed to a string label), including if that value is
   * a datetime.date or datetime.datetime object (as those are converted
   * to integers for storage in Excel). Returns |False| otherwise,
   * including when this category collection is empty. It also returns
   * False when this category collection is hierarchical, because
   * hierarchical categories can only be written as string labels.
   */
  get areNumeric() {
    if (this.depth !== 1) {
      return false;
    }
    const firstCatLabel = this.at(0).label;
    return typeof firstCatLabel === "number" || firstCatLabel instanceof Number || firstCatLabel instanceof Date;
  }

  /**
   * The number of hierarchy levels in this category graph. Returns 0 if
   * it contains no categories.
   */
  get depth(): number {
    const categories = this._categories;
    if (categories.length === 0) {
      return 0;
    }
    const firstDepth = categories[0].depth;
    for (const category of categories.slice(1)) {
      if (category.depth !== firstDepth) {
        throw new ValueError("category depth not uniform");
      }
    }
    return firstDepth;
  }

  /**
   * The offset of *category* in the overall sequence of leaf categories.
   * A non-leaf category gets the index of its first sub-category.
   */
  index(category: any) {
    let index, thisCategory;
    index = 0;
    for (const thisCategory of this._categories) {
      if (category === thisCategory) {
        return index;
      }
      index += thisCategory.leafCount;
    }
    throw new ValueError("category not in top-level categories");
  }

  /**
   * The number of leaf-level categories in this hierarchy. The return
   * value is the same as that of `len()` only when the hierarchy is
   * single level.
   */
  get leafCount(): number {
    return this._categories.reduce((n: number, c: any) => n + c.leafCount, 0);
  }

  /**
   * A generator of (idx, label) sequences representing the category
   * hierarchy from the bottom up. The first level contains all leaf
   * categories, and each subsequent is the next level up.
   */
  get levels(): Generator<[number, any][]> {
    const levels = function* (categories: Iterable<any>): Generator<[number, any][]> {
      const cats = [...categories];
      const subCategories = cats.flatMap((c) => [...c.subCategories]);
      if (subCategories.length) {
        yield* levels(subCategories);
      }
      yield cats.map((cat): [number, any] => [cat.idx, cat.label]);
    };
    return levels(this);
  }

  /**
   * Read/write. Return a string representing the number format used in
   * Excel to format these category values, e.g. '0.0' or 'mm/dd/yyyy'.
   * This string is only relevant when the categories are numeric or date
   * type, although it returns 'General' without error when the categories
   * are string labels. Assigning |None| causes the default number format
   * to be used, based on the type of the category labels.
   */
  get numberFormat() {
    const GENERAL = "General";
    if (this._numberFormat !== null) {
      return this._numberFormat;
    }
    if (this.depth !== 1) {
      return GENERAL;
    }
    const firstCatLabel = this.at(0).label;
    if (firstCatLabel instanceof Date) {
      return "yyyy\\-mm\\-dd";
    }
    return GENERAL;
  }

  set numberFormat(value: any) {
    this._numberFormat = value;
  }
}

/**
 * A chart category, primarily having a label to be displayed on the
 * category axis, but also able to be configured in a hierarchy for support
 * of multi-level category charts.
 */
export class Category {
  declare _label: any;
  declare _parent: any;
  declare _subCategories: any;

  constructor(label: any, parent: any) {
    this._label = label;
    this._parent = parent;
    this._subCategories = [];
  }

  /**
   * Return a newly created |data.Category| object having *label* and
   * appended to the end of the sub-category sequence for this category.
   */
  addSubCategory(label: any) {
    const category = new Category(label, this);
    this._subCategories.push(category);
    return category;
  }

  /**
   * The number of hierarchy levels rooted at this category node. Returns
   * 1 if this category has no sub-categories.
   */
  get depth(): number {
    const subCategories = this._subCategories;
    if (subCategories.length === 0) {
      return 1;
    }
    const firstDepth = subCategories[0].depth;
    for (const category of subCategories.slice(1)) {
      if (category.depth !== firstDepth) {
        throw new ValueError("category depth not uniform");
      }
    }
    return (firstDepth + 1);
  }

  /**
   * The offset of this category in the overall sequence of leaf
   * categories. A non-leaf category gets the index of its first
   * sub-category.
   */
  get idx() {
    return this._parent.index(this);
  }

  /**
   * The offset of *sub_category* in the overall sequence of leaf
   * categories.
   */
  index(subCategory: any) {
    let index, thisSubCategory;
    index = this._parent.index(this);
    for (const thisSubCategory of this._subCategories) {
      if (subCategory === thisSubCategory) {
        return index;
      }
      index += thisSubCategory.leafCount;
    }
    throw new ValueError("sub_category not in this category");
  }

  /**
   * The number of leaf category nodes under this category. Returns
   * 1 if this category has no sub-categories.
   */
  get leafCount() {
    if (this._subCategories.length === 0) {
      return 1;
    }
    return this._subCategories.reduce((n: number, category: any) => n + category.leafCount, 0);
  }

  /**
   * The value that appears on the axis for this category. The label can
   * be a string, a number, or a datetime.date or datetime.datetime
   * object.
   */
  get label() {
    return (this._label !== null && this._label !== undefined ? this._label : "");
  }

  /**
   * The string representation of the numeric (or date) label of this
   * category, suitable for use in the XML `c:pt` element for this
   * category. The optional *date_1904* parameter specifies the epoch used
   * for calculating Excel date numbers.
   */
  numericStrVal(date1904: any = false) {
    const label = this._label;
    if (label instanceof Date) {
      return `${this._excelDateNumber(date1904)}.0`;
    }
    // python str(): JS can't tell 2.0 from 2, so an integral float is written "2" (python "2.0")
    return pyFormatValue(this._label);
  }

  /** The sequence of child categories for this category. */
  get subCategories() {
    return this._subCategories;
  }

  /**
   * Return an integer representing the date label of this category as the
   * number of days since January 1, 1900 (or 1904 if date_1904 is
   * |True|).
   */
  _excelDateNumber(date1904: any): number {
    let excelDayNumber;
    const label: Date = this._label;
    // the date part only (a datetime's time of day is dropped), naive dates held as UTC
    const date = new Date(label.getTime());
    const date_ = date.setUTCHours(0, 0, 0, 0);
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    excelDayNumber = Math.round((date_ - epoch) / 86400000);
    if ((!date1904 && excelDayNumber > 59)) {
      excelDayNumber += 1;
    }
    return excelDayNumber;
  }
}

/**
 * |ChartData| is simply an alias for |CategoryChartData| and may be removed
 * in a future release. All new development should use |CategoryChartData|
 * for creating or replacing the data in chart types other than XY and
 * Bubble.
 */
export class ChartData extends CategoryChartData {
}

/**
 * The data specific to a particular category chart series. It provides
 * access to the series label, the series data points, and an optional
 * number format to be applied to each data point not having a specified
 * number format.
 */
export class CategorySeriesData extends _BaseSeriesData {

  /**
   * Return a CategoryDataPoint object newly created with value *value*,
   * an optional *number_format*, and appended to this sequence.
   */
  addDataPoint(value: any, numberFormat: any = null) {
    const dataPoint = new CategoryDataPoint(this, value, numberFormat);
    this.append(dataPoint);
    return dataPoint;
  }

  /**
   * The |data.Categories| object that provides access to the category
   * objects for this series.
   */
  get categories() {
    return this._chartData.categories;
  }

  /**
   * The Excel worksheet reference to the categories for this chart (not
   * including the column heading).
   */
  get categoriesRef() {
    return this._chartData.categoriesRef;
  }

  /**
   * A sequence containing the (Y) value of each datapoint in this series,
   * in data point order.
   */
  get values() {
    return [...this._dataPoints].map((dp) => dp.value);
  }

  /**
   * The Excel worksheet reference to the (Y) values for this series (not
   * including the column heading).
   */
  get valuesRef() {
    return this._chartData.valuesRef(this);
  }
}

/**
 * A specialized ChartData object suitable for use with an XY (aka. scatter)
 * chart. Unlike ChartData, it has no category sequence. Rather, each data
 * point of each series specifies both an X and a Y value.
 */
export class XyChartData extends _BaseChartData {

  /**
   * Return an |XySeriesData| object newly created and added at the end of
   * this sequence, identified by *name* and values formatted with
   * *number_format*.
   */
  addSeries(name: any, numberFormat: any = null) {
    const seriesData = new XySeriesData(this, name, numberFormat);
    this.append(seriesData);
    return seriesData;
  }

  /**
   * The worksheet writer object to which layout and writing of the Excel
   * worksheet for this chart will be delegated.
   */
  override get _workbookWriter(): any {
    return lazy(this, "_workbookWriter", () => {
      return new XyWorkbookWriter(this);
    });
  }
}

/**
 * A specialized ChartData object suitable for use with a bubble chart.
 * A bubble chart is essentially an XY chart where the markers are scaled to
 * provide a third quantitative dimension to the exhibit.
 */
export class BubbleChartData extends XyChartData {

  /**
   * Return a |BubbleSeriesData| object newly created and added at the end
   * of this sequence, and having series named *name* and values formatted
   * with *number_format*.
   */
  addSeries(name: any, numberFormat: any = null) {
    const seriesData = new BubbleSeriesData(this, name, numberFormat);
    this.append(seriesData);
    return seriesData;
  }

  /**
   * The Excel worksheet reference for the range containing the bubble
   * sizes for *series*.
   */
  bubbleSizesRef(series: any) {
    return this._workbookWriter.bubbleSizesRef(series);
  }

  /**
   * The worksheet writer object to which layout and writing of the Excel
   * worksheet for this chart will be delegated.
   */
  override get _workbookWriter(): any {
    return lazy(this, "_workbookWriter", () => {
      return new BubbleWorkbookWriter(this);
    });
  }
}

/**
 * The data specific to a particular XY chart series. It provides access to
 * the series label, the series data points, and an optional number format
 * to be applied to each data point not having a specified number format.
 *
 * The sequence of data points in an XY series is significant; lines are
 * plotted following the sequence of points, even if that causes a line
 * segment to "travel backward" (implying a multi-valued function). The data
 * points are not automatically sorted into increasing order by X value.
 */
export class XySeriesData extends _BaseSeriesData {

  /**
   * Return an XyDataPoint object newly created with values *x* and *y*,
   * and appended to this sequence.
   */
  addDataPoint(x: any, y: any, numberFormat: any = null) {
    const dataPoint = new XyDataPoint(this, x, y, numberFormat);
    this.append(dataPoint);
    return dataPoint;
  }
}

/**
 * The data specific to a particular Bubble chart series. It provides access
 * to the series label, the series data points, and an optional number
 * format to be applied to each data point not having a specified number
 * format.
 *
 * The sequence of data points in a bubble chart series is maintained
 * throughout the chart building process because a data point has no unique
 * identifier and can only be retrieved by index.
 */
export class BubbleSeriesData extends XySeriesData {

  /**
   * Append a new BubbleDataPoint object having the values *x*, *y*, and
   * *size*. The optional *number_format* is used to format the Y value.
   * If not provided, the number format is inherited from the series data.
   */
  addDataPoint(x: any, y: any, size: any, numberFormat: any = null) {
    const dataPoint = new BubbleDataPoint(this, x, y, size, numberFormat);
    this.append(dataPoint);
    return dataPoint;
  }

  /**
   * A sequence containing the bubble size for each datapoint in this
   * series, in data point order.
   */
  get bubbleSizes() {
    return [...this._dataPoints].map((dp) => dp.bubbleSize);
  }

  /**
   * The Excel worksheet reference for the range containing the bubble
   * sizes for this series.
   */
  get bubbleSizesRef() {
    return this._chartData.bubbleSizesRef(this);
  }
}

/**
 * A data point in a category chart series. Provides access to the value of
 * the datapoint and the number format with which it should appear in the
 * Excel file.
 */
export class CategoryDataPoint extends _BaseDataPoint {
  declare _value: any;

  constructor(seriesData: any, value: any, numberFormat: any) {
    super(seriesData, numberFormat);
    this._value = value;
  }

  /** The (Y) value for this category data point. */
  get value() {
    return this._value;
  }
}

/**
 * A data point in an XY chart series. Provides access to the x and y values
 * of the datapoint.
 */
export class XyDataPoint extends _BaseDataPoint {
  declare _x: any;
  declare _y: any;

  constructor(seriesData: any, x: any, y: any, numberFormat: any) {
    super(seriesData, numberFormat);
    this._x = x;
    this._y = y;
  }

  /** The X value for this XY data point. */
  get x() {
    return this._x;
  }

  /** The Y value for this XY data point. */
  get y() {
    return this._y;
  }
}

/**
 * A data point in a bubble chart series. Provides access to the x, y, and
 * size values of the datapoint.
 */
export class BubbleDataPoint extends XyDataPoint {
  declare _size: any;

  constructor(seriesData: any, x: any, y: any, size: any, numberFormat: any) {
    super(seriesData, x, y, numberFormat);
    this._size = size;
  }

  /** The value representing the size of the bubble for this data point. */
  get bubbleSize() {
    return this._size;
  }
}
