// Chart builder and related objects. (`pptx.chart.xlsx`)
import { ValueError, NotImplementedError } from "../errors.js";
import { BytesIO } from "../fileio.js";
import { pyFormatValue, pyPercent } from "../util.js";
import { Workbook, type Worksheet } from "../xlsx/workbook.js";

/** Base class for workbook writers, providing shared members. */
export class _BaseWorkbookWriter {
  declare _chartData: any;

  constructor(chartData: any) {
    this._chartData = chartData;
  }

  /** bytes for Excel file containing chart_data. */
  get xlsxBlob(): Uint8Array {
    const xlsxFile = new BytesIO();
    // python: `with self._open_worksheet(xlsx_file) as (workbook, worksheet):`
    const ctx = this._openWorksheet(xlsxFile);
    const [workbook, worksheet] = ctx.__enter__();
    this._populateWorksheet(workbook, worksheet);
    ctx.__exit__(null, null, null);
    return xlsxFile.getvalue();
  }

  /**
   * Enable XlsxWriter Worksheet object to be opened, operated on, and
   * then automatically closed within a `with` statement. A filename or
   * stream object (such as an `io.BytesIO` instance) is expected as
   * *xlsx_file*.
   */
  _openWorksheet(xlsxFile: any): { __enter__(): [Workbook, Worksheet]; __exit__(...exc: unknown[]): void } {
    // python's @contextmanager generator: the workbook is closed only when the block completes
    let workbook: Workbook;
    return {
      __enter__: (): [Workbook, Worksheet] => {
        workbook = new Workbook(xlsxFile, { "in_memory": true });
        const worksheet = workbook.addWorksheet();
        return [workbook, worksheet];
      },
      __exit__: (excType: unknown = null) => {
        if (excType === null || excType === undefined) workbook.close();
      },
    };
  }

  /**
   * Must be overridden by each subclass to provide the particulars of
   * writing the spreadsheet data.
   */
  _populateWorksheet(workbook: any, worksheet: any): void {
    throw new NotImplementedError("must be provided by each subclass");
  }
}

/**
 * Determines Excel worksheet layout and can write an Excel workbook from
 * a CategoryChartData object. Serves as the authority for Excel worksheet
 * ranges.
 */
export class CategoryWorkbookWriter extends _BaseWorkbookWriter {

  /**
   * The Excel worksheet reference to the categories for this chart (not
   * including the column heading).
   */
  get categoriesRef() {
    const categories = this._chartData.categories;
    if (categories.depth === 0) {
      throw new ValueError("chart data contains no categories");
    }
    const rightCol = String.fromCharCode(("A".charCodeAt(0) + categories.depth) - 1);
    const bottomRow = (categories.leafCount + 1);
    return `Sheet1!$A$2:$${rightCol}$${bottomRow}`;
  }

  /**
   * Return the Excel worksheet reference to the cell containing the name
   * for *series*. This also serves as the column heading for the series
   * values.
   */
  seriesNameRef(series: any) {
    return `Sheet1!$${this._seriesColLetter(series)}$1`;
  }

  /**
   * The Excel worksheet reference to the values for this series (not
   * including the column heading).
   */
  valuesRef(series: any) {
    return `Sheet1!$${pyFormatValue(this._seriesColLetter(series))}$2:$${pyFormatValue(this._seriesColLetter(series))}$${pyFormatValue((series.length + 1))}`;
  }

  /**
   * Return str Excel column reference like 'BQ' for *column_number*.
   *
   * *column_number* is an int in the range 1-16384 inclusive, where
   * 1 maps to column 'A'.
   */
  static _columnReference(columnNumber: any) {
    let colRef, remainder;
    if ((columnNumber < 1 || columnNumber > 16384)) {
      throw new ValueError("column_number must be in range 1-16384");
    }
    colRef = "";
    while (columnNumber) {
      remainder = pyPercent(columnNumber, [26]);
      if (remainder === 0) {
        remainder = 26;
      }
      const colLetter = String.fromCharCode(("A".charCodeAt(0) + remainder) - 1);
      colRef = (colLetter + colRef);
      columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return colRef;
  }

  /**
   * Write the chart data contents to *worksheet* in category chart
   * layout. Write categories starting in the first column starting in
   * the second row, and proceeding one column per category level (for
   * charts having multi-level categories). Write series as columns
   * starting in the next following column, placing the series title in
   * the first cell.
   */
  _populateWorksheet(workbook: any, worksheet: any) {
    this._writeCategories(workbook, worksheet);
    this._writeSeries(workbook, worksheet);
  }

  /**
   * The letter of the Excel worksheet column in which the data for a
   * series appears.
   */
  _seriesColLetter(series: any) {
    const columnNumber = ((1 + series.categories.depth) + series.index);
    return (this.constructor as typeof CategoryWorkbookWriter)._columnReference(columnNumber);
  }

  /**
   * Write the categories column(s) to *worksheet*. Categories start in
   * the first column starting in the second row, and proceeding one
   * column per category level (for charts having multi-level categories).
   * A date category is formatted as a date. All others are formatted
   * `General`.
   */
  _writeCategories(workbook: any, worksheet: any) {
    const categories = this._chartData.categories;
    const numFormat = workbook.addFormat({ "num_format": categories.numberFormat });
    const depth = categories.depth;
    for (const [idx, level] of [...categories.levels].entries()) {
      const col = ((depth - idx) - 1);
      this._writeCatColumn(worksheet, col, level, numFormat);
    }
  }

  /**
   * Write a category column defined by *level* to *worksheet* at offset
   * *col* and formatted with *num_format*.
   */
  _writeCatColumn(worksheet: any, col: any, level: any, numFormat: any) {
    worksheet.setColumn(col, col, 10);
    for (const [off, name] of level) {
      const row = (off + 1);
      worksheet.write(row, col, name, numFormat);
    }
  }

  /**
   * Write the series column(s) to *worksheet*. Series start in the column
   * following the last categories column, placing the series title in the
   * first cell.
   */
  _writeSeries(workbook: any, worksheet: any) {
    const colOffset = this._chartData.categories.depth;
    for (const [idx, series] of [...this._chartData].entries()) {
      const numFormat = workbook.addFormat({ "num_format": series.numberFormat });
      const seriesCol = (idx + colOffset);
      worksheet.write(0, seriesCol, series.name);
      worksheet.writeColumn(1, seriesCol, series.values, numFormat);
    }
  }
}

/**
 * Determines Excel worksheet layout and can write an Excel workbook from XY
 * chart data. Serves as the authority for Excel worksheet ranges.
 */
export class XyWorkbookWriter extends _BaseWorkbookWriter {

  /**
   * Return the Excel worksheet reference to the cell containing the name
   * for *series*. This also serves as the column heading for the series
   * Y values.
   */
  seriesNameRef(series: any) {
    const row = (this.seriesTableRowOffset(series) + 1);
    return `Sheet1!$B$${row}`;
  }

  /**
   * Return the number of rows preceding the data table for *series* in
   * the Excel worksheet.
   */
  seriesTableRowOffset(series: any) {
    const titleAndSpacerRows = (series.index * 2);
    const dataPointRows = series.dataPointOffset;
    return (titleAndSpacerRows + dataPointRows);
  }

  /**
   * The Excel worksheet reference to the X values for this chart (not
   * including the column label).
   */
  xValuesRef(series: any) {
    const topRow = (this.seriesTableRowOffset(series) + 2);
    const bottomRow = ((topRow + series.length) - 1);
    return `Sheet1!$A$${topRow}:$A$${bottomRow}`;
  }

  /**
   * The Excel worksheet reference to the Y values for this chart (not
   * including the column label).
   */
  yValuesRef(series: any) {
    const topRow = (this.seriesTableRowOffset(series) + 2);
    const bottomRow = ((topRow + series.length) - 1);
    return `Sheet1!$B$${topRow}:$B$${bottomRow}`;
  }

  /**
   * Write chart data contents to *worksheet* in the standard XY chart
   * layout. Write the data for each series to a separate two-column
   * table, X values in column A and Y values in column B. Place the
   * series label in the first (heading) cell of the column.
   */
  _populateWorksheet(workbook: any, worksheet: any) {
    const chartNumFormat = workbook.addFormat({ "num_format": this._chartData.numberFormat });
    for (const series of this._chartData) {
      const seriesNumFormat = workbook.addFormat({ "num_format": series.numberFormat });
      const offset = this.seriesTableRowOffset(series);
      worksheet.writeColumn((offset + 1), 0, series.xValues, chartNumFormat);
      worksheet.write(offset, 1, series.name);
      worksheet.writeColumn((offset + 1), 1, series.yValues, seriesNumFormat);
    }
  }
}

/**
 * Service object that knows how to write an Excel workbook from bubble
 * chart data.
 */
export class BubbleWorkbookWriter extends XyWorkbookWriter {

  /**
   * The Excel worksheet reference to the range containing the bubble
   * sizes for *series* (not including the column heading cell).
   */
  bubbleSizesRef(series: any) {
    const topRow = (this.seriesTableRowOffset(series) + 2);
    const bottomRow = ((topRow + series.length) - 1);
    return `Sheet1!$C$${topRow}:$C$${bottomRow}`;
  }

  /**
   * Write chart data contents to *worksheet* in the bubble chart layout.
   * Write the data for each series to a separate three-column table with
   * X values in column A, Y values in column B, and bubble sizes in
   * column C. Place the series label in the first (heading) cell of the
   * values column.
   */
  _populateWorksheet(workbook: any, worksheet: any) {
    const chartNumFormat = workbook.addFormat({ "num_format": this._chartData.numberFormat });
    for (const series of this._chartData) {
      const seriesNumFormat = workbook.addFormat({ "num_format": series.numberFormat });
      const offset = this.seriesTableRowOffset(series);
      worksheet.writeColumn((offset + 1), 0, series.xValues, chartNumFormat);
      worksheet.write(offset, 1, series.name);
      worksheet.writeColumn((offset + 1), 1, series.yValues, seriesNumFormat);
      worksheet.write(offset, 2, "Size");
      worksheet.writeColumn((offset + 1), 2, series.bubbleSizes, chartNumFormat);
    }
  }
}
