// Chart part objects, including Chart and Charts (`pptx.parts.chart`).
import type { XL_CHART_TYPE } from "../enum/chart.js";
import { NotImplementedError } from "../errors.js";
import { CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT } from "../opc/constants.js";
import { XmlPart, type Part } from "../opc/package.js";
import type { Package } from "../package.js";
import { lazy } from "../util.js";
import { EmbeddedXlsxPart } from "./embeddedpackage.js";

/** What `ChartPart.new()` needs from a `ChartData` object (`pptx.chart.data`, stage 5). */
export interface ChartDataSource {
  xmlBytes(chartType: XL_CHART_TYPE): Uint8Array;
  readonly xlsxBlob: Uint8Array;
}

/**
 * A chart part.
 *
 * Corresponds to parts having partnames matching ppt/charts/chart[1-9][0-9]*.xml
 */
export class ChartPart extends XmlPart {
  static partnameTemplate = "/ppt/charts/chart%d.xml";

  /**
   * A new `ChartPart` instance added to `package_` containing a chart of `chartType` and
   * depicting `chartData`.
   */
  static new(chartType: XL_CHART_TYPE, chartData: ChartDataSource, package_: Package): ChartPart {
    const chartPart = this.load(package_.nextPartname(this.partnameTemplate), CT.DML_CHART, package_, chartData.xmlBytes(chartType)) as ChartPart;
    chartPart.chartWorkbook.updateFromXlsxBlob(chartData.xlsxBlob);
    return chartPart;
  }

  /** `Chart` object representing the chart in this part. */
  get chart(): never {
    throw new NotImplementedError("chart objects (pptx.chart) are ported in stage 5, see gljeremy/pptx-ts#5");
  }

  /**
   * The `ChartWorkbook` object providing access to the external chart data in a linked or
   * embedded Excel workbook.
   */
  get chartWorkbook(): ChartWorkbook {
    return lazy(this, "chartWorkbook", () => new ChartWorkbook(this._element, this));
  }
}

/** Provides access to external chart data in a linked or embedded Excel workbook. */
export class ChartWorkbook {
  declare _chartSpace: any;
  declare _chartPart: ChartPart;

  constructor(chartSpace: any, chartPart: ChartPart) {
    this._chartSpace = chartSpace;
    this._chartPart = chartPart;
  }

  /**
   * Replace the Excel spreadsheet in the related `EmbeddedXlsxPart` with `xlsxBlob`.
   *
   * Adds a new `EmbeddedXlsxPart` if there isn't one.
   */
  updateFromXlsxBlob(xlsxBlob: Uint8Array): void {
    const xlsxPart = this.xlsxPart;
    if (xlsxPart === null) {
      this.xlsxPart = EmbeddedXlsxPart.new(xlsxBlob, this._chartPart.package);
      return;
    }
    xlsxPart.blob = xlsxBlob;
  }

  /**
   * Optional `EmbeddedXlsxPart` object containing data for this chart.
   *
   * This related part has its rId at `c:chartSpace/c:externalData/@rId`. null if there is no
   * `<c:externalData>` element.
   */
  get xlsxPart(): Part | null {
    const xlsxPartRId = this._chartSpace.xlsxPartRId;
    return xlsxPartRId === null || xlsxPartRId === undefined ? null : this._chartPart.relatedPart(xlsxPartRId);
  }
  /** Set the related `EmbeddedXlsxPart` to `xlsxPart`, adding a `c:externalData` element if needed. */
  set xlsxPart(xlsxPart: Part) {
    const rId = this._chartPart.relateTo(xlsxPart, RT.PACKAGE);
    const externalData = this._chartSpace.getOrAddExternalData();
    externalData.rId = rId;
  }
}
