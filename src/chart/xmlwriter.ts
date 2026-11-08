// Composers for default chart XML for various chart types. (`pptx.chart.xmlwriter`)
import { KeyError, NotImplementedError } from "../errors.js";
import { EnumDict } from "../enum/base.js";
import { pyFormat, pyFormatValue, xmlEscape } from "../util.js";
import { XL_CHART_TYPE } from "../enum/chart.js";
import { parseXml } from "../oxml/parser.js";
import { nsdecls } from "../oxml/ns.js";

/**
 * Factory function returning appropriate XML writer object for
 * *chart_type*, loaded with *chart_type* and *chart_data*.
 */
export function ChartXmlWriter(chartType: any, chartData: any) {
  const XL_CT = XL_CHART_TYPE;
  let BuilderCls;
  try {
    BuilderCls = new EnumDict<any>([[XL_CT.AREA, _AreaChartXmlWriter], [XL_CT.AREA_STACKED, _AreaChartXmlWriter], [XL_CT.AREA_STACKED_100, _AreaChartXmlWriter], [XL_CT.BAR_CLUSTERED, _BarChartXmlWriter], [XL_CT.BAR_STACKED, _BarChartXmlWriter], [XL_CT.BAR_STACKED_100, _BarChartXmlWriter], [XL_CT.BUBBLE, _BubbleChartXmlWriter], [XL_CT.BUBBLE_THREE_D_EFFECT, _BubbleChartXmlWriter], [XL_CT.COLUMN_CLUSTERED, _BarChartXmlWriter], [XL_CT.COLUMN_STACKED, _BarChartXmlWriter], [XL_CT.COLUMN_STACKED_100, _BarChartXmlWriter], [XL_CT.DOUGHNUT, _DoughnutChartXmlWriter], [XL_CT.DOUGHNUT_EXPLODED, _DoughnutChartXmlWriter], [XL_CT.LINE, _LineChartXmlWriter], [XL_CT.LINE_MARKERS, _LineChartXmlWriter], [XL_CT.LINE_MARKERS_STACKED, _LineChartXmlWriter], [XL_CT.LINE_MARKERS_STACKED_100, _LineChartXmlWriter], [XL_CT.LINE_STACKED, _LineChartXmlWriter], [XL_CT.LINE_STACKED_100, _LineChartXmlWriter], [XL_CT.PIE, _PieChartXmlWriter], [XL_CT.PIE_EXPLODED, _PieChartXmlWriter], [XL_CT.RADAR, _RadarChartXmlWriter], [XL_CT.RADAR_FILLED, _RadarChartXmlWriter], [XL_CT.RADAR_MARKERS, _RadarChartXmlWriter], [XL_CT.XY_SCATTER, _XyChartXmlWriter], [XL_CT.XY_SCATTER_LINES, _XyChartXmlWriter], [XL_CT.XY_SCATTER_LINES_NO_MARKERS, _XyChartXmlWriter], [XL_CT.XY_SCATTER_SMOOTH, _XyChartXmlWriter], [XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS, _XyChartXmlWriter]]).getItem(chartType);
  } catch (e) {
    if (!(e instanceof KeyError)) throw e;
    throw new NotImplementedError(`XML writer for chart type ${chartType} not yet implemented`);
  }
  return new BuilderCls(chartType, chartData);
}

/** Return a |_BaseSeriesXmlRewriter| subclass appropriate to *chart_type*. */
export function SeriesXmlRewriterFactory(chartType: any, chartData: any) {
  const XL_CT = XL_CHART_TYPE;
  const RewriterCls = new EnumDict<any>([[XL_CT.BUBBLE, _BubbleSeriesXmlRewriter], [XL_CT.BUBBLE_THREE_D_EFFECT, _BubbleSeriesXmlRewriter], [XL_CT.XY_SCATTER, _XySeriesXmlRewriter], [XL_CT.XY_SCATTER_LINES, _XySeriesXmlRewriter], [XL_CT.XY_SCATTER_LINES_NO_MARKERS, _XySeriesXmlRewriter], [XL_CT.XY_SCATTER_SMOOTH, _XySeriesXmlRewriter], [XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS, _XySeriesXmlRewriter]]).get(chartType, _CategorySeriesXmlRewriter);
  return new RewriterCls(chartData);
}

/**
 * Generates XML text (unicode) for a default chart, like the one added by
 * PowerPoint when you click the *Add Column Chart* button on the ribbon.
 * Differentiated XML for different chart types is provided by subclasses.
 */
export class _BaseChartXmlWriter {
  declare _chartType: any;
  declare _chartData: any;
  declare _seriesSeq: any;

  constructor(chartType: any, seriesSeq: any) {
    this._chartType = chartType;
    this._chartData = seriesSeq;
    this._seriesSeq = [...seriesSeq];
  }

  /**
   * The full XML stream for the chart specified by this chart builder, as
   * unicode text. This method must be overridden by each subclass.
   */
  get xml(): string {
    throw new NotImplementedError("must be implemented by all subclasses");
  }
}

/** Provides shared members for series XML writers. */
export class _BaseSeriesXmlWriter {
  declare _series: any;
  declare _date1904: any;

  constructor(series: any, date1904: any = false) {
    this._series = series;
    this._date1904 = date1904;
  }

  /** The XML-escaped name for this series. */
  get name() {
    return xmlEscape(this._series.name);
  }

  /**
   * Return the ``<c:numRef>`` element specified by the parameters as
   * unicode text.
   */
  numRefXml(wkshtRef: any, numberFormat: any, values: any) {
    const ptXml = this.ptXml(values);
    return `            <c:numRef>
              <c:f>${pyFormatValue(wkshtRef)}</c:f>
              <c:numCache>
                <c:formatCode>${pyFormatValue(numberFormat)}</c:formatCode>
${ptXml}              </c:numCache>
            </c:numRef>
`;
  }

  /**
   * Return the ``<c:ptCount>`` and sequence of ``<c:pt>`` elements
   * corresponding to *values* as a single unicode text string.
   * `c:ptCount` refers to the number of `c:pt` elements in this sequence.
   * The `idx` attribute value for `c:pt` elements locates the data point
   * in the overall data point sequence of the chart and is started at
   * *offset*.
   */
  ptXml(values: any) {
    let xml, idx, value;
    xml = `                <c:ptCount val="${pyFormatValue(values.length)}"/>
`;
    const ptTmpl = "                <c:pt idx=\"{idx}\">\n                  <c:v>{value}</c:v>\n                </c:pt>\n";
    for (const [idx, value] of [...values].entries()) {
      if (value === null) {
        continue;
      }
      xml += pyFormat(ptTmpl, { "idx": idx, "value": value });
    }
    return xml;
  }

  /**
   * Return a ``<c:tx>`` oxml element for this series, containing the
   * series name.
   */
  get tx() {
    const xml = pyFormat(this._txTmpl, { "wksht_ref": this._series.nameRef, "series_name": this.name, "nsdecls": ` ${nsdecls("c")}` });
    return parseXml(xml);
  }

  /**
   * Return the ``<c:tx>`` (tx is short for 'text') element for this
   * series as unicode text. This element contains the series name.
   */
  get txXml() {
    return pyFormat(this._txTmpl, { "wksht_ref": this._series.nameRef, "series_name": this.name, "nsdecls": "" });
  }

  /**
   * The string formatting template for the ``<c:tx>`` element for this
   * series, containing the series title and spreadsheet range reference.
   */
  get _txTmpl() {
    return "          <c:tx{nsdecls}>\n            <c:strRef>\n              <c:f>{wksht_ref}</c:f>\n              <c:strCache>\n                <c:ptCount val=\"1\"/>\n                <c:pt idx=\"0\">\n                  <c:v>{series_name}</c:v>\n                </c:pt>\n              </c:strCache>\n            </c:strRef>\n          </c:tx>\n";
  }
}

/** Base class for series XML rewriters. */
export class _BaseSeriesXmlRewriter {
  declare _chartData: any;

  constructor(chartData: any) {
    this._chartData = chartData;
  }

  /**
   * Rewrite the series data under *chartSpace* using the chart data
   * contents. All series-level formatting is left undisturbed. If
   * the chart data contains fewer series than *chartSpace*, the extra
   * series in *chartSpace* are deleted. If *chart_data* contains more
   * series than the *chartSpace* element, new series are added to the
   * last plot in the chart and series formatting is "cloned" from the
   * last series in that plot.
   */
  replaceSeriesData(chartSpace: any) {
    const [plotArea, date1904] = [chartSpace.plotArea, chartSpace.date1904Val];
    const chartData = this._chartData;
    this._adjustSerCount(plotArea, chartData.length);
    const sers = plotArea.sers;
    const seriesSeq = [...chartData];
    // python zip(): stops at the shorter sequence
    for (let i = 0; i < Math.min(sers.length, seriesSeq.length); i++) {
      this._rewriteSerData(sers[i], seriesSeq[i], date1904);
    }
  }

  /**
   * Add `c:ser` elements to the last xChart element in *plotArea*, cloned
   * from the last `c:ser` child of that last xChart.
   */
  _addClonedSers(plotArea: any, count: any) {
    let lastSer;
    const cloneSer = (ser: any) => {
      const newSer = ser.deepcopy();
      newSer.idx.val = plotArea.nextIdx;
      newSer.order.val = plotArea.nextOrder;
      ser.addnext(newSer);
      return newSer;
    };
    lastSer = plotArea.lastSer;
    for (let _ = 0; _ < count; _++) {
      lastSer = cloneSer(lastSer);
    }
  }

  /**
   * Adjust the number of c:ser elements in *plotArea* to *new_ser_count*.
   * Excess c:ser elements are deleted from the end, along with any xChart
   * elements that are left empty as a result. Series elements are
   * considered in xChart + series order. Any new c:ser elements required
   * are added to the last xChart element and cloned from the last c:ser
   * element in that xChart.
   */
  _adjustSerCount(plotArea: any, newSerCount: any) {
    const serCountDiff = (newSerCount - plotArea.sers.length);
    if (serCountDiff > 0) {
      this._addClonedSers(plotArea, serCountDiff);
    } else if (serCountDiff < 0) {
      this._trimSerCountBy(plotArea, Math.abs(serCountDiff));
    }
  }

  /**
   * Rewrite selected child elements of *ser* based on the values in
   * *series_data*.
   */
  _rewriteSerData(ser: any, seriesData: any, date1904: any) {
    throw new NotImplementedError("must be implemented by each subclass");
  }

  /**
   * Remove the last *count* ser elements from *plotArea*. Any xChart
   * elements having no ser child elements after trimming are also
   * removed.
   */
  _trimSerCountBy(plotArea: any, count: any) {
    let ser, xChart, parent;
    const extraSers = plotArea.sers.slice(-count);
    for (const ser of extraSers) {
      parent = ser.getparent();
      parent.remove(ser);
    }
    const extraXCharts = [...plotArea.iterXCharts()].filter((xChart) => xChart.sers.length === 0);
    for (const xChart of extraXCharts) {
      parent = xChart.getparent();
      parent.remove(xChart);
    }
  }
}

/** Provides specialized methods particular to the ``<c:areaChart>`` element. */
export class _AreaChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:areaChart>
${this._groupingXml}        <c:varyColors val="0"/>
${this._serXml}        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:axId val="-2101159928"/>
        <c:axId val="-2100718248"/>
      </c:areaChart>
${this._catAxXml}      <c:valAx>
        <c:axId val="-2100718248"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2101159928"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="zero"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _catAxXml() {
    const categories = this._chartData.categories;
    if (categories.areDates) {
      return `      <c:dateAx>
        <c:axId val="-2101159928"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="${pyFormatValue(categories.numberFormat)}" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2100718248"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
`;
    }
    return "      <c:catAx>\n        <c:axId val=\"-2101159928\"/>\n        <c:scaling>\n          <c:orientation val=\"minMax\"/>\n        </c:scaling>\n        <c:delete val=\"0\"/>\n        <c:axPos val=\"b\"/>\n        <c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n        <c:majorTickMark val=\"out\"/>\n        <c:minorTickMark val=\"none\"/>\n        <c:tickLblPos val=\"nextTo\"/>\n        <c:crossAx val=\"-2100718248\"/>\n        <c:crosses val=\"autoZero\"/>\n        <c:auto val=\"1\"/>\n        <c:lblAlgn val=\"ctr\"/>\n        <c:lblOffset val=\"100\"/>\n        <c:noMultiLvlLbl val=\"0\"/>\n      </c:catAx>\n";
  }

  get _groupingXml() {
    const val = new EnumDict([[XL_CHART_TYPE.AREA, "standard"], [XL_CHART_TYPE.AREA_STACKED, "stacked"], [XL_CHART_TYPE.AREA_STACKED_100, "percentStacked"]]).getItem(this._chartType);
    return `        <c:grouping val="${val}"/>
`;
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _CategorySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${xmlWriter.catXml}${xmlWriter.valXml}        </c:ser>
`;
    }
    return xml;
  }
}

/** Provides specialized methods particular to the ``<c:barChart>`` element. */
export class _BarChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:barChart>
${this._barDirXml}${this._groupingXml}${this._serXml}${this._overlapXml}        <c:axId val="-2068027336"/>
        <c:axId val="-2113994440"/>
      </c:barChart>
${this._catAxXml}      <c:valAx>
        <c:axId val="-2113994440"/>
        <c:scaling/>
        <c:delete val="0"/>
        <c:axPos val="${pyFormatValue(this._valAxPos)}"/>
        <c:majorGridlines/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2068027336"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _barDirXml() {
    const XL = XL_CHART_TYPE;
    const barTypes = [XL.BAR_CLUSTERED, XL.BAR_STACKED, XL.BAR_STACKED_100];
    const colTypes = [XL.COLUMN_CLUSTERED, XL.COLUMN_STACKED, XL.COLUMN_STACKED_100];
    if (barTypes.includes(this._chartType)) {
      return "        <c:barDir val=\"bar\"/>\n";
    } else if (colTypes.includes(this._chartType)) {
      return "        <c:barDir val=\"col\"/>\n";
    }
    throw new NotImplementedError(`no _barDir_xml() for chart type ${this._chartType}`);
  }

  get _catAxPos() {
    return new EnumDict([[XL_CHART_TYPE.BAR_CLUSTERED, "l"], [XL_CHART_TYPE.BAR_STACKED, "l"], [XL_CHART_TYPE.BAR_STACKED_100, "l"], [XL_CHART_TYPE.COLUMN_CLUSTERED, "b"], [XL_CHART_TYPE.COLUMN_STACKED, "b"], [XL_CHART_TYPE.COLUMN_STACKED_100, "b"]]).getItem(this._chartType);
  }

  get _catAxXml() {
    const categories = this._chartData.categories;
    if (categories.areDates) {
      return `      <c:dateAx>
        <c:axId val="-2068027336"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${pyFormatValue(this._catAxPos)}"/>
        <c:numFmt formatCode="${pyFormatValue(categories.numberFormat)}" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2113994440"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
`;
    }
    return `      <c:catAx>
        <c:axId val="-2068027336"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${pyFormatValue(this._catAxPos)}"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2113994440"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
`;
  }

  get _groupingXml() {
    const XL = XL_CHART_TYPE;
    const clusteredTypes = [XL.BAR_CLUSTERED, XL.COLUMN_CLUSTERED];
    const stackedTypes = [XL.BAR_STACKED, XL.COLUMN_STACKED];
    const percentStackedTypes = [XL.BAR_STACKED_100, XL.COLUMN_STACKED_100];
    if (clusteredTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"clustered\"/>\n";
    } else if (stackedTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"stacked\"/>\n";
    } else if (percentStackedTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"percentStacked\"/>\n";
    }
    throw new NotImplementedError(`no _grouping_xml() for chart type ${this._chartType}`);
  }

  get _overlapXml() {
    const XL = XL_CHART_TYPE;
    const percentStackedTypes = [XL.BAR_STACKED, XL.BAR_STACKED_100, XL.COLUMN_STACKED, XL.COLUMN_STACKED_100];
    if (percentStackedTypes.includes(this._chartType)) {
      return "        <c:overlap val=\"100\"/>\n";
    }
    return "";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _CategorySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${xmlWriter.catXml}${xmlWriter.valXml}        </c:ser>
`;
    }
    return xml;
  }

  get _valAxPos() {
    return new EnumDict([[XL_CHART_TYPE.BAR_CLUSTERED, "b"], [XL_CHART_TYPE.BAR_STACKED, "b"], [XL_CHART_TYPE.BAR_STACKED_100, "b"], [XL_CHART_TYPE.COLUMN_CLUSTERED, "l"], [XL_CHART_TYPE.COLUMN_STACKED, "l"], [XL_CHART_TYPE.COLUMN_STACKED_100, "l"]]).getItem(this._chartType);
  }
}

/**
 * Provides specialized methods particular to the ``<c:doughnutChart>``
 * element.
 */
export class _DoughnutChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:doughnutChart>
        <c:varyColors val="1"/>
${this._serXml}        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
          <c:showLeaderLines val="1"/>
        </c:dLbls>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _explosionXml() {
    if (this._chartType === XL_CHART_TYPE.DOUGHNUT_EXPLODED) {
      return "          <c:explosion val=\"25\"/>\n";
    }
    return "";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _CategorySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${this._explosionXml}${xmlWriter.catXml}${xmlWriter.valXml}        </c:ser>
`;
    }
    return xml;
  }
}

/** Provides specialized methods particular to the ``<c:lineChart>`` element. */
export class _LineChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:lineChart>
${this._groupingXml}        <c:varyColors val="0"/>
${this._serXml}        <c:marker val="1"/>
        <c:smooth val="0"/>
        <c:axId val="2118791784"/>
        <c:axId val="2140495176"/>
      </c:lineChart>
${this._catAxXml}      <c:valAx>
        <c:axId val="2140495176"/>
        <c:scaling/>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2118791784"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _catAxXml() {
    const categories = this._chartData.categories;
    if (categories.areDates) {
      return `      <c:dateAx>
        <c:axId val="2118791784"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="${pyFormatValue(categories.numberFormat)}" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2140495176"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
`;
    }
    return "      <c:catAx>\n        <c:axId val=\"2118791784\"/>\n        <c:scaling>\n          <c:orientation val=\"minMax\"/>\n        </c:scaling>\n        <c:delete val=\"0\"/>\n        <c:axPos val=\"b\"/>\n        <c:majorTickMark val=\"out\"/>\n        <c:minorTickMark val=\"none\"/>\n        <c:tickLblPos val=\"nextTo\"/>\n        <c:crossAx val=\"2140495176\"/>\n        <c:crosses val=\"autoZero\"/>\n        <c:auto val=\"1\"/>\n        <c:lblAlgn val=\"ctr\"/>\n        <c:lblOffset val=\"100\"/>\n        <c:noMultiLvlLbl val=\"0\"/>\n      </c:catAx>\n";
  }

  get _groupingXml() {
    const XL = XL_CHART_TYPE;
    const standardTypes = [XL.LINE, XL.LINE_MARKERS];
    const stackedTypes = [XL.LINE_STACKED, XL.LINE_MARKERS_STACKED];
    const percentStackedTypes = [XL.LINE_STACKED_100, XL.LINE_MARKERS_STACKED_100];
    if (standardTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"standard\"/>\n";
    } else if (stackedTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"stacked\"/>\n";
    } else if (percentStackedTypes.includes(this._chartType)) {
      return "        <c:grouping val=\"percentStacked\"/>\n";
    }
    throw new NotImplementedError(`no _grouping_xml() for chart type ${this._chartType}`);
  }

  get _markerXml() {
    const XL = XL_CHART_TYPE;
    const noMarkerTypes = [XL.LINE, XL.LINE_STACKED, XL.LINE_STACKED_100];
    if (noMarkerTypes.includes(this._chartType)) {
      return "          <c:marker>\n            <c:symbol val=\"none\"/>\n          </c:marker>\n";
    }
    return "";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _CategorySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${this._markerXml}${xmlWriter.catXml}${xmlWriter.valXml}          <c:smooth val="0"/>
        </c:ser>
`;
    }
    return xml;
  }
}

/** Provides specialized methods particular to the ``<c:pieChart>`` element. */
export class _PieChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
${this._serXml}      </c:pieChart>
    </c:plotArea>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _explosionXml() {
    if (this._chartType === XL_CHART_TYPE.PIE_EXPLODED) {
      return "          <c:explosion val=\"25\"/>\n";
    }
    return "";
  }

  get _serXml() {
    const xmlWriter = new _CategorySeriesXmlWriter(this._chartData.at(0));
    const xml = `        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
${xmlWriter.txXml}${this._explosionXml}${xmlWriter.catXml}${xmlWriter.valXml}        </c:ser>
`;
    return xml;
  }
}

/** Generates XML for the ``<c:radarChart>`` element. */
export class _RadarChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    return `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:roundedCorners val="0"/>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
      <c14:style val="118"/>
    </mc:Choice>
    <mc:Fallback>
      <c:style val="18"/>
    </mc:Fallback>
  </mc:AlternateContent>
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:radarChart>
        <c:radarStyle val="${pyFormatValue(this._radarStyle)}"/>
        <c:varyColors val="0"/>
${this._serXml}        <c:axId val="2073612648"/>
        <c:axId val="-2112772216"/>
      </c:radarChart>
      <c:catAx>
        <c:axId val="2073612648"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2112772216"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="-2112772216"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="cross"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2073612648"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
  }

  get _markerXml() {
    if (this._chartType === XL_CHART_TYPE.RADAR) {
      return "          <c:marker>\n            <c:symbol val=\"none\"/>\n          </c:marker>\n";
    }
    return "";
  }

  get _radarStyle() {
    if (this._chartType === XL_CHART_TYPE.RADAR_FILLED) {
      return "filled";
    }
    return "marker";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _CategorySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${this._markerXml}${xmlWriter.catXml}${xmlWriter.valXml}          <c:smooth val="0"/>
        </c:ser>
`;
    }
    return xml;
  }
}

/** Generates XML for the ``<c:scatterChart>`` element. */
export class _XyChartXmlWriter extends _BaseChartXmlWriter {

  get xml() {
    const xml = `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="${this._scatterStyleVal}"/>
        <c:varyColors val="0"/>
${this._serXml}        <c:axId val="-2128940872"/>
        <c:axId val="-2129643912"/>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="-2128940872"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2129643912"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="-2129643912"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2128940872"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
    return xml;
  }

  get _markerXml() {
    const noMarkerTypes = [XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS];
    if (noMarkerTypes.includes(this._chartType)) {
      return "          <c:marker>\n            <c:symbol val=\"none\"/>\n          </c:marker>\n";
    }
    return "";
  }

  get _scatterStyleVal() {
    const smoothTypes = [XL_CHART_TYPE.XY_SCATTER_SMOOTH, XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS];
    if (smoothTypes.includes(this._chartType)) {
      return "smoothMarker";
    }
    return "lineMarker";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _XySeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}${this._spPrXml}${this._markerXml}${xmlWriter.xValXml}${xmlWriter.yValXml}          <c:smooth val="0"/>
        </c:ser>
`;
    }
    return xml;
  }

  get _spPrXml() {
    if (this._chartType === XL_CHART_TYPE.XY_SCATTER) {
      return "          <c:spPr>\n            <a:ln w=\"47625\">\n              <a:noFill/>\n            </a:ln>\n          </c:spPr>\n";
    }
    return "";
  }
}

/**
 * Provides specialized methods particular to the ``<c:bubbleChart>``
 * element.
 */
export class _BubbleChartXmlWriter extends _XyChartXmlWriter {

  get xml() {
    const xml = `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:bubbleChart>
        <c:varyColors val="0"/>
${this._serXml}        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:bubbleScale val="100"/>
        <c:showNegBubbles val="0"/>
        <c:axId val="-2115720072"/>
        <c:axId val="-2115723560"/>
      </c:bubbleChart>
      <c:valAx>
        <c:axId val="-2115720072"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2115723560"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="-2115723560"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="-2115720072"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="1800"/>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
</c:chartSpace>
`;
    return xml;
  }

  get _bubble3DVal() {
    if (this._chartType === XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT) {
      return "1";
    }
    return "0";
  }

  get _serXml() {
    let xml, series;
    xml = "";
    for (const series of this._chartData) {
      const xmlWriter = new _BubbleSeriesXmlWriter(series);
      xml += `        <c:ser>
          <c:idx val="${pyFormatValue(series.index)}"/>
          <c:order val="${pyFormatValue(series.index)}"/>
${xmlWriter.txXml}          <c:invertIfNegative val="0"/>
${xmlWriter.xValXml}${xmlWriter.yValXml}${xmlWriter.bubbleSizeXml}          <c:bubble3D val="${pyFormatValue(this._bubble3DVal)}"/>
        </c:ser>
`;
    }
    return xml;
  }
}

/** Generates XML snippets particular to a category chart series. */
export class _CategorySeriesXmlWriter extends _BaseSeriesXmlWriter {

  /**
   * Return the ``<c:cat>`` element XML for this series, as an oxml
   * element.
   */
  get cat() {
    const categories = this._series.categories;
    if (categories.areNumeric) {
      return parseXml(pyFormat(this._numRefCatTmpl, { "wksht_ref": this._series.categoriesRef, "number_format": categories.numberFormat, "cat_count": categories.leafCount, "cat_pt_xml": this._catNumPtXml, "nsdecls": ` ${nsdecls("c")}` }));
    }
    if (categories.depth === 1) {
      return parseXml(pyFormat(this._catTmpl, { "wksht_ref": this._series.categoriesRef, "cat_count": categories.leafCount, "cat_pt_xml": this._catPtXml, "nsdecls": ` ${nsdecls("c")}` }));
    }
    return parseXml(pyFormat(this._multiLvlCatTmpl, { "wksht_ref": this._series.categoriesRef, "cat_count": categories.leafCount, "lvl_xml": this._lvlXml(categories), "nsdecls": ` ${nsdecls("c")}` }));
  }

  /**
   * The unicode XML snippet for the ``<c:cat>`` element for this series,
   * containing the category labels and spreadsheet reference.
   */
  get catXml() {
    const categories = this._series.categories;
    if (categories.areNumeric) {
      return pyFormat(this._numRefCatTmpl, { "wksht_ref": this._series.categoriesRef, "number_format": categories.numberFormat, "cat_count": categories.leafCount, "cat_pt_xml": this._catNumPtXml, "nsdecls": "" });
    }
    if (categories.depth === 1) {
      return pyFormat(this._catTmpl, { "wksht_ref": this._series.categoriesRef, "cat_count": categories.leafCount, "cat_pt_xml": this._catPtXml, "nsdecls": "" });
    }
    return pyFormat(this._multiLvlCatTmpl, { "wksht_ref": this._series.categoriesRef, "cat_count": categories.leafCount, "lvl_xml": this._lvlXml(categories), "nsdecls": "" });
  }

  /** The ``<c:val>`` XML for this series, as an oxml element. */
  get val() {
    const xml = pyFormat(this._valTmpl, { "nsdecls": ` ${nsdecls("c")}`, "values_ref": this._series.valuesRef, "number_format": this._series.numberFormat, "val_count": this._series.length, "val_pt_xml": this._valPtXml });
    return parseXml(xml);
  }

  /**
   * Return the unicode XML snippet for the ``<c:val>`` element describing
   * this series, containing the series values and their spreadsheet range
   * reference.
   */
  get valXml() {
    return pyFormat(this._valTmpl, { "nsdecls": "", "values_ref": this._series.valuesRef, "number_format": this._series.numberFormat, "val_count": this._series.length, "val_pt_xml": this._valPtXml });
  }

  /**
   * The unicode XML snippet for the ``<c:pt>`` elements when category
   * labels are numeric (including date type).
   */
  get _catNumPtXml() {
    let xml, idx, category;
    xml = "";
    for (const [idx, category] of [...this._series.categories].entries()) {
      xml += `                <c:pt idx="${pyFormatValue(idx)}">
                  <c:v>${pyFormatValue(category.numericStrVal(this._date1904))}</c:v>
                </c:pt>
`;
    }
    return xml;
  }

  /**
   * The unicode XML snippet for the ``<c:pt>`` elements containing the
   * category names for this series.
   */
  get _catPtXml() {
    let xml, idx, category;
    xml = "";
    for (const [idx, category] of [...this._series.categories].entries()) {
      xml += `                <c:pt idx="${pyFormatValue(idx)}">
                  <c:v>${pyFormatValue(xmlEscape(String(category.label)))}</c:v>
                </c:pt>
`;
    }
    return xml;
  }

  /**
   * The template for the ``<c:cat>`` element for this series, containing
   * the category labels and spreadsheet reference.
   */
  get _catTmpl() {
    return "          <c:cat{nsdecls}>\n            <c:strRef>\n              <c:f>{wksht_ref}</c:f>\n              <c:strCache>\n                <c:ptCount val=\"{cat_count}\"/>\n{cat_pt_xml}              </c:strCache>\n            </c:strRef>\n          </c:cat>\n";
  }

  /**
   * The unicode XML snippet for the ``<c:lvl>`` elements containing
   * multi-level category names.
   */
  _lvlXml(categories: any) {
    let xml, level, idx, name;
    const lvlPtXml = (level: any) => {
      let xml = "";
      for (const [idx, name] of level) {
        xml += `                  <c:pt idx="${idx}">
                    <c:v>${xmlEscape(`${name}`)}</c:v>
                  </c:pt>
`;
      }
      return xml;
    };
    xml = "";
    for (const level of categories.levels) {
      xml += `                <c:lvl>
${lvlPtXml(level)}                </c:lvl>
`;
    }
    return xml;
  }

  /**
   * The template for the ``<c:cat>`` element for this series when there
   * are multi-level (nested) categories.
   */
  get _multiLvlCatTmpl() {
    return "          <c:cat{nsdecls}>\n            <c:multiLvlStrRef>\n              <c:f>{wksht_ref}</c:f>\n              <c:multiLvlStrCache>\n                <c:ptCount val=\"{cat_count}\"/>\n{lvl_xml}              </c:multiLvlStrCache>\n            </c:multiLvlStrRef>\n          </c:cat>\n";
  }

  /**
   * The template for the ``<c:cat>`` element for this series when the
   * labels are numeric (or date) values.
   */
  get _numRefCatTmpl() {
    return "          <c:cat{nsdecls}>\n            <c:numRef>\n              <c:f>{wksht_ref}</c:f>\n              <c:numCache>\n                <c:formatCode>{number_format}</c:formatCode>\n                <c:ptCount val=\"{cat_count}\"/>\n{cat_pt_xml}              </c:numCache>\n            </c:numRef>\n          </c:cat>\n";
  }

  /**
   * The unicode XML snippet containing the ``<c:pt>`` elements containing
   * the values for this series.
   */
  get _valPtXml() {
    let xml, idx, value;
    xml = "";
    for (const [idx, value] of [...this._series.values].entries()) {
      if (value === null) {
        continue;
      }
      xml += `                <c:pt idx="${idx}">
                  <c:v>${pyFormatValue(value)}</c:v>
                </c:pt>
`;
    }
    return xml;
  }

  /**
   * The template for the ``<c:val>`` element for this series, containing
   * the series values and their spreadsheet range reference.
   */
  get _valTmpl() {
    return "          <c:val{nsdecls}>\n            <c:numRef>\n              <c:f>{values_ref}</c:f>\n              <c:numCache>\n                <c:formatCode>{number_format}</c:formatCode>\n                <c:ptCount val=\"{val_count}\"/>\n{val_pt_xml}              </c:numCache>\n            </c:numRef>\n          </c:val>\n";
  }
}

/** Generates XML snippets particular to an XY series. */
export class _XySeriesXmlWriter extends _BaseSeriesXmlWriter {

  /**
   * Return the ``<c:xVal>`` element for this series as an oxml element.
   * This element contains the X values for this series.
   */
  get xVal() {
    const xml = pyFormat(this._xValTmpl, { "nsdecls": ` ${nsdecls("c")}`, "numRef_xml": this.numRefXml(this._series.xValuesRef, this._series.numberFormat, this._series.xValues) });
    return parseXml(xml);
  }

  /**
   * Return the ``<c:xVal>`` element for this series as unicode text. This
   * element contains the X values for this series.
   */
  get xValXml() {
    return pyFormat(this._xValTmpl, { "nsdecls": "", "numRef_xml": this.numRefXml(this._series.xValuesRef, this._series.numberFormat, this._series.xValues) });
  }

  /**
   * Return the ``<c:yVal>`` element for this series as an oxml element.
   * This element contains the Y values for this series.
   */
  get yVal() {
    const xml = pyFormat(this._yValTmpl, { "nsdecls": ` ${nsdecls("c")}`, "numRef_xml": this.numRefXml(this._series.yValuesRef, this._series.numberFormat, this._series.yValues) });
    return parseXml(xml);
  }

  /**
   * Return the ``<c:yVal>`` element for this series as unicode text. This
   * element contains the Y values for this series.
   */
  get yValXml() {
    return pyFormat(this._yValTmpl, { "nsdecls": "", "numRef_xml": this.numRefXml(this._series.yValuesRef, this._series.numberFormat, this._series.yValues) });
  }

  /**
   * The template for the ``<c:xVal>`` element for this series, containing
   * the X values and their spreadsheet range reference.
   */
  get _xValTmpl() {
    return "          <c:xVal{nsdecls}>\n{numRef_xml}          </c:xVal>\n";
  }

  /**
   * The template for the ``<c:yVal>`` element for this series, containing
   * the Y values and their spreadsheet range reference.
   */
  get _yValTmpl() {
    return "          <c:yVal{nsdecls}>\n{numRef_xml}          </c:yVal>\n";
  }
}

/** Generates XML snippets particular to a bubble chart series. */
export class _BubbleSeriesXmlWriter extends _XySeriesXmlWriter {

  /**
   * Return the ``<c:bubbleSize>`` element for this series as an oxml
   * element. This element contains the bubble size values for this
   * series.
   */
  get bubbleSize() {
    const xml = pyFormat(this._bubbleSizeTmpl, { "nsdecls": ` ${nsdecls("c")}`, "numRef_xml": this.numRefXml(this._series.bubbleSizesRef, this._series.numberFormat, this._series.bubbleSizes) });
    return parseXml(xml);
  }

  /**
   * Return the ``<c:bubbleSize>`` element for this series as unicode
   * text. This element contains the bubble size values for all the
   * data points in the chart.
   */
  get bubbleSizeXml() {
    return pyFormat(this._bubbleSizeTmpl, { "nsdecls": "", "numRef_xml": this.numRefXml(this._series.bubbleSizesRef, this._series.numberFormat, this._series.bubbleSizes) });
  }

  /**
   * The template for the ``<c:bubbleSize>`` element for this series,
   * containing the bubble size values and their spreadsheet range
   * reference.
   */
  get _bubbleSizeTmpl() {
    return "          <c:bubbleSize{nsdecls}>\n{numRef_xml}          </c:bubbleSize>\n";
  }
}

/** A series rewriter suitable for bubble charts. */
export class _BubbleSeriesXmlRewriter extends _BaseSeriesXmlRewriter {

  /**
   * Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
   * of *ser* based on the values in *series_data*.
   */
  _rewriteSerData(ser: any, seriesData: any, date1904: any) {
    ser._removeTx();
    ser._removeXVal();
    ser._removeYVal();
    ser._removeBubbleSize();
    const xmlWriter = new _BubbleSeriesXmlWriter(seriesData);
    ser._insertTx(xmlWriter.tx);
    ser._insertXVal(xmlWriter.xVal);
    ser._insertYVal(xmlWriter.yVal);
    ser._insertBubbleSize(xmlWriter.bubbleSize);
  }
}

/** A series rewriter suitable for category charts. */
export class _CategorySeriesXmlRewriter extends _BaseSeriesXmlRewriter {

  /**
   * Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
   * of *ser* based on the values in *series_data*.
   */
  _rewriteSerData(ser: any, seriesData: any, date1904: any) {
    ser._removeTx();
    ser._removeCat();
    ser._removeVal();
    const xmlWriter = new _CategorySeriesXmlWriter(seriesData, date1904);
    ser._insertTx(xmlWriter.tx);
    ser._insertCat(xmlWriter.cat);
    ser._insertVal(xmlWriter.val);
  }
}

/** A series rewriter suitable for XY (aka. scatter) charts. */
export class _XySeriesXmlRewriter extends _BaseSeriesXmlRewriter {

  /**
   * Rewrite the ``<c:tx>``, ``<c:xVal>`` and ``<c:yVal>`` child elements
   * of *ser* based on the values in *series_data*.
   */
  _rewriteSerData(ser: any, seriesData: any, date1904: any) {
    ser._removeTx();
    ser._removeXVal();
    ser._removeYVal();
    const xmlWriter = new _XySeriesXmlWriter(seriesData);
    ser._insertTx(xmlWriter.tx);
    ser._insertXVal(xmlWriter.xVal);
    ser._insertYVal(xmlWriter.yVal);
  }
}
