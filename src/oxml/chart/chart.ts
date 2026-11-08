// Custom element classes for top-level chart-related XML elements. (`pptx.oxml.chart.chart`)
import { defineChildren } from "../xmlchemy.js";
import { pyBool } from "../../util.js";
import { parseXml } from "../parser.js";
import { CT_Title } from "./shared.js";
import { nsdecls, qn } from "../ns.js";
import { ST_Style, XsdString } from "../simpletypes.js";
import { CT_TextBody } from "../text.js";
import { BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne } from "../xmlchemy.js";

/** `c:chart` custom element class. */
export class CT_Chart extends BaseOxmlElement {
  declare title: any;
  declare getOrAddTitle: () => any;
  declare _addTitle: (attrs?: Record<string, any>) => any;
  declare _insertTitle: (child: any) => any;
  declare _removeTitle: () => void;
  declare autoTitleDeleted: any;
  declare getOrAddAutoTitleDeleted: () => any;
  declare _addAutoTitleDeleted: (attrs?: Record<string, any>) => any;
  declare _insertAutoTitleDeleted: (child: any) => any;
  declare _newAutoTitleDeleted: () => any;
  declare _removeAutoTitleDeleted: () => void;
  declare plotArea: any;
  declare legend: any;
  declare getOrAddLegend: () => any;
  declare _addLegend: (attrs?: Record<string, any>) => any;
  declare _insertLegend: (child: any) => any;
  declare _newLegend: () => any;
  declare _removeLegend: () => void;
  declare rId: string;
  static {
    const _tagSeq = ["c:title", "c:autoTitleDeleted", "c:pivotFmts", "c:view3D", "c:floor", "c:sideWall", "c:backWall", "c:plotArea", "c:legend", "c:plotVisOnly", "c:dispBlanksAs", "c:showDLblsOverMax", "c:extLst"] as const;
    defineChildren(this, {
      title: new ZeroOrOne("c:title", _tagSeq.slice(1)),
      autoTitleDeleted: new ZeroOrOne("c:autoTitleDeleted", _tagSeq.slice(2)),
      plotArea: new OneAndOnlyOne("c:plotArea"),
      legend: new ZeroOrOne("c:legend", _tagSeq.slice(9)),
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }

  /** True if this chart has a legend defined, False otherwise. */
  get hasLegend() {
    const legend = this.legend;
    if (legend === null) {
      return false;
    }
    return true;
  }

  set hasLegend(boolValue: any) {
    if (pyBool(boolValue) === false) {
      this._removeLegend();
    } else if (this.legend === null) {
      this._addLegend();
    }
  }

  /** Return a new `c:chart` element. */
  static newChart(rId: string): CT_Chart {
    return parseXml(`<c:chart ${nsdecls("c")} ${nsdecls("r")} r:id="${rId}"/>`) as CT_Chart;
  }

  _newTitle() {
    return CT_Title.newTitle();
  }
}

/** `c:chartSpace` root element of a chart part. */
export class CT_ChartSpace extends BaseOxmlElement {
  declare date1904: any;
  declare getOrAddDate1904: () => any;
  declare _addDate1904: (attrs?: Record<string, any>) => any;
  declare _insertDate1904: (child: any) => any;
  declare _newDate1904: () => any;
  declare _removeDate1904: () => void;
  declare style: any;
  declare getOrAddStyle: () => any;
  declare _addStyle: (attrs?: Record<string, any>) => any;
  declare _insertStyle: (child: any) => any;
  declare _newStyle: () => any;
  declare _removeStyle: () => void;
  declare chart: any;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _removeTxPr: () => void;
  declare externalData: any;
  declare getOrAddExternalData: () => any;
  declare _insertExternalData: (child: any) => any;
  declare _newExternalData: () => any;
  declare _removeExternalData: () => void;
  static {
    const _tagSeq = ["c:date1904", "c:lang", "c:roundedCorners", "c:style", "c:clrMapOvr", "c:pivotSource", "c:protection", "c:chart", "c:spPr", "c:txPr", "c:externalData", "c:printSettings", "c:userShapes", "c:extLst"] as const;
    defineChildren(this, {
      date1904: new ZeroOrOne("c:date1904", _tagSeq.slice(1)),
      style: new ZeroOrOne("c:style", _tagSeq.slice(4)),
      chart: new OneAndOnlyOne("c:chart"),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(10)),
      externalData: new ZeroOrOne("c:externalData", _tagSeq.slice(11)),
    });
  }

  get catAxLst() {
    return this.chart.plotArea.catAxLst;
  }

  /**
   * Return |True| if the `c:date1904` child element resolves truthy,
   * |False| otherwise. This value indicates whether date number values
   * are based on the 1900 or 1904 epoch.
   */
  get date1904Val(): boolean {
    const date1904 = this.date1904;
    if (date1904 === null) {
      return false;
    }
    return date1904.val;
  }

  get dateAxLst() {
    return this.xpath("c:chart/c:plotArea/c:dateAx");
  }

  /** Return the `c:title` grandchild, newly created if not present. */
  getOrAddTitle() {
    return this.chart.getOrAddTitle();
  }

  /**
   * Return the required `c:chartSpace/c:chart/c:plotArea` grandchild
   * element.
   */
  get plotArea() {
    return this.chart.plotArea;
  }

  get valAxLst() {
    return this.chart.plotArea.valAxLst;
  }

  /**
   * The string in the required ``r:id`` attribute of the
   * `<c:externalData>` child, or |None| if no externalData element is
   * present.
   */
  get xlsxPartRId() {
    const externalData = this.externalData;
    if (externalData === null) {
      return null;
    }
    return externalData.rId;
  }

  /**
   * Always add a ``<c:autoUpdate val="0"/>`` child so auto-updating
   * behavior is off by default.
   */
  _addExternalData() {
    const externalData = this._newExternalData();
    externalData._addAutoUpdate({ val: false });
    this._insertExternalData(externalData);
    return externalData;
  }

  _newTxPr() {
    return CT_TextBody.newTxPr();
  }
}

/**
 * `<c:externalData>` element, defining link to embedded Excel package part
 * containing the chart data.
 */
export class CT_ExternalData extends BaseOxmlElement {
  declare autoUpdate: any;
  declare getOrAddAutoUpdate: () => any;
  declare _addAutoUpdate: (attrs?: Record<string, any>) => any;
  declare _insertAutoUpdate: (child: any) => any;
  declare _newAutoUpdate: () => any;
  declare _removeAutoUpdate: () => void;
  declare rId: any;
  static {
    defineChildren(this, {
      autoUpdate: new ZeroOrOne("c:autoUpdate"),
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }
}

/** ``<c:plotArea>`` element. */
export class CT_PlotArea extends BaseOxmlElement {
  declare catAxLst: any[];
  declare _addCatAx: (attrs?: Record<string, any>) => any;
  declare _insertCatAx: (child: any) => any;
  declare _newCatAx: () => any;
  declare valAxLst: any[];
  declare _addValAx: (attrs?: Record<string, any>) => any;
  declare _insertValAx: (child: any) => any;
  declare _newValAx: () => any;
  static {
    defineChildren(this, {
      catAx: new ZeroOrMore("c:catAx"),
      valAx: new ZeroOrMore("c:valAx"),
    });
  }

  /**
   * Generate each of the `c:ser` elements in this chart, ordered first by
   * the document order of the containing xChart element, then by their
   * ordering within the xChart element (not necessarily document order).
   */
  *iterSers() {
    for (const xChart of this.iterXCharts() as Iterable<any>) {
      for (const ser of xChart.iterSers()) {
        yield ser;
      }
    }
  }

  /** Generate each xChart child element in document. */
  *iterXCharts() {
    const plotTags = [qn("c:area3DChart"), qn("c:areaChart"), qn("c:bar3DChart"), qn("c:barChart"), qn("c:bubbleChart"), qn("c:doughnutChart"), qn("c:line3DChart"), qn("c:lineChart"), qn("c:ofPieChart"), qn("c:pie3DChart"), qn("c:pieChart"), qn("c:radarChart"), qn("c:scatterChart"), qn("c:stockChart"), qn("c:surface3DChart"), qn("c:surfaceChart")];
    for (const child of this.iterchildren()) {
      if (!plotTags.includes(child.tag)) {
        continue;
      }
      yield child;
    }
  }

  /**
   * Return the last `<c:ser>` element in the last xChart element, based
   * on series order (not necessarily the same element as document order).
   */
  get lastSer() {
    const lastXChart: any = this.xCharts.at(-1);
    const sers = lastXChart.sers;
    if (sers.length === 0) {
      return null;
    }
    return sers.at(-1);
  }

  /**
   * Return the next available `c:ser/c:idx` value within the scope of
   * this chart, the maximum idx value found on existing series,
   * incremented by one.
   */
  get nextIdx() {
    const idxVals = [...this.sers].map((s) => s.idx.val);
    if (idxVals.length === 0) {
      return 0;
    }
    return (Math.max(...idxVals) + 1);
  }

  /**
   * Return the next available `c:ser/c:order` value within the scope of
   * this chart, the maximum order value found on existing series,
   * incremented by one.
   */
  get nextOrder() {
    const orderVals = [...this.sers].map((s) => s.order.val);
    if (orderVals.length === 0) {
      return 0;
    }
    return (Math.max(...orderVals) + 1);
  }

  /**
   * Return a sequence containing all the `c:ser` elements in this chart,
   * ordered first by the document order of the containing xChart element,
   * then by their ordering within the xChart element (not necessarily
   * document order).
   */
  get sers() {
    return [...this.iterSers()];
  }

  /**
   * Return a sequence containing all the `c:{x}Chart` elements in this
   * chart, in document order.
   */
  get xCharts() {
    return [...this.iterXCharts()];
  }
}

/** ``<c:style>`` element; defines the chart style. */
export class CT_Style extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_Style),
    });
  }
}
