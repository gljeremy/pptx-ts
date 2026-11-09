// Plot-related oxml objects. (`pptx.oxml.chart.plot`)
import { defineChildren } from "../xmlchemy.js";
import { pyCmp } from "../../util.js";
import { CT_DLbls } from "./datalabel.js";
import { ST_BarDir, ST_BubbleScale, ST_GapAmount, ST_Grouping, ST_Overlap } from "../simpletypes.js";
import { BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, ZeroOrMore, ZeroOrOne } from "../xmlchemy.js";

/** Base class for barChart, lineChart, and other plot elements. */
export class BaseChartElement extends BaseOxmlElement {
  declare grouping: any;

  /**
   * Return the `c:cat` element of the first series in this xChart, or
   * |None| if not present.
   */
  get cat() {
    const cats = this.xpath("./c:ser[1]/c:cat");
    return (cats.length ? cats[0] : null);
  }

  /**
   * Return the value of the `c:ptCount` descendent of this xChart
   * element. Its parent can be one of three element types. This value
   * represents the true number of (leaf) categories, although they might
   * not all have a corresponding `c:pt` sibling; a category with no label
   * does not get a `c:pt` element. Returns 0 if there is no `c:ptCount`
   * descendent.
   */
  get catPtCount() {
    const catPtCounts = this.xpath("./c:ser//c:cat//c:ptCount");
    if (catPtCounts.length === 0) {
      return 0;
    }
    return catPtCounts[0].val;
  }

  /**
   * Return a sequence representing the `c:pt` elements under the `c:cat`
   * element of the first series in this xChart element. A category having
   * no value will have no corresponding `c:pt` element; |None| will
   * appear in that position in such cases. Items appear in `idx` order.
   * Only those in the first ``<c:lvl>`` element are included in the case
   * of multi-level categories.
   */
  get catPts() {
    let catPts;
    catPts = this.xpath("./c:ser[1]/c:cat//c:lvl[1]/c:pt");
    if (catPts.length === 0) {
      catPts = this.xpath("./c:ser[1]/c:cat//c:pt");
    }
    const catPtDict = new Map<number, any>([...catPts].map((pt: any) => [pt.idx, pt]));
    return Array.from({ length: this.catPtCount }, (_, idx) => catPtDict.get(idx) ?? null);
  }

  /**
   * Return the value of the ``./c:grouping{val=?}`` attribute, taking
   * defaults into account when items are not present.
   */
  get groupingVal() {
    const grouping = this.grouping;
    if (grouping === null) {
      return ST_Grouping.STANDARD;
    }
    const val = grouping.val;
    if (val === null) {
      return ST_Grouping.STANDARD;
    }
    return val;
  }

  /**
   * Generate each ``<c:ser>`` child element in this xChart in
   * c:order/@val sequence (not document or c:idx order).
   */
  iterSers() {
    const serOrder = (ser: any) => {
      return ser.order.val;
    };
    return [...[...this.xpath("./c:ser")].sort((a, b) => pyCmp((serOrder)(a), (serOrder)(b)))];
  }

  /**
   * Sequence of ``<c:ser>`` child elements in this xChart in c:order/@val
   * sequence (not document or c:idx order).
   */
  get sers() {
    return [...this.iterSers()];
  }

  _newDLbls() {
    return CT_DLbls.newDLbls();
  }
}

/** ``<c:area3DChart>`` element. */
export class CT_Area3DChart extends BaseChartElement {
  declare grouping: any;
  declare getOrAddGrouping: () => any;
  declare _addGrouping: (attrs?: Record<string, any>) => any;
  declare _insertGrouping: (child: any) => any;
  declare _newGrouping: () => any;
  declare _removeGrouping: () => void;
  static {
    defineChildren(this, {
      grouping: new ZeroOrOne("c:grouping", ["c:varyColors", "c:ser", "c:dLbls", "c:dropLines", "c:gapDepth", "c:axId"]),
    });
  }
}

/** ``<c:areaChart>`` element. */
export class CT_AreaChart extends BaseChartElement {
  declare grouping: any;
  declare getOrAddGrouping: () => any;
  declare _addGrouping: (attrs?: Record<string, any>) => any;
  declare _insertGrouping: (child: any) => any;
  declare _newGrouping: () => any;
  declare _removeGrouping: () => void;
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  static {
    const _tagSeq = ["c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:dropLines", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      grouping: new ZeroOrOne("c:grouping", _tagSeq.slice(1)),
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(2)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(3)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(4)),
    });
  }
}

/** ``<c:barChart>`` element. */
export class CT_BarChart extends BaseChartElement {
  declare barDir: any;
  declare grouping: any;
  declare getOrAddGrouping: () => any;
  declare _addGrouping: (attrs?: Record<string, any>) => any;
  declare _insertGrouping: (child: any) => any;
  declare _newGrouping: () => any;
  declare _removeGrouping: () => void;
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  declare gapWidth: any;
  declare getOrAddGapWidth: () => any;
  declare _addGapWidth: (attrs?: Record<string, any>) => any;
  declare _insertGapWidth: (child: any) => any;
  declare _newGapWidth: () => any;
  declare _removeGapWidth: () => void;
  declare overlap: any;
  declare getOrAddOverlap: () => any;
  declare _addOverlap: (attrs?: Record<string, any>) => any;
  declare _insertOverlap: (child: any) => any;
  declare _newOverlap: () => any;
  declare _removeOverlap: () => void;
  static {
    const _tagSeq = ["c:barDir", "c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:gapWidth", "c:overlap", "c:serLines", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      barDir: new OneAndOnlyOne("c:barDir"),
      grouping: new ZeroOrOne("c:grouping", _tagSeq.slice(2)),
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(3)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(4)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(5)),
      gapWidth: new ZeroOrOne("c:gapWidth", _tagSeq.slice(6)),
      overlap: new ZeroOrOne("c:overlap", _tagSeq.slice(7)),
    });
  }

  /**
   * Return the value of the ``./c:grouping{val=?}`` attribute, taking
   * defaults into account when items are not present.
   */
  get groupingVal() {
    const grouping = this.grouping;
    if (grouping === null) {
      return ST_Grouping.CLUSTERED;
    }
    const val = grouping.val;
    if (val === null) {
      return ST_Grouping.CLUSTERED;
    }
    return val;
  }
}

/**
 * ``<c:barDir>`` child of a barChart element, specifying the orientation of
 * the bars, 'bar' if they are horizontal and 'col' if they are vertical.
 */
export class CT_BarDir extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_BarDir, ST_BarDir.COL),
    });
  }
}

/** ``<c:bubbleChart>`` custom element class */
export class CT_BubbleChart extends BaseChartElement {
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  declare bubble3D: any;
  declare getOrAddBubble3D: () => any;
  declare _addBubble3D: (attrs?: Record<string, any>) => any;
  declare _insertBubble3D: (child: any) => any;
  declare _newBubble3D: () => any;
  declare _removeBubble3D: () => void;
  declare bubbleScale: any;
  declare getOrAddBubbleScale: () => any;
  declare _addBubbleScale: (attrs?: Record<string, any>) => any;
  declare _insertBubbleScale: (child: any) => any;
  declare _newBubbleScale: () => any;
  declare _removeBubbleScale: () => void;
  static {
    const _tagSeq = ["c:varyColors", "c:ser", "c:dLbls", "c:axId", "c:bubble3D", "c:bubbleScale", "c:showNegBubbles", "c:sizeRepresents", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(2)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(3)),
      bubble3D: new ZeroOrOne("c:bubble3D", _tagSeq.slice(5)),
      bubbleScale: new ZeroOrOne("c:bubbleScale", _tagSeq.slice(6)),
    });
  }
}

/** ``<c:bubbleScale>`` custom element class */
export class CT_BubbleScale extends BaseChartElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_BubbleScale, 100),
    });
  }
}

/** ``<c:doughnutChart>`` element. */
export class CT_DoughnutChart extends BaseChartElement {
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  static {
    const _tagSeq = ["c:varyColors", "c:ser", "c:dLbls", "c:firstSliceAng", "c:holeSize", "c:extLst"] as const;
    defineChildren(this, {
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(1)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(2)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(3)),
    });
  }
}

/**
 * ``<c:gapWidth>`` child of ``<c:barChart>`` element, also used for other
 * purposes like error bars.
 */
export class CT_GapAmount extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_GapAmount, 150),
    });
  }
}

/**
 * ``<c:grouping>`` child of an xChart element, specifying a value like
 * 'clustered' or 'stacked'. Also used for variants with the same tag name
 * like CT_BarGrouping.
 */
export class CT_Grouping extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_Grouping),
    });
  }
}

/** ``<c:lineChart>`` custom element class */
export class CT_LineChart extends BaseChartElement {
  declare grouping: any;
  declare getOrAddGrouping: () => any;
  declare _addGrouping: (attrs?: Record<string, any>) => any;
  declare _insertGrouping: (child: any) => any;
  declare _newGrouping: () => any;
  declare _removeGrouping: () => void;
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  static {
    const _tagSeq = ["c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:dropLines", "c:hiLowLines", "c:upDownBars", "c:marker", "c:smooth", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      grouping: new ZeroOrOne("c:grouping", _tagSeq.slice(1)),
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(2)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(3)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(4)),
    });
  }
}

/**
 * ``<c:overlap>`` element specifying bar overlap as an integer percentage
 * of bar width, in range -100 to 100.
 */
export class CT_Overlap extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_Overlap, 0),
    });
  }
}

/** ``<c:pieChart>`` custom element class */
export class CT_PieChart extends BaseChartElement {
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  static {
    const _tagSeq = ["c:varyColors", "c:ser", "c:dLbls", "c:firstSliceAng", "c:extLst"] as const;
    defineChildren(this, {
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(1)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(2)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(3)),
    });
  }
}

/** ``<c:radarChart>`` custom element class */
export class CT_RadarChart extends BaseChartElement {
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _newDLbls: () => any;
  declare _removeDLbls: () => void;
  static {
    const _tagSeq = ["c:radarStyle", "c:varyColors", "c:ser", "c:dLbls", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(2)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(3)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(4)),
    });
  }
}

/** ``<c:scatterChart>`` custom element class */
export class CT_ScatterChart extends BaseChartElement {
  declare varyColors: any;
  declare getOrAddVaryColors: () => any;
  declare _addVaryColors: (attrs?: Record<string, any>) => any;
  declare _insertVaryColors: (child: any) => any;
  declare _newVaryColors: () => any;
  declare _removeVaryColors: () => void;
  declare serLst: any[];
  declare _addSer: (attrs?: Record<string, any>) => any;
  declare _insertSer: (child: any) => any;
  declare _newSer: () => any;
  static {
    const _tagSeq = ["c:scatterStyle", "c:varyColors", "c:ser", "c:dLbls", "c:axId", "c:extLst"] as const;
    defineChildren(this, {
      varyColors: new ZeroOrOne("c:varyColors", _tagSeq.slice(2)),
      ser: new ZeroOrMore("c:ser", _tagSeq.slice(3)),
    });
  }
}
