// Series-related oxml objects. (`pptx.oxml.chart.series`)
import { defineChildren } from "../xmlchemy.js";
import { pyInt, pyFloat } from "../../util.js";
import { CT_DLbls } from "./datalabel.js";
import { XsdUnsignedInt } from "../simpletypes.js";
import { BaseOxmlElement, OneAndOnlyOne, OxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne } from "../xmlchemy.js";

/**
 * ``<c:cat>`` custom element class used in category charts to specify
 * category labels and hierarchy.
 */
export class CT_AxDataSource extends BaseOxmlElement {
  declare multiLvlStrRef: any;
  declare getOrAddMultiLvlStrRef: () => any;
  declare _addMultiLvlStrRef: (attrs?: Record<string, any>) => any;
  declare _insertMultiLvlStrRef: (child: any) => any;
  declare _newMultiLvlStrRef: () => any;
  declare _removeMultiLvlStrRef: () => void;
  static {
    defineChildren(this, {
      multiLvlStrRef: new ZeroOrOne("c:multiLvlStrRef", []),
    });
  }

  /**
   * Return a list containing the `c:lvl` descendent elements in document
   * order. These will only be present when the required single child
   * is a `c:multiLvlStrRef` element. Returns an empty list when no
   * `c:lvl` descendent elements are present.
   */
  get lvls() {
    return this.xpath(".//c:lvl");
  }
}

/**
 * ``<c:dPt>`` custom element class, containing visual properties for a data
 * point.
 */
export class CT_DPt extends BaseOxmlElement {
  declare idx: any;
  declare marker: any;
  declare getOrAddMarker: () => any;
  declare _addMarker: (attrs?: Record<string, any>) => any;
  declare _insertMarker: (child: any) => any;
  declare _newMarker: () => any;
  declare _removeMarker: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  static {
    const _tagSeq = ["c:idx", "c:invertIfNegative", "c:marker", "c:bubble3D", "c:explosion", "c:spPr", "c:pictureOptions", "c:extLst"] as const;
    defineChildren(this, {
      idx: new OneAndOnlyOne("c:idx"),
      marker: new ZeroOrOne("c:marker", _tagSeq.slice(3)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(6)),
    });
  }

  /**
   * Return a newly created "loose" `c:dPt` element containing its default
   * subtree.
   */
  static newDPt() {
    const dPt = OxmlElement("c:dPt");
    dPt.append(OxmlElement("c:idx"));
    return dPt;
  }
}

/**
 * ``<c:lvl>`` custom element class used in multi-level categories to
 * specify a level of hierarchy.
 */
export class CT_Lvl extends BaseOxmlElement {
  declare ptLst: any[];
  declare _addPt: (attrs?: Record<string, any>) => any;
  declare _insertPt: (child: any) => any;
  declare _newPt: () => any;
  static {
    defineChildren(this, {
      pt: new ZeroOrMore("c:pt", []),
    });
  }
}

/**
 * ``<c:yVal>`` custom element class used in XY and bubble charts, and
 * perhaps others.
 */
export class CT_NumDataSource extends BaseOxmlElement {
  declare numRef: any;
  static {
    defineChildren(this, {
      numRef: new OneAndOnlyOne("c:numRef"),
    });
  }

  /**
   * Return the value of `./c:numRef/c:numCache/c:ptCount/@val`,
   * specifying how many `c:pt` elements are in this numeric data cache.
   * Returns 0 if no `c:ptCount` element is present, as this is the least
   * disruptive way to degrade when no cached point data is available.
   * This situation is not expected, but is valid according to the schema.
   */
  get ptCountVal() {
    const results = this.xpath(".//c:ptCount/@val");
    return (results.length ? pyInt(results[0]) : 0);
  }

  /**
   * Return the Y value for data point *idx* in this cache, or None if no
   * value is present for that data point.
   */
  ptV(idx: any) {
    const results = this.xpath(`.//c:pt[@idx=${idx}]`);
    return (results.length ? results[0].value : null);
  }
}

/**
 * ``<c:ser>`` custom element class. Note there are several different series
 * element types in the schema, such as ``CT_LineSer`` and ``CT_BarSer``,
 * but they all share the same tag name. This class acts as a composite and
 * depends on the caller not to do anything invalid for a series belonging
 * to a particular plot type.
 */
export class CT_SeriesComposite extends BaseOxmlElement {
  declare idx: any;
  declare order: any;
  declare tx: any;
  declare getOrAddTx: () => any;
  declare _addTx: (attrs?: Record<string, any>) => any;
  declare _insertTx: (child: any) => any;
  declare _newTx: () => any;
  declare _removeTx: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  declare invertIfNegative: any;
  declare getOrAddInvertIfNegative: () => any;
  declare _addInvertIfNegative: (attrs?: Record<string, any>) => any;
  declare _insertInvertIfNegative: (child: any) => any;
  declare _newInvertIfNegative: () => any;
  declare _removeInvertIfNegative: () => void;
  declare marker: any;
  declare getOrAddMarker: () => any;
  declare _addMarker: (attrs?: Record<string, any>) => any;
  declare _insertMarker: (child: any) => any;
  declare _newMarker: () => any;
  declare _removeMarker: () => void;
  declare dPtLst: any[];
  declare _addDPt: (attrs?: Record<string, any>) => any;
  declare _insertDPt: (child: any) => any;
  declare dLbls: any;
  declare getOrAddDLbls: () => any;
  declare _addDLbls: (attrs?: Record<string, any>) => any;
  declare _insertDLbls: (child: any) => any;
  declare _removeDLbls: () => void;
  declare cat: any;
  declare getOrAddCat: () => any;
  declare _addCat: (attrs?: Record<string, any>) => any;
  declare _insertCat: (child: any) => any;
  declare _newCat: () => any;
  declare _removeCat: () => void;
  declare val: any;
  declare getOrAddVal: () => any;
  declare _addVal: (attrs?: Record<string, any>) => any;
  declare _insertVal: (child: any) => any;
  declare _newVal: () => any;
  declare _removeVal: () => void;
  declare xVal: any;
  declare getOrAddXVal: () => any;
  declare _addXVal: (attrs?: Record<string, any>) => any;
  declare _insertXVal: (child: any) => any;
  declare _newXVal: () => any;
  declare _removeXVal: () => void;
  declare yVal: any;
  declare getOrAddYVal: () => any;
  declare _addYVal: (attrs?: Record<string, any>) => any;
  declare _insertYVal: (child: any) => any;
  declare _newYVal: () => any;
  declare _removeYVal: () => void;
  declare smooth: any;
  declare getOrAddSmooth: () => any;
  declare _addSmooth: (attrs?: Record<string, any>) => any;
  declare _insertSmooth: (child: any) => any;
  declare _newSmooth: () => any;
  declare _removeSmooth: () => void;
  declare bubbleSize: any;
  declare getOrAddBubbleSize: () => any;
  declare _addBubbleSize: (attrs?: Record<string, any>) => any;
  declare _insertBubbleSize: (child: any) => any;
  declare _newBubbleSize: () => any;
  declare _removeBubbleSize: () => void;
  static {
    const _tagSeq = ["c:idx", "c:order", "c:tx", "c:spPr", "c:invertIfNegative", "c:pictureOptions", "c:marker", "c:explosion", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:cat", "c:val", "c:xVal", "c:yVal", "c:shape", "c:smooth", "c:bubbleSize", "c:bubble3D", "c:extLst"] as const;
    defineChildren(this, {
      idx: new OneAndOnlyOne("c:idx"),
      order: new OneAndOnlyOne("c:order"),
      tx: new ZeroOrOne("c:tx", _tagSeq.slice(3)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(4)),
      invertIfNegative: new ZeroOrOne("c:invertIfNegative", _tagSeq.slice(5)),
      marker: new ZeroOrOne("c:marker", _tagSeq.slice(7)),
      dPt: new ZeroOrMore("c:dPt", _tagSeq.slice(9)),
      dLbls: new ZeroOrOne("c:dLbls", _tagSeq.slice(10)),
      cat: new ZeroOrOne("c:cat", _tagSeq.slice(13)),
      val: new ZeroOrOne("c:val", _tagSeq.slice(14)),
      xVal: new ZeroOrOne("c:xVal", _tagSeq.slice(15)),
      yVal: new ZeroOrOne("c:yVal", _tagSeq.slice(16)),
      smooth: new ZeroOrOne("c:smooth", _tagSeq.slice(18)),
      bubbleSize: new ZeroOrOne("c:bubbleSize", _tagSeq.slice(19)),
    });
  }

  /**
   * Return the number of bubble size values as reflected in the `val`
   * attribute of `./c:bubbleSize//c:ptCount`, or 0 if not present.
   */
  get bubbleSizePtCountVal() {
    const vals = this.xpath("./c:bubbleSize//c:ptCount/@val");
    if (vals.length === 0) {
      return 0;
    }
    return pyInt(vals[0]);
  }

  /**
   * Return the number of categories as reflected in the `val` attribute
   * of `./c:cat//c:ptCount`, or 0 if not present.
   */
  get catPtCountVal() {
    const vals = this.xpath("./c:cat//c:ptCount/@val");
    if (vals.length === 0) {
      return 0;
    }
    return pyInt(vals[0]);
  }

  /**
   * Return the `c:dLbl` element representing the label for the data point
   * at offset *idx* in this series, or |None| if not present.
   */
  getDLbl(idx: any) {
    const dLbls = this.dLbls;
    if (dLbls === null) {
      return null;
    }
    return dLbls.getDLblForPoint(idx);
  }

  /**
   * Return the `c:dLbl` element representing the label of the point at
   * offset *idx* in this series, newly created if not yet present.
   */
  getOrAddDLbl(idx: any) {
    const dLbls = this.getOrAddDLbls();
    return dLbls.getOrAddDLblForPoint(idx);
  }

  /**
   * Return the `c:dPt` child representing the visual properties of the
   * data point at index *idx*.
   */
  getOrAddDPtForPoint(idx: any) {
    const matches = this.xpath(`c:dPt[c:idx[@val="${idx}"]]`);
    if (matches.length) {
      return matches[0];
    }
    const dPt = this._addDPt();
    dPt.idx.val = idx;
    return dPt;
  }

  /**
   * Return the number of X values as reflected in the `val` attribute of
   * `./c:xVal//c:ptCount`, or 0 if not present.
   */
  get xValPtCountVal() {
    const vals = this.xpath("./c:xVal//c:ptCount/@val");
    if (vals.length === 0) {
      return 0;
    }
    return pyInt(vals[0]);
  }

  /**
   * Return the number of Y values as reflected in the `val` attribute of
   * `./c:yVal//c:ptCount`, or 0 if not present.
   */
  get yValPtCountVal() {
    const vals = this.xpath("./c:yVal//c:ptCount/@val");
    if (vals.length === 0) {
      return 0;
    }
    return pyInt(vals[0]);
  }

  /** Override metaclass method that creates `c:dLbls` element. */
  _newDLbls() {
    return CT_DLbls.newDLbls();
  }

  /**
   * Overrides the metaclass generated method to get `c:dPt` with minimal
   * subtree.
   */
  _newDPt() {
    return CT_DPt.newDPt();
  }
}

/**
 * ``<c:pt>`` element, can be either CT_StrVal or CT_NumVal complex type.
 * Using this class for both, differentiating as needed.
 */
export class CT_StrVal_NumVal_Composite extends BaseOxmlElement {
  declare v: any;
  declare idx: any;
  static {
    defineChildren(this, {
      v: new OneAndOnlyOne("c:v"),
      idx: new RequiredAttribute("idx", XsdUnsignedInt),
    });
  }

  /** The float value of the text in the required ``<c:v>`` child. */
  get value() {
    return pyFloat(this.v.text);
  }
}
