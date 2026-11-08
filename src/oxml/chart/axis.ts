// Axis-related oxml objects. (`pptx.oxml.chart.axis`)
import { defineChildren } from "../xmlchemy.js";
import { XL_AXIS_CROSSES, XL_TICK_LABEL_POSITION, XL_TICK_MARK } from "../../enum/chart.js";
import { CT_Title } from "./shared.js";
import { ST_AxisUnit, ST_LblOffset, ST_Orientation } from "../simpletypes.js";
import { CT_TextBody } from "../text.js";
import { BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrOne } from "../xmlchemy.js";

/** Base class for catAx, dateAx, valAx, and perhaps other axis elements. */
export class BaseAxisElement extends BaseOxmlElement {
  // children every axis element type declares
  declare scaling: any;
  declare getOrAddTxPr: () => any;

  /**
   * ``<a:defRPr>`` great-great-grandchild element, added with its
   * ancestors if not present.
   */
  get defRPr() {
    const txPr = this.getOrAddTxPr();
    const defRPr = txPr.defRPr;
    return defRPr;
  }

  /**
   * Value of `val` attribute of `c:scaling/c:orientation` grandchild element.
   *
   * Defaults to `ST_Orientation.MIN_MAX` if attribute or any ancestors are not
   * present.
   */
  get orientation() {
    const orientation = this.scaling.orientation;
    if (orientation === null) {
      return ST_Orientation.MIN_MAX;
    }
    return orientation.val;
  }

  set orientation(value: any) {
    this.scaling._removeOrientation();
    if (value === ST_Orientation.MAX_MIN) {
      this.scaling.getOrAddOrientation().val = value;
    }
  }

  _newTitle() {
    return CT_Title.newTitle();
  }

  _newTxPr() {
    return CT_TextBody.newTxPr();
  }
}

/** Used for `c:majorUnit` and `c:minorUnit` elements, and others. */
export class CT_AxisUnit extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_AxisUnit),
    });
  }
}

/** `c:catAx` element, defining a category axis. */
export class CT_CatAx extends BaseAxisElement {
  declare scaling: any;
  declare delete_: any;
  declare getOrAddDelete_: () => any;
  declare _addDelete_: (attrs?: Record<string, any>) => any;
  declare _insertDelete_: (child: any) => any;
  declare _newDelete_: () => any;
  declare _removeDelete_: () => void;
  declare majorGridlines: any;
  declare getOrAddMajorGridlines: () => any;
  declare _addMajorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMajorGridlines: (child: any) => any;
  declare _newMajorGridlines: () => any;
  declare _removeMajorGridlines: () => void;
  declare minorGridlines: any;
  declare getOrAddMinorGridlines: () => any;
  declare _addMinorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMinorGridlines: (child: any) => any;
  declare _newMinorGridlines: () => any;
  declare _removeMinorGridlines: () => void;
  declare title: any;
  declare getOrAddTitle: () => any;
  declare _addTitle: (attrs?: Record<string, any>) => any;
  declare _insertTitle: (child: any) => any;
  declare _newTitle: () => any;
  declare _removeTitle: () => void;
  declare numFmt: any;
  declare getOrAddNumFmt: () => any;
  declare _addNumFmt: (attrs?: Record<string, any>) => any;
  declare _insertNumFmt: (child: any) => any;
  declare _newNumFmt: () => any;
  declare _removeNumFmt: () => void;
  declare majorTickMark: any;
  declare getOrAddMajorTickMark: () => any;
  declare _addMajorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMajorTickMark: (child: any) => any;
  declare _newMajorTickMark: () => any;
  declare _removeMajorTickMark: () => void;
  declare minorTickMark: any;
  declare getOrAddMinorTickMark: () => any;
  declare _addMinorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMinorTickMark: (child: any) => any;
  declare _newMinorTickMark: () => any;
  declare _removeMinorTickMark: () => void;
  declare tickLblPos: any;
  declare getOrAddTickLblPos: () => any;
  declare _addTickLblPos: (attrs?: Record<string, any>) => any;
  declare _insertTickLblPos: (child: any) => any;
  declare _newTickLblPos: () => any;
  declare _removeTickLblPos: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _newTxPr: () => any;
  declare _removeTxPr: () => void;
  declare crosses: any;
  declare getOrAddCrosses: () => any;
  declare _addCrosses: (attrs?: Record<string, any>) => any;
  declare _insertCrosses: (child: any) => any;
  declare _newCrosses: () => any;
  declare _removeCrosses: () => void;
  declare crossesAt: any;
  declare getOrAddCrossesAt: () => any;
  declare _addCrossesAt: (attrs?: Record<string, any>) => any;
  declare _insertCrossesAt: (child: any) => any;
  declare _newCrossesAt: () => any;
  declare _removeCrossesAt: () => void;
  declare lblOffset: any;
  declare getOrAddLblOffset: () => any;
  declare _addLblOffset: (attrs?: Record<string, any>) => any;
  declare _insertLblOffset: (child: any) => any;
  declare _newLblOffset: () => any;
  declare _removeLblOffset: () => void;
  static {
    const _tagSeq = ["c:axId", "c:scaling", "c:delete", "c:axPos", "c:majorGridlines", "c:minorGridlines", "c:title", "c:numFmt", "c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", "c:crosses", "c:crossesAt", "c:auto", "c:lblAlgn", "c:lblOffset", "c:tickLblSkip", "c:tickMarkSkip", "c:noMultiLvlLbl", "c:extLst"] as const;
    defineChildren(this, {
      scaling: new OneAndOnlyOne("c:scaling"),
      delete_: new ZeroOrOne("c:delete", _tagSeq.slice(3)),
      majorGridlines: new ZeroOrOne("c:majorGridlines", _tagSeq.slice(5)),
      minorGridlines: new ZeroOrOne("c:minorGridlines", _tagSeq.slice(6)),
      title: new ZeroOrOne("c:title", _tagSeq.slice(7)),
      numFmt: new ZeroOrOne("c:numFmt", _tagSeq.slice(8)),
      majorTickMark: new ZeroOrOne("c:majorTickMark", _tagSeq.slice(9)),
      minorTickMark: new ZeroOrOne("c:minorTickMark", _tagSeq.slice(10)),
      tickLblPos: new ZeroOrOne("c:tickLblPos", _tagSeq.slice(11)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(12)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(13)),
      crosses: new ZeroOrOne("c:crosses", _tagSeq.slice(15)),
      crossesAt: new ZeroOrOne("c:crossesAt", _tagSeq.slice(16)),
      lblOffset: new ZeroOrOne("c:lblOffset", _tagSeq.slice(19)),
    });
  }
}

/**
 * Used for `c:majorGridlines` and `c:minorGridlines`.
 *
 * Specifies gridlines visual properties such as color and width.
 */
export class CT_ChartLines extends BaseOxmlElement {
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  static {
    defineChildren(this, {
      spPr: new ZeroOrOne("c:spPr", []),
    });
  }
}

/** `c:crosses` element, specifying where the other axis crosses this one. */
export class CT_Crosses extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", XL_AXIS_CROSSES),
    });
  }
}

/** `c:dateAx` element, defining a date (category) axis. */
export class CT_DateAx extends BaseAxisElement {
  declare scaling: any;
  declare delete_: any;
  declare getOrAddDelete_: () => any;
  declare _addDelete_: (attrs?: Record<string, any>) => any;
  declare _insertDelete_: (child: any) => any;
  declare _newDelete_: () => any;
  declare _removeDelete_: () => void;
  declare majorGridlines: any;
  declare getOrAddMajorGridlines: () => any;
  declare _addMajorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMajorGridlines: (child: any) => any;
  declare _newMajorGridlines: () => any;
  declare _removeMajorGridlines: () => void;
  declare minorGridlines: any;
  declare getOrAddMinorGridlines: () => any;
  declare _addMinorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMinorGridlines: (child: any) => any;
  declare _newMinorGridlines: () => any;
  declare _removeMinorGridlines: () => void;
  declare title: any;
  declare getOrAddTitle: () => any;
  declare _addTitle: (attrs?: Record<string, any>) => any;
  declare _insertTitle: (child: any) => any;
  declare _newTitle: () => any;
  declare _removeTitle: () => void;
  declare numFmt: any;
  declare getOrAddNumFmt: () => any;
  declare _addNumFmt: (attrs?: Record<string, any>) => any;
  declare _insertNumFmt: (child: any) => any;
  declare _newNumFmt: () => any;
  declare _removeNumFmt: () => void;
  declare majorTickMark: any;
  declare getOrAddMajorTickMark: () => any;
  declare _addMajorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMajorTickMark: (child: any) => any;
  declare _newMajorTickMark: () => any;
  declare _removeMajorTickMark: () => void;
  declare minorTickMark: any;
  declare getOrAddMinorTickMark: () => any;
  declare _addMinorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMinorTickMark: (child: any) => any;
  declare _newMinorTickMark: () => any;
  declare _removeMinorTickMark: () => void;
  declare tickLblPos: any;
  declare getOrAddTickLblPos: () => any;
  declare _addTickLblPos: (attrs?: Record<string, any>) => any;
  declare _insertTickLblPos: (child: any) => any;
  declare _newTickLblPos: () => any;
  declare _removeTickLblPos: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _newTxPr: () => any;
  declare _removeTxPr: () => void;
  declare crosses: any;
  declare getOrAddCrosses: () => any;
  declare _addCrosses: (attrs?: Record<string, any>) => any;
  declare _insertCrosses: (child: any) => any;
  declare _newCrosses: () => any;
  declare _removeCrosses: () => void;
  declare crossesAt: any;
  declare getOrAddCrossesAt: () => any;
  declare _addCrossesAt: (attrs?: Record<string, any>) => any;
  declare _insertCrossesAt: (child: any) => any;
  declare _newCrossesAt: () => any;
  declare _removeCrossesAt: () => void;
  declare lblOffset: any;
  declare getOrAddLblOffset: () => any;
  declare _addLblOffset: (attrs?: Record<string, any>) => any;
  declare _insertLblOffset: (child: any) => any;
  declare _newLblOffset: () => any;
  declare _removeLblOffset: () => void;
  static {
    const _tagSeq = ["c:axId", "c:scaling", "c:delete", "c:axPos", "c:majorGridlines", "c:minorGridlines", "c:title", "c:numFmt", "c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", "c:crosses", "c:crossesAt", "c:auto", "c:lblOffset", "c:baseTimeUnit", "c:majorUnit", "c:majorTimeUnit", "c:minorUnit", "c:minorTimeUnit", "c:extLst"] as const;
    defineChildren(this, {
      scaling: new OneAndOnlyOne("c:scaling"),
      delete_: new ZeroOrOne("c:delete", _tagSeq.slice(3)),
      majorGridlines: new ZeroOrOne("c:majorGridlines", _tagSeq.slice(5)),
      minorGridlines: new ZeroOrOne("c:minorGridlines", _tagSeq.slice(6)),
      title: new ZeroOrOne("c:title", _tagSeq.slice(7)),
      numFmt: new ZeroOrOne("c:numFmt", _tagSeq.slice(8)),
      majorTickMark: new ZeroOrOne("c:majorTickMark", _tagSeq.slice(9)),
      minorTickMark: new ZeroOrOne("c:minorTickMark", _tagSeq.slice(10)),
      tickLblPos: new ZeroOrOne("c:tickLblPos", _tagSeq.slice(11)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(12)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(13)),
      crosses: new ZeroOrOne("c:crosses", _tagSeq.slice(15)),
      crossesAt: new ZeroOrOne("c:crossesAt", _tagSeq.slice(16)),
      lblOffset: new ZeroOrOne("c:lblOffset", _tagSeq.slice(18)),
    });
  }
}

/** `c:lblOffset` custom element class. */
export class CT_LblOffset extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_LblOffset, 100),
    });
  }
}

/**
 * `c:xAx/c:scaling/c:orientation` element, defining category order.
 *
 * Used to reverse the order categories appear in on a bar chart so they start at the
 * top rather than the bottom. Because we read top-to-bottom, the default way looks odd
 * to many and perhaps most folks. Also applicable to value and date axes.
 */
export class CT_Orientation extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_Orientation, ST_Orientation.MIN_MAX),
    });
  }
}

/**
 * `c:scaling` element.
 *
 * Defines axis scale characteristics such as maximum value, log vs. linear, etc.
 */
export class CT_Scaling extends BaseOxmlElement {
  declare orientation: any;
  declare getOrAddOrientation: () => any;
  declare _addOrientation: (attrs?: Record<string, any>) => any;
  declare _insertOrientation: (child: any) => any;
  declare _newOrientation: () => any;
  declare _removeOrientation: () => void;
  declare max: any;
  declare getOrAddMax: () => any;
  declare _addMax: (attrs?: Record<string, any>) => any;
  declare _insertMax: (child: any) => any;
  declare _newMax: () => any;
  declare _removeMax: () => void;
  declare min: any;
  declare getOrAddMin: () => any;
  declare _addMin: (attrs?: Record<string, any>) => any;
  declare _insertMin: (child: any) => any;
  declare _newMin: () => any;
  declare _removeMin: () => void;
  static {
    const _tagSeq = ["c:logBase", "c:orientation", "c:max", "c:min", "c:extLst"] as const;
    defineChildren(this, {
      orientation: new ZeroOrOne("c:orientation", _tagSeq.slice(2)),
      max: new ZeroOrOne("c:max", _tagSeq.slice(3)),
      min: new ZeroOrOne("c:min", _tagSeq.slice(4)),
    });
  }

  /**
   * The float value of the ``<c:max>`` child element, or |None| if no max
   * element is present.
   */
  get maximum() {
    const max = this.max;
    if (max === null) {
      return null;
    }
    return max.val;
  }

  set maximum(value: any) {
    this._removeMax();
    if (value === null) {
      return;
    }
    this._addMax({ val: value });
  }

  /**
   * The float value of the ``<c:min>`` child element, or |None| if no min
   * element is present.
   */
  get minimum() {
    const min = this.min;
    if (min === null) {
      return null;
    }
    return min.val;
  }

  set minimum(value: any) {
    this._removeMin();
    if (value === null) {
      return;
    }
    this._addMin({ val: value });
  }
}

/** `c:tickLblPos` element. */
export class CT_TickLblPos extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", XL_TICK_LABEL_POSITION),
    });
  }
}

/** Used for `c:minorTickMark` and `c:majorTickMark`. */
export class CT_TickMark extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", XL_TICK_MARK, XL_TICK_MARK.CROSS),
    });
  }
}

/** `c:valAx` element, defining a value axis. */
export class CT_ValAx extends BaseAxisElement {
  declare scaling: any;
  declare delete_: any;
  declare getOrAddDelete_: () => any;
  declare _addDelete_: (attrs?: Record<string, any>) => any;
  declare _insertDelete_: (child: any) => any;
  declare _newDelete_: () => any;
  declare _removeDelete_: () => void;
  declare majorGridlines: any;
  declare getOrAddMajorGridlines: () => any;
  declare _addMajorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMajorGridlines: (child: any) => any;
  declare _newMajorGridlines: () => any;
  declare _removeMajorGridlines: () => void;
  declare minorGridlines: any;
  declare getOrAddMinorGridlines: () => any;
  declare _addMinorGridlines: (attrs?: Record<string, any>) => any;
  declare _insertMinorGridlines: (child: any) => any;
  declare _newMinorGridlines: () => any;
  declare _removeMinorGridlines: () => void;
  declare title: any;
  declare getOrAddTitle: () => any;
  declare _addTitle: (attrs?: Record<string, any>) => any;
  declare _insertTitle: (child: any) => any;
  declare _newTitle: () => any;
  declare _removeTitle: () => void;
  declare numFmt: any;
  declare getOrAddNumFmt: () => any;
  declare _addNumFmt: (attrs?: Record<string, any>) => any;
  declare _insertNumFmt: (child: any) => any;
  declare _newNumFmt: () => any;
  declare _removeNumFmt: () => void;
  declare majorTickMark: any;
  declare getOrAddMajorTickMark: () => any;
  declare _addMajorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMajorTickMark: (child: any) => any;
  declare _newMajorTickMark: () => any;
  declare _removeMajorTickMark: () => void;
  declare minorTickMark: any;
  declare getOrAddMinorTickMark: () => any;
  declare _addMinorTickMark: (attrs?: Record<string, any>) => any;
  declare _insertMinorTickMark: (child: any) => any;
  declare _newMinorTickMark: () => any;
  declare _removeMinorTickMark: () => void;
  declare tickLblPos: any;
  declare getOrAddTickLblPos: () => any;
  declare _addTickLblPos: (attrs?: Record<string, any>) => any;
  declare _insertTickLblPos: (child: any) => any;
  declare _newTickLblPos: () => any;
  declare _removeTickLblPos: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _newTxPr: () => any;
  declare _removeTxPr: () => void;
  declare crossAx: any;
  declare getOrAddCrossAx: () => any;
  declare _addCrossAx: (attrs?: Record<string, any>) => any;
  declare _insertCrossAx: (child: any) => any;
  declare _newCrossAx: () => any;
  declare _removeCrossAx: () => void;
  declare crosses: any;
  declare getOrAddCrosses: () => any;
  declare _addCrosses: (attrs?: Record<string, any>) => any;
  declare _insertCrosses: (child: any) => any;
  declare _newCrosses: () => any;
  declare _removeCrosses: () => void;
  declare crossesAt: any;
  declare getOrAddCrossesAt: () => any;
  declare _addCrossesAt: (attrs?: Record<string, any>) => any;
  declare _insertCrossesAt: (child: any) => any;
  declare _newCrossesAt: () => any;
  declare _removeCrossesAt: () => void;
  declare majorUnit: any;
  declare getOrAddMajorUnit: () => any;
  declare _addMajorUnit: (attrs?: Record<string, any>) => any;
  declare _insertMajorUnit: (child: any) => any;
  declare _newMajorUnit: () => any;
  declare _removeMajorUnit: () => void;
  declare minorUnit: any;
  declare getOrAddMinorUnit: () => any;
  declare _addMinorUnit: (attrs?: Record<string, any>) => any;
  declare _insertMinorUnit: (child: any) => any;
  declare _newMinorUnit: () => any;
  declare _removeMinorUnit: () => void;
  static {
    const _tagSeq = ["c:axId", "c:scaling", "c:delete", "c:axPos", "c:majorGridlines", "c:minorGridlines", "c:title", "c:numFmt", "c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", "c:crosses", "c:crossesAt", "c:crossBetween", "c:majorUnit", "c:minorUnit", "c:dispUnits", "c:extLst"] as const;
    defineChildren(this, {
      scaling: new OneAndOnlyOne("c:scaling"),
      delete_: new ZeroOrOne("c:delete", _tagSeq.slice(3)),
      majorGridlines: new ZeroOrOne("c:majorGridlines", _tagSeq.slice(5)),
      minorGridlines: new ZeroOrOne("c:minorGridlines", _tagSeq.slice(6)),
      title: new ZeroOrOne("c:title", _tagSeq.slice(7)),
      numFmt: new ZeroOrOne("c:numFmt", _tagSeq.slice(8)),
      majorTickMark: new ZeroOrOne("c:majorTickMark", _tagSeq.slice(9)),
      minorTickMark: new ZeroOrOne("c:minorTickMark", _tagSeq.slice(10)),
      tickLblPos: new ZeroOrOne("c:tickLblPos", _tagSeq.slice(11)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(12)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(13)),
      crossAx: new ZeroOrOne("c:crossAx", _tagSeq.slice(14)),
      crosses: new ZeroOrOne("c:crosses", _tagSeq.slice(15)),
      crossesAt: new ZeroOrOne("c:crossesAt", _tagSeq.slice(16)),
      majorUnit: new ZeroOrOne("c:majorUnit", _tagSeq.slice(18)),
      minorUnit: new ZeroOrOne("c:minorUnit", _tagSeq.slice(19)),
    });
  }
}
