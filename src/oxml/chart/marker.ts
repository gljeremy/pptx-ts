// Series-related oxml objects. (`pptx.oxml.chart.marker`)
import { defineChildren } from "../xmlchemy.js";
import { XL_MARKER_STYLE } from "../../enum/chart.js";
import { ST_MarkerSize } from "../simpletypes.js";
import { BaseOxmlElement, RequiredAttribute, ZeroOrOne } from "../xmlchemy.js";

/**
 * `c:marker` custom element class, containing visual properties for a data
 * point marker on line-type charts.
 */
export class CT_Marker extends BaseOxmlElement {
  declare symbol: any;
  declare getOrAddSymbol: () => any;
  declare _addSymbol: (attrs?: Record<string, any>) => any;
  declare _insertSymbol: (child: any) => any;
  declare _newSymbol: () => any;
  declare _removeSymbol: () => void;
  declare size: any;
  declare getOrAddSize: () => any;
  declare _addSize: (attrs?: Record<string, any>) => any;
  declare _insertSize: (child: any) => any;
  declare _newSize: () => any;
  declare _removeSize: () => void;
  declare spPr: any;
  declare getOrAddSpPr: () => any;
  declare _addSpPr: (attrs?: Record<string, any>) => any;
  declare _insertSpPr: (child: any) => any;
  declare _newSpPr: () => any;
  declare _removeSpPr: () => void;
  static {
    const _tagSeq = ["c:symbol", "c:size", "c:spPr", "c:extLst"] as const;
    defineChildren(this, {
      symbol: new ZeroOrOne("c:symbol", _tagSeq.slice(1)),
      size: new ZeroOrOne("c:size", _tagSeq.slice(2)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(3)),
    });
  }

  /**
   * Return the value of `./c:size/@val`, specifying the size of this
   * marker in points. Returns |None| if no `c:size` element is present or
   * its val attribute is not present.
   */
  get sizeVal() {
    const size = this.size;
    if (size === null) {
      return null;
    }
    return size.val;
  }

  /**
   * Return the value of `./c:symbol/@val`, specifying the shape of this
   * marker. Returns |None| if no `c:symbol` element is present.
   */
  get symbolVal() {
    const symbol = this.symbol;
    if (symbol === null) {
      return null;
    }
    return symbol.val;
  }
}

/**
 * `c:size` custom element class, specifying the size (in points) of a data
 * point marker for a line, XY, or radar chart.
 */
export class CT_MarkerSize extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_MarkerSize),
    });
  }
}

/**
 * `c:symbol` custom element class, specifying the shape of a data point
 * marker for a line, XY, or radar chart.
 */
export class CT_MarkerStyle extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", XL_MARKER_STYLE),
    });
  }
}
