// lxml custom element classes for legend-related XML elements. (`pptx.oxml.chart.legend`)
import { defineChildren } from "../xmlchemy.js";
import { XL_LEGEND_POSITION } from "../../enum/chart.js";
import { CT_TextBody } from "../text.js";
import { BaseOxmlElement, OptionalAttribute, ZeroOrOne } from "../xmlchemy.js";

/** ``<c:legend>`` custom element class */
export class CT_Legend extends BaseOxmlElement {
  declare legendPos: any;
  declare getOrAddLegendPos: () => any;
  declare _addLegendPos: (attrs?: Record<string, any>) => any;
  declare _insertLegendPos: (child: any) => any;
  declare _newLegendPos: () => any;
  declare _removeLegendPos: () => void;
  declare layout: any;
  declare getOrAddLayout: () => any;
  declare _addLayout: (attrs?: Record<string, any>) => any;
  declare _insertLayout: (child: any) => any;
  declare _newLayout: () => any;
  declare _removeLayout: () => void;
  declare overlay: any;
  declare getOrAddOverlay: () => any;
  declare _addOverlay: (attrs?: Record<string, any>) => any;
  declare _insertOverlay: (child: any) => any;
  declare _newOverlay: () => any;
  declare _removeOverlay: () => void;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _removeTxPr: () => void;
  static {
    const _tagSeq = ["c:legendPos", "c:legendEntry", "c:layout", "c:overlay", "c:spPr", "c:txPr", "c:extLst"] as const;
    defineChildren(this, {
      legendPos: new ZeroOrOne("c:legendPos", _tagSeq.slice(1)),
      layout: new ZeroOrOne("c:layout", _tagSeq.slice(3)),
      overlay: new ZeroOrOne("c:overlay", _tagSeq.slice(4)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(6)),
    });
  }

  /**
   * `./c:txPr/a:p/a:pPr/a:defRPr` great-great-grandchild element, added
   * with its ancestors if not present.
   */
  get defRPr() {
    const txPr = this.getOrAddTxPr();
    const defRPr = txPr.defRPr;
    return defRPr;
  }

  /**
   * The float value in ./c:layout/c:manualLayout/c:x when
   * ./c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that
   * XPath expression has no match.
   */
  get horzOffset() {
    const layout = this.layout;
    if (layout === null) {
      return 0.0;
    }
    return layout.horzOffset;
  }

  set horzOffset(offset: any) {
    const layout = this.getOrAddLayout();
    layout.horzOffset = offset;
  }

  _newTxPr() {
    return CT_TextBody.newTxPr();
  }
}

/**
 * ``<c:legendPos>`` element specifying position of legend with respect to
 * chart as a member of ST_LegendPos.
 */
export class CT_LegendPos extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", XL_LEGEND_POSITION, XL_LEGEND_POSITION.RIGHT),
    });
  }
}
