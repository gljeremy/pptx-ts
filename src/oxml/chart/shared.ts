// Shared oxml objects for charts. (`pptx.oxml.chart.shared`)
import { defineChildren } from "../xmlchemy.js";
import { pyBool } from "../../util.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { ST_LayoutMode, XsdBoolean, XsdDouble, XsdString, XsdUnsignedInt } from "../simpletypes.js";
import { BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrOne } from "../xmlchemy.js";

/** Common complex type used for elements having a True/False value. */
export class CT_Boolean extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", XsdBoolean, true),
    });
  }
}

/**
 * Always spells out the `val` attribute, e.g. `val=1`.
 *
 * At least one boolean element is improperly interpreted by one or more
 * versions of PowerPoint. The `c:overlay` element is interpreted as |False|
 * when no `val` attribute is present, contrary to the behavior described in
 * the schema. A remedy for this is to interpret a missing `val` attribute
 * as |True| (consistent with the spec), but always write the attribute
 * whenever there is occasion for changing the element.
 */
export class CT_Boolean_Explicit extends BaseOxmlElement {
  declare _val: any;
  static {
    defineChildren(this, {
      _val: new OptionalAttribute("val", XsdBoolean, true),
    });
  }

  get val() {
    return this._val;
  }

  set val(value: any) {
    const valStr = (pyBool(value) === true ? "1" : "0");
    this.set("val", valStr);
  }
}

/** Used for floating point values. */
export class CT_Double extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", XsdDouble),
    });
  }
}

/** ``<c:layout>`` custom element class */
export class CT_Layout extends BaseOxmlElement {
  declare manualLayout: any;
  declare getOrAddManualLayout: () => any;
  declare _addManualLayout: (attrs?: Record<string, any>) => any;
  declare _insertManualLayout: (child: any) => any;
  declare _newManualLayout: () => any;
  declare _removeManualLayout: () => void;
  static {
    defineChildren(this, {
      manualLayout: new ZeroOrOne("c:manualLayout", ["c:extLst"]),
    });
  }

  /**
   * The float value in ./c:manualLayout/c:x when
   * c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
   * expression finds no match.
   */
  get horzOffset() {
    const manualLayout = this.manualLayout;
    if (manualLayout === null) {
      return 0.0;
    }
    return manualLayout.horzOffset;
  }

  set horzOffset(offset: any) {
    if (offset === 0.0) {
      this._removeManualLayout();
      return;
    }
    const manualLayout = this.getOrAddManualLayout();
    manualLayout.horzOffset = offset;
  }
}

/**
 * Used for ``<c:xMode>``, ``<c:yMode>``, ``<c:wMode>``, and ``<c:hMode>``
 * child elements of CT_ManualLayout.
 */
export class CT_LayoutMode extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", ST_LayoutMode, ST_LayoutMode.FACTOR),
    });
  }
}

/** ``<c:manualLayout>`` custom element class */
export class CT_ManualLayout extends BaseOxmlElement {
  declare xMode: any;
  declare getOrAddXMode: () => any;
  declare _addXMode: (attrs?: Record<string, any>) => any;
  declare _insertXMode: (child: any) => any;
  declare _newXMode: () => any;
  declare _removeXMode: () => void;
  declare x: any;
  declare getOrAddX: () => any;
  declare _addX: (attrs?: Record<string, any>) => any;
  declare _insertX: (child: any) => any;
  declare _newX: () => any;
  declare _removeX: () => void;
  static {
    const _tagSeq = ["c:layoutTarget", "c:xMode", "c:yMode", "c:wMode", "c:hMode", "c:x", "c:y", "c:w", "c:h", "c:extLst"] as const;
    defineChildren(this, {
      xMode: new ZeroOrOne("c:xMode", _tagSeq.slice(2)),
      x: new ZeroOrOne("c:x", _tagSeq.slice(6)),
    });
  }

  /**
   * The float value in ./c:x@val when ./c:xMode@val == "factor". 0.0 when
   * ./c:x is not present or ./c:xMode@val != "factor".
   */
  get horzOffset() {
    const [x, xMode] = [this.x, this.xMode];
    if ((x === null || xMode === null || xMode.val !== ST_LayoutMode.FACTOR)) {
      return 0.0;
    }
    return x.val;
  }

  set horzOffset(offset: any) {
    this.getOrAddXMode().val = ST_LayoutMode.FACTOR;
    this.getOrAddX().val = offset;
  }
}

/**
 * ``<c:numFmt>`` element specifying the formatting for number labels on a
 * tick mark or data point.
 */
export class CT_NumFmt extends BaseOxmlElement {
  declare formatCode: any;
  declare sourceLinked: any;
  static {
    defineChildren(this, {
      formatCode: new RequiredAttribute("formatCode", XsdString),
      sourceLinked: new OptionalAttribute("sourceLinked", XsdBoolean),
    });
  }
}

/** `c:title` custom element class. */
export class CT_Title extends BaseOxmlElement {
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
  static {
    const _tagSeq = ["c:tx", "c:layout", "c:overlay", "c:spPr", "c:txPr", "c:extLst"] as const;
    defineChildren(this, {
      tx: new ZeroOrOne("c:tx", _tagSeq.slice(1)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(4)),
    });
  }

  /**
   * Return `c:tx/c:rich`, newly created if not present.
   *
   * Return the `c:rich` grandchild at `c:tx/c:rich`. Both the `c:tx` and
   * `c:rich` elements are created if not already present. Any
   * `c:tx/c:strRef` element is removed. (Such an element would contain
   * a cell reference for the axis title text in the chart's Excel
   * worksheet.)
   */
  getOrAddTxRich() {
    const tx = this.getOrAddTx();
    tx._removeStrRef();
    return tx.getOrAddRich();
  }

  /** Return `c:tx/c:rich` or |None| if not present. */
  get txRich() {
    const richs = this.xpath("c:tx/c:rich");
    if (richs.length === 0) {
      return null;
    }
    return richs[0];
  }

  /** Return "loose" `c:title` element containing default children. */
  static newTitle() {
    return parseXml(`<c:title ${nsdecls("c")}>  <c:layout/>  <c:overlay val="0"/></c:title>`);
  }
}

/**
 * ``<c:tx>`` element containing the text for a label on a data point or
 * other chart item.
 */
export class CT_Tx extends BaseOxmlElement {
  declare strRef: any;
  declare getOrAddStrRef: () => any;
  declare _addStrRef: (attrs?: Record<string, any>) => any;
  declare _insertStrRef: (child: any) => any;
  declare _newStrRef: () => any;
  declare _removeStrRef: () => void;
  declare rich: any;
  declare getOrAddRich: () => any;
  declare _addRich: (attrs?: Record<string, any>) => any;
  declare _insertRich: (child: any) => any;
  declare _removeRich: () => void;
  static {
    defineChildren(this, {
      strRef: new ZeroOrOne("c:strRef"),
      rich: new ZeroOrOne("c:rich"),
    });
  }

  _newRich() {
    return parseXml(`<c:rich ${nsdecls("c", "a")}>  <a:bodyPr/>  <a:lstStyle/>  <a:p>    <a:pPr>      <a:defRPr/>    </a:pPr>  </a:p></c:rich>`);
  }
}

/** ``<c:idx>`` element and others. */
export class CT_UnsignedInt extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", XsdUnsignedInt),
    });
  }
}
