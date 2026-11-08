// Chart data-label related oxml objects. (`pptx.oxml.chart.datalabel`)
import { defineChildren } from "../xmlchemy.js";
import { XL_DATA_LABEL_POSITION } from "../../enum/chart.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { CT_TextBody } from "../text.js";
import { BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne } from "../xmlchemy.js";

/**
 * ``<c:dLbl>`` element specifying the properties of the data label for an
 * individual data point.
 */
export class CT_DLbl extends BaseOxmlElement {
  declare idx: any;
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
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _removeTxPr: () => void;
  declare dLblPos: any;
  declare getOrAddDLblPos: () => any;
  declare _addDLblPos: (attrs?: Record<string, any>) => any;
  declare _insertDLblPos: (child: any) => any;
  declare _newDLblPos: () => any;
  declare _removeDLblPos: () => void;
  static {
    const _tagSeq = ["c:idx", "c:layout", "c:tx", "c:numFmt", "c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize", "c:separator", "c:extLst"] as const;
    defineChildren(this, {
      idx: new OneAndOnlyOne("c:idx"),
      tx: new ZeroOrOne("c:tx", _tagSeq.slice(3)),
      spPr: new ZeroOrOne("c:spPr", _tagSeq.slice(5)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(6)),
      dLblPos: new ZeroOrOne("c:dLblPos", _tagSeq.slice(7)),
    });
  }

  /**
   * Return the `c:rich` descendant representing the text frame of the
   * data label, newly created if not present. Any existing `c:strRef`
   * element is removed along with its contents.
   */
  getOrAddRich() {
    const tx = this.getOrAddTx();
    tx._removeStrRef();
    return tx.getOrAddRich();
  }

  /** Return the `c:tx[c:rich]` subtree, newly created if not present. */
  getOrAddTxRich() {
    const tx = this.getOrAddTx();
    tx._removeStrRef();
    tx.getOrAddRich();
    return tx;
  }

  /**
   * The integer value of the `val` attribute on the required `c:idx`
   * child.
   */
  get idxVal() {
    return this.idx.val;
  }

  /**
   * Return a newly created "loose" `c:dLbl` element.
   *
   * The `c:dLbl` element contains the same (fairly extensive) default
   * subtree added by PowerPoint when an individual data label is
   * customized in the UI. Note that the idx value must be set by the
   * client. Failure to set the idx value will likely result in any
   * changes not being visible and may result in a repair error on open.
   */
  static newDLbl() {
    return parseXml(`<c:dLbl ${nsdecls("c", "a")}>
  <c:idx val="666"/>
  <c:spPr/>
  <c:txPr>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr/>
      </a:pPr>
    </a:p>
  </c:txPr>
  <c:showLegendKey val="0"/>
  <c:showVal val="1"/>
  <c:showCatName val="0"/>
  <c:showSerName val="0"/>
  <c:showPercent val="0"/>
  <c:showBubbleSize val="0"/>
</c:dLbl>`);
  }

  /** Remove any `c:tx[c:rich]` child, or do nothing if not present. */
  removeTxRich() {
    const matches = this.xpath("c:tx[c:rich]");
    if (matches.length === 0) {
      return;
    }
    const tx = matches[0];
    this.remove(tx);
  }

  _newTxPr() {
    return CT_TextBody.newTxPr();
  }
}

/**
 * ``<c:dLblPos>`` element specifying the positioning of a data label with
 * respect to its data point.
 */
export class CT_DLblPos extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", XL_DATA_LABEL_POSITION),
    });
  }
}

/** `c:dLbls` element specifying properties for a set of data labels. */
export class CT_DLbls extends BaseOxmlElement {
  declare dLblLst: any[];
  declare _addDLbl: (attrs?: Record<string, any>) => any;
  declare _insertDLbl: (child: any) => any;
  declare numFmt: any;
  declare getOrAddNumFmt: () => any;
  declare _addNumFmt: (attrs?: Record<string, any>) => any;
  declare _insertNumFmt: (child: any) => any;
  declare _newNumFmt: () => any;
  declare _removeNumFmt: () => void;
  declare txPr: any;
  declare getOrAddTxPr: () => any;
  declare _addTxPr: (attrs?: Record<string, any>) => any;
  declare _insertTxPr: (child: any) => any;
  declare _removeTxPr: () => void;
  declare dLblPos: any;
  declare getOrAddDLblPos: () => any;
  declare _addDLblPos: (attrs?: Record<string, any>) => any;
  declare _insertDLblPos: (child: any) => any;
  declare _newDLblPos: () => any;
  declare _removeDLblPos: () => void;
  declare showLegendKey: any;
  declare getOrAddShowLegendKey: () => any;
  declare _addShowLegendKey: (attrs?: Record<string, any>) => any;
  declare _insertShowLegendKey: (child: any) => any;
  declare _removeShowLegendKey: () => void;
  declare showVal: any;
  declare getOrAddShowVal: () => any;
  declare _addShowVal: (attrs?: Record<string, any>) => any;
  declare _insertShowVal: (child: any) => any;
  declare _removeShowVal: () => void;
  declare showCatName: any;
  declare getOrAddShowCatName: () => any;
  declare _addShowCatName: (attrs?: Record<string, any>) => any;
  declare _insertShowCatName: (child: any) => any;
  declare _removeShowCatName: () => void;
  declare showSerName: any;
  declare getOrAddShowSerName: () => any;
  declare _addShowSerName: (attrs?: Record<string, any>) => any;
  declare _insertShowSerName: (child: any) => any;
  declare _removeShowSerName: () => void;
  declare showPercent: any;
  declare getOrAddShowPercent: () => any;
  declare _addShowPercent: (attrs?: Record<string, any>) => any;
  declare _insertShowPercent: (child: any) => any;
  declare _removeShowPercent: () => void;
  static {
    const _tagSeq = ["c:dLbl", "c:numFmt", "c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize", "c:separator", "c:showLeaderLines", "c:leaderLines", "c:extLst"] as const;
    defineChildren(this, {
      dLbl: new ZeroOrMore("c:dLbl", _tagSeq.slice(1)),
      numFmt: new ZeroOrOne("c:numFmt", _tagSeq.slice(2)),
      txPr: new ZeroOrOne("c:txPr", _tagSeq.slice(4)),
      dLblPos: new ZeroOrOne("c:dLblPos", _tagSeq.slice(5)),
      showLegendKey: new ZeroOrOne("c:showLegendKey", _tagSeq.slice(6)),
      showVal: new ZeroOrOne("c:showVal", _tagSeq.slice(7)),
      showCatName: new ZeroOrOne("c:showCatName", _tagSeq.slice(8)),
      showSerName: new ZeroOrOne("c:showSerName", _tagSeq.slice(9)),
      showPercent: new ZeroOrOne("c:showPercent", _tagSeq.slice(10)),
    });
  }

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
   * Return the `c:dLbl` child representing the label for the data point
   * at index *idx*.
   */
  getDLblForPoint(idx: any) {
    const matches = this.xpath(`c:dLbl[c:idx[@val="${idx}"]]`);
    if (matches.length) {
      return matches[0];
    }
    return null;
  }

  /**
   * Return the `c:dLbl` element representing the label of the point at
   * index *idx*.
   */
  getOrAddDLblForPoint(idx: any) {
    const matches = this.xpath(`c:dLbl[c:idx[@val="${idx}"]]`);
    if (matches.length) {
      return matches[0];
    }
    return this._insertDLblInSequence(idx);
  }

  /** Return a newly created "loose" `c:dLbls` element. */
  static newDLbls() {
    return parseXml(`<c:dLbls ${nsdecls("c")}>
  <c:showLegendKey val="0"/>
  <c:showVal val="0"/>
  <c:showCatName val="0"/>
  <c:showSerName val="0"/>
  <c:showPercent val="0"/>
  <c:showBubbleSize val="0"/>
  <c:showLeaderLines val="1"/>
</c:dLbls>`);
  }

  /**
   * Return a newly created `c:dLbl` element having `c:idx` child of *idx*
   * and inserted in numeric sequence among the `c:dLbl` children of this
   * element.
   */
  _insertDLblInSequence(idx: any) {
    const newDLbl = this._newDLbl();
    newDLbl.idx.val = idx;
    // python's loop variable outlives the loop: the last `c:dLbl`, if any
    let dLbl: any = null;
    for (dLbl of this.dLblLst) {
      if (dLbl.idxVal > idx) {
        dLbl.addprevious(newDLbl);
        return newDLbl;
      }
    }
    if (dLbl !== null) {
      dLbl.addnext(newDLbl);
    } else {
      this.insert(0, newDLbl);
    }
    return newDLbl;
  }

  _newDLbl() {
    return CT_DLbl.newDLbl();
  }

  /**
   * Return a new `c:showCatName` with value initialized.
   *
   * This method is called by the metaclass-generated code whenever a new
   * `c:showCatName` element is required. In this case, it defaults to
   * `val=true`, which is not what we need so we override to make val
   * explicitly False.
   */
  _newShowCatName() {
    return parseXml(`<c:showCatName ${nsdecls("c")} val="0"/>`);
  }

  _newShowLegendKey() {
    return parseXml(`<c:showLegendKey ${nsdecls("c")} val="0"/>`);
  }

  _newShowPercent() {
    return parseXml(`<c:showPercent ${nsdecls("c")} val="0"/>`);
  }

  _newShowSerName() {
    return parseXml(`<c:showSerName ${nsdecls("c")} val="0"/>`);
  }

  _newShowVal() {
    return parseXml(`<c:showVal ${nsdecls("c")} val="0"/>`);
  }

  _newTxPr() {
    return CT_TextBody.newTxPr();
  }
}
