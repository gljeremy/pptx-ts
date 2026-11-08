// lxml custom element classes for XML elements related to the Connector shape. (`pptx.oxml.shapes.connector`)
import { defineChildren } from "../xmlchemy.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { BaseShapeElement } from "./shared.js";
import { ST_DrawingElementId, XsdUnsignedInt } from "../simpletypes.js";
import { BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne } from "../xmlchemy.js";
import type { CT_ShapeProperties } from "./shared.js";

/**
 * A `a:stCxn` or `a:endCxn` element.
 *
 * Specifies a connection between an end-point of a connector and a shape connection point.
 */
export class CT_Connection extends BaseShapeElement {
  declare id: any;
  declare idx: any;
  static {
    defineChildren(this, {
      id: new RequiredAttribute("id", ST_DrawingElementId),
      idx: new RequiredAttribute("idx", XsdUnsignedInt),
    });
  }
}

/** A line/connector shape `p:cxnSp` element */
export class CT_Connector extends BaseShapeElement {
  declare nvCxnSpPr: any;
  declare spPr: CT_ShapeProperties;
  static {
    const _tagSeq = ["p:nvCxnSpPr", "p:spPr", "p:style", "p:extLst"] as const;
    defineChildren(this, {
      nvCxnSpPr: new OneAndOnlyOne("p:nvCxnSpPr"),
      spPr: new OneAndOnlyOne("p:spPr"),
    });
  }

  /** Return a new `p:cxnSp` element tree configured as a base connector. */
  static newCxnSp(id_: number, name: string, prst: string, x: number, y: number, cx: number, cy: number, flipH: boolean, flipV: boolean): CT_Connector {
    const flip = ((flipH ? " flipH=\"1\"" : "") + (flipV ? " flipV=\"1\"" : ""));
    return parseXml(`<p:cxnSp ${nsdecls("a", "p")}>
  <p:nvCxnSpPr>
    <p:cNvPr id="${id_}" name="${name}"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm${flip}>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${cx}" cy="${cy}"/>
    </a:xfrm>
    <a:prstGeom prst="${prst}">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2">
      <a:schemeClr val="accent1"/>
    </a:lnRef>
    <a:fillRef idx="0">
      <a:schemeClr val="accent1"/>
    </a:fillRef>
    <a:effectRef idx="1">
      <a:schemeClr val="accent1"/>
    </a:effectRef>
    <a:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </a:fontRef>
  </p:style>
</p:cxnSp>`);
  }
}

/**
 * `p:nvCxnSpPr` element, container for the non-visual properties of
 * a connector, such as name, id, etc.
 */
export class CT_ConnectorNonVisual extends BaseOxmlElement {
  declare cNvPr: any;
  declare cNvCxnSpPr: any;
  declare nvPr: any;
  static {
    defineChildren(this, {
      cNvPr: new OneAndOnlyOne("p:cNvPr"),
      cNvCxnSpPr: new OneAndOnlyOne("p:cNvCxnSpPr"),
      nvPr: new OneAndOnlyOne("p:nvPr"),
    });
  }
}

/**
 * `p:cNvCxnSpPr` element, container for the non-visual properties specific
 * to a connector shape, such as connections and connector locking.
 */
export class CT_NonVisualConnectorProperties extends BaseOxmlElement {
  declare stCxn: any;
  declare getOrAddStCxn: () => any;
  declare _addStCxn: (attrs?: Record<string, any>) => any;
  declare _insertStCxn: (child: any) => any;
  declare _newStCxn: () => any;
  declare _removeStCxn: () => void;
  declare endCxn: any;
  declare getOrAddEndCxn: () => any;
  declare _addEndCxn: (attrs?: Record<string, any>) => any;
  declare _insertEndCxn: (child: any) => any;
  declare _newEndCxn: () => any;
  declare _removeEndCxn: () => void;
  static {
    const _tagSeq = ["a:cxnSpLocks", "a:stCxn", "a:endCxn", "a:extLst"] as const;
    defineChildren(this, {
      stCxn: new ZeroOrOne("a:stCxn", _tagSeq.slice(2)),
      endCxn: new ZeroOrOne("a:endCxn", _tagSeq.slice(3)),
    });
  }
}
