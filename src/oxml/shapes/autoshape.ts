// lxml custom element classes for shape-related XML elements. (`pptx.oxml.shapes.autoshape`)
import { ValueError } from "../../errors.js";
import { defineChildren } from "../xmlchemy.js";
import { pyPercent } from "../../util.js";
import { MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER } from "../../enum/shapes.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { BaseShapeElement } from "./shared.js";
import { ST_Coordinate, ST_PositiveCoordinate, XsdBoolean, XsdString } from "../simpletypes.js";
import { CT_TextBody } from "../text.js";
import { BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrMore, ZeroOrOne } from "../xmlchemy.js";
import type { CT_ApplicationNonVisualDrawingProps, CT_NonVisualDrawingProps, CT_ShapeProperties } from "./shared.js";
import type { Length } from "../../util.js";

/** `a:pt` custom element class. */
export class CT_AdjPoint2D extends BaseOxmlElement {
  declare x: number;
  declare y: number;
  static {
    defineChildren(this, {
      x: new RequiredAttribute("x", ST_Coordinate),
      y: new RequiredAttribute("y", ST_Coordinate),
    });
  }
}

/** `a:custGeom` custom element class. */
export class CT_CustomGeometry2D extends BaseOxmlElement {
  declare getOrAddPathLst: (...args: any[]) => CT_Path2DList;
  declare pathLst: CT_Path2DList | null;
  declare _addPathLst: (attrs?: Record<string, any>) => CT_Path2DList;
  declare _insertPathLst: (child: CT_Path2DList) => CT_Path2DList;
  declare _newPathLst: () => CT_Path2DList;
  declare _removePathLst: () => void;
  static {
    const _tagSeq = ["a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst", "a:rect", "a:pathLst"] as const;
    defineChildren(this, {
      pathLst: new ZeroOrOne("a:pathLst", _tagSeq.slice(6)),
    });
  }
}

/**
 * `a:gd` custom element class.
 *
 * Defines a "guide", corresponding to a yellow diamond-shaped handle on an autoshape.
 */
export class CT_GeomGuide extends BaseOxmlElement {
  declare name: string;
  declare fmla: string;
  static {
    defineChildren(this, {
      name: new RequiredAttribute("name", XsdString),
      fmla: new RequiredAttribute("fmla", XsdString),
    });
  }
}

/** `a:avLst` custom element class. */
export class CT_GeomGuideList extends BaseOxmlElement {
  declare _addGd: (...args: any[]) => CT_GeomGuide;
  declare gdLst: CT_GeomGuide[];
  declare _insertGd: (child: any) => any;
  declare _newGd: () => any;
  static {
    defineChildren(this, {
      gd: new ZeroOrMore("a:gd"),
    });
  }
}

/** `p:cNvSpPr` custom element class. */
export class CT_NonVisualDrawingShapeProps extends BaseShapeElement {
  declare spLocks: any;
  declare getOrAddSpLocks: () => any;
  declare _addSpLocks: (attrs?: Record<string, any>) => any;
  declare _insertSpLocks: (child: any) => any;
  declare _newSpLocks: () => any;
  declare _removeSpLocks: () => void;
  declare txBox: boolean | null;
  static {
    defineChildren(this, {
      spLocks: new ZeroOrOne("a:spLocks"),
      txBox: new OptionalAttribute("txBox", XsdBoolean),
    });
  }
}

/** `a:path` custom element class. */
export class CT_Path2D extends BaseOxmlElement {
  declare _addClose: (...args: any[]) => CT_Path2DClose;
  declare _addLnTo: (...args: any[]) => CT_Path2DLineTo;
  declare _addMoveTo: (...args: any[]) => CT_Path2DMoveTo;
  declare closeLst: any[];
  declare _insertClose: (child: any) => any;
  declare _newClose: () => any;
  declare lnToLst: any[];
  declare _insertLnTo: (child: any) => any;
  declare _newLnTo: () => any;
  declare moveToLst: any[];
  declare _insertMoveTo: (child: any) => any;
  declare _newMoveTo: () => any;
  declare w: number | null;
  declare h: number | null;
  static {
    defineChildren(this, {
      close: new ZeroOrMore("a:close", []),
      lnTo: new ZeroOrMore("a:lnTo", []),
      moveTo: new ZeroOrMore("a:moveTo", []),
      w: new OptionalAttribute("w", ST_PositiveCoordinate),
      h: new OptionalAttribute("h", ST_PositiveCoordinate),
    });
  }

  /**
   * Return a newly created `a:close` element.
   *
   * The new `a:close` element is appended to this `a:path` element.
   */
  addClose(): CT_Path2DClose {
    return this._addClose();
  }

  /**
   * Return a newly created `a:lnTo` subtree with end point *(x, y)*.
   *
   * The new `a:lnTo` element is appended to this `a:path` element.
   */
  addLnTo(x: number, y: number): CT_Path2DLineTo {
    const lnTo = this._addLnTo();
    const pt = lnTo._addPt();
    [pt.x, pt.y] = [x, y];
    return lnTo;
  }

  /**
   * Return a newly created `a:moveTo` subtree with point `(x, y)`.
   *
   * The new `a:moveTo` element is appended to this `a:path` element.
   */
  addMoveTo(x: number, y: number) {
    const moveTo = this._addMoveTo();
    const pt = moveTo._addPt();
    [pt.x, pt.y] = [x, y];
    return moveTo;
  }
}

/** `a:close` custom element class. */
export class CT_Path2DClose extends BaseOxmlElement {
}

/** `a:lnTo` custom element class. */
export class CT_Path2DLineTo extends BaseOxmlElement {
  declare _addPt: (...args: any[]) => CT_AdjPoint2D;
  declare pt: any;
  declare getOrAddPt: () => any;
  declare _insertPt: (child: any) => any;
  declare _newPt: () => any;
  declare _removePt: () => void;
  static {
    defineChildren(this, {
      pt: new ZeroOrOne("a:pt", []),
    });
  }
}

/** `a:pathLst` custom element class. */
export class CT_Path2DList extends BaseOxmlElement {
  declare _addPath: (...args: any[]) => CT_Path2D;
  declare pathLst: any[];
  declare _insertPath: (child: any) => any;
  declare _newPath: () => any;
  static {
    defineChildren(this, {
      path: new ZeroOrMore("a:path", []),
    });
  }

  /** Return a newly created `a:path` child element. */
  addPath(w: number, h: number) {
    const path = this._addPath();
    [path.w, path.h] = [w, h];
    return path;
  }
}

/** `a:moveTo` custom element class. */
export class CT_Path2DMoveTo extends BaseOxmlElement {
  declare _addPt: (...args: any[]) => CT_AdjPoint2D;
  declare pt: any;
  declare getOrAddPt: () => any;
  declare _insertPt: (child: any) => any;
  declare _newPt: () => any;
  declare _removePt: () => void;
  static {
    defineChildren(this, {
      pt: new ZeroOrOne("a:pt", []),
    });
  }
}

/** `a:prstGeom` custom element class. */
export class CT_PresetGeometry2D extends BaseOxmlElement {
  declare _addAvLst: (...args: any[]) => CT_GeomGuideList;
  declare _removeAvLst: (...args: any[]) => null;
  declare avLst: CT_GeomGuideList | null;
  declare getOrAddAvLst: () => CT_GeomGuideList;
  declare _insertAvLst: (child: CT_GeomGuideList) => CT_GeomGuideList;
  declare _newAvLst: () => CT_GeomGuideList;
  declare prst: MSO_AUTO_SHAPE_TYPE;
  static {
    defineChildren(this, {
      avLst: new ZeroOrOne("a:avLst"),
      prst: new RequiredAttribute("prst", MSO_AUTO_SHAPE_TYPE),
    });
  }

  /** Sequence of `a:gd` element children of `a:avLst`. Empty if none are present. */
  get gdLst(): CT_GeomGuide[] {
    const avLst = this.avLst;
    if (avLst === null) {
      return [];
    }
    return avLst.gdLst;
  }

  /** Replace any `a:gd` element children of `a:avLst` with ones forme from `guides`. */
  rewriteGuides(guides: [string, number][]) {
    this._removeAvLst();
    const avLst = this._addAvLst();
    for (const [name, val] of guides) {
      const gd = avLst._addGd();
      gd.name = name;
      gd.fmla = `val ${val}`;
    }
  }
}

/** `p:sp` custom element class. */
export class CT_Shape extends BaseShapeElement {
  declare getOrAddTxBody: (...args: any[]) => CT_TextBody;
  declare nvSpPr: CT_ShapeNonVisual;
  declare spPr: CT_ShapeProperties;
  declare _addTxBody: (attrs?: Record<string, any>) => CT_TextBody;
  declare _insertTxBody: (child: CT_TextBody) => CT_TextBody;
  declare _removeTxBody: () => void;
  static {
    defineChildren(this, {
      nvSpPr: new OneAndOnlyOne("p:nvSpPr"),
      spPr: new OneAndOnlyOne("p:spPr"),
      txBody: new ZeroOrOne("p:txBody", ["p:extLst"]),
    });
  }

  addPath(w: number, h: number): CT_Path2D {
    const custGeom = this.spPr.custGeom;
    if (custGeom === null) {
      throw new ValueError("shape must be freeform");
    }
    const pathLst = custGeom.getOrAddPathLst();
    return pathLst.addPath(w, h);
  }

  /** Return the `a:ln` grandchild element, newly added if not present. */
  getOrAddLn() {
    return this.spPr.getOrAddLn();
  }

  /**
   * True if this shape has custom geometry, i.e. is a freeform shape.
   *
   * A shape has custom geometry if it has a `p:spPr/a:custGeom`
   * descendant (instead of `p:spPr/a:prstGeom`).
   */
  get hasCustomGeometry() {
    return this.spPr.custGeom !== null;
  }

  /**
   * True if this shape is an auto shape.
   *
   * A shape is an auto shape if it has a `a:prstGeom` element and does not have a txBox="1"
   * attribute on cNvSpPr.
   */
  get isAutoshape() {
    const prstGeom = this.prstGeom;
    if (prstGeom === null) {
      return false;
    }
    return this.nvSpPr.cNvSpPr.txBox !== true;
  }

  /**
   * True if this shape is a text box.
   *
   * A shape is a text box if it has a `txBox` attribute on cNvSpPr that resolves to |True|.
   * The default when the txBox attribute is missing is |False|.
   */
  get isTextbox() {
    return this.nvSpPr.cNvSpPr.txBox === true;
  }

  /** `a:ln` grand-child element or |None| if not present. */
  get ln() {
    return this.spPr.ln;
  }

  /** Return a new `p:sp` element tree configured as a base auto shape. */
  static newAutoshapeSp(id_: number, name: string, prst: string, left: number, top: number, width: number, height: number): CT_Shape {
    const xml = pyPercent(`<p:sp ${nsdecls("a", "p")}>
  <p:nvSpPr>
    <p:cNvPr id="${"%d"}" name="${"%s"}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${"%d"}" y="${"%d"}"/>
      <a:ext cx="${"%d"}" cy="${"%d"}"/>
    </a:xfrm>
    <a:prstGeom prst="${"%s"}">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1">
      <a:schemeClr val="accent1"/>
    </a:lnRef>
    <a:fillRef idx="3">
      <a:schemeClr val="accent1"/>
    </a:fillRef>
    <a:effectRef idx="2">
      <a:schemeClr val="accent1"/>
    </a:effectRef>
    <a:fontRef idx="minor">
      <a:schemeClr val="lt1"/>
    </a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
    </a:p>
  </p:txBody>
</p:sp>`, [id_, name, left, top, width, height, prst]);
    return parseXml(xml);
  }

  /**
   * Return new `p:sp` element tree configured as freeform shape.
   *
   * The returned shape has a `a:custGeom` subtree but no paths in its
   * path list.
   */
  static newFreeformSp(shapeId: number, name: string, x: number, y: number, cx: number, cy: number) {
    const xml = pyPercent(`<p:sp ${nsdecls("a", "p")}>
  <p:nvSpPr>
    <p:cNvPr id="${"%d"}" name="${"%s"}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${"%d"}" y="${"%d"}"/>
      <a:ext cx="${"%d"}" cy="${"%d"}"/>
    </a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="l" t="t" r="r" b="b"/>
      <a:pathLst/>
    </a:custGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1">
      <a:schemeClr val="accent1"/>
    </a:lnRef>
    <a:fillRef idx="3">
      <a:schemeClr val="accent1"/>
    </a:fillRef>
    <a:effectRef idx="2">
      <a:schemeClr val="accent1"/>
    </a:effectRef>
    <a:fontRef idx="minor">
      <a:schemeClr val="lt1"/>
    </a:fontRef>
  </p:style>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
    </a:p>
  </p:txBody>
</p:sp>`, [shapeId, name, x, y, cx, cy]);
    return parseXml(xml);
  }

  /** Return a new `p:sp` element tree configured as a placeholder shape. */
  static newPlaceholderSp(id_: number, name: string, phType: PP_PLACEHOLDER, orient: string, sz: any, idx: any): CT_Shape {
    const sp = parseXml(`<p:sp ${nsdecls("a", "p")}>
  <p:nvSpPr>
    <p:cNvPr id="${id_}" name="${name}"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1"/>
    </p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
</p:sp>`);
    const ph = sp.nvSpPr.nvPr.getOrAddPh();
    ph.type = phType;
    ph.idx = idx;
    ph.orient = orient;
    ph.sz = sz;
    const placeholderTypesThatHaveATextFrame = [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT];
    if (placeholderTypesThatHaveATextFrame.includes(phType)) {
      sp.append(CT_TextBody.new());
    }
    return sp;
  }

  /** Return a new `p:sp` element tree configured as a base textbox shape. */
  static newTextboxSp(id_: any, name: any, left: any, top: any, width: any, height: any) {
    const tmpl = CT_Shape._textboxSpTmpl();
    const xml = pyPercent(tmpl, [id_, name, left, top, width, height]);
    const sp = parseXml(xml);
    return sp;
  }

  /** Value of `prst` attribute of `a:prstGeom` element or |None| if not present. */
  get prst() {
    const prstGeom = this.prstGeom;
    if (prstGeom === null) {
      return null;
    }
    return prstGeom.prst;
  }

  /**
   * Reference to `a:prstGeom` child element.
   *
   * |None| if this shape doesn't have one, for example, if it's a placeholder shape.
   */
  get prstGeom(): CT_PresetGeometry2D | null {
    return this.spPr.prstGeom;
  }

  _newTxBody() {
    return CT_TextBody.newPTxBody();
  }

  static _textboxSpTmpl() {
    return `<p:sp ${nsdecls("a", "p")}>
  <p:nvSpPr>
    <p:cNvPr id="${"%d"}" name="${"%s"}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${"%d"}" y="${"%d"}"/>
      <a:ext cx="${"%d"}" cy="${"%d"}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="none">
      <a:spAutoFit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>`;
  }
}

/** `p:nvSpPr` custom element class. */
export class CT_ShapeNonVisual extends BaseShapeElement {
  declare cNvPr: CT_NonVisualDrawingProps;
  declare cNvSpPr: CT_NonVisualDrawingShapeProps;
  declare nvPr: CT_ApplicationNonVisualDrawingProps;
  static {
    defineChildren(this, {
      cNvPr: new OneAndOnlyOne("p:cNvPr"),
      cNvSpPr: new OneAndOnlyOne("p:cNvSpPr"),
      nvPr: new OneAndOnlyOne("p:nvPr"),
    });
  }
}
