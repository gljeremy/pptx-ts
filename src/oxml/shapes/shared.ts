// Common shape-related oxml objects. (`pptx.oxml.shapes.shared`)
import { defineChildren } from "../xmlchemy.js";
import { pyBool } from "../../util.js";
import { ValueError } from "../../errors.js";

import { CT_GradientFillProperties } from "../dml/fill.js";
import { PP_PLACEHOLDER } from "../../enum/shapes.js";
import { qn } from "../ns.js";
import { ST_Angle, ST_Coordinate, ST_Direction, ST_DrawingElementId, ST_LineWidth, ST_PlaceholderSize, ST_PositiveCoordinate, XsdBoolean, XsdString, XsdUnsignedInt } from "../simpletypes.js";
import { BaseOxmlElement, Choice, OptionalAttribute, OxmlElement, RequiredAttribute, ZeroOrOne, ZeroOrOneChoice } from "../xmlchemy.js";
import { Emu } from "../../util.js";
import type { CT_Hyperlink } from "../action.js";
import type { CT_CustomGeometry2D, CT_PresetGeometry2D } from "./autoshape.js";
import type { Length } from "../../util.js";

/** Provides common behavior for shape element classes like CT_Shape, CT_Picture, etc. */
export class BaseShapeElement extends BaseOxmlElement {
  declare spPr: CT_ShapeProperties;

  get cx(): number | null {
    return this._getXfrmAttr("cx");
  }

  set cx(value: any) {
    this._setXfrmAttr("cx", value);
  }

  get cy(): number | null {
    return this._getXfrmAttr("cy");
  }

  set cy(value: any) {
    this._setXfrmAttr("cy", value);
  }

  get flipH() {
    return pyBool(this._getXfrmAttr("flipH"));
  }

  set flipH(value: any) {
    this._setXfrmAttr("flipH", value);
  }

  get flipV() {
    return pyBool(this._getXfrmAttr("flipV"));
  }

  set flipV(value: any) {
    this._setXfrmAttr("flipV", value);
  }

  /**
   * Return the `a:xfrm` grandchild element, newly-added if not present.
   *
   * This version works for `p:sp`, `p:cxnSp`, and `p:pic` elements, others will need to
   * override.
   */
  getOrAddXfrm() {
    return this.spPr.getOrAddXfrm();
  }

  /**
   * True if this shape element has a `p:ph` descendant, indicating it
   * is a placeholder shape. False otherwise.
   */
  get hasPhElm() {
    return this.ph !== null;
  }

  /** The `p:ph` descendant element if there is one, None otherwise. */
  get ph(): CT_Placeholder | null {
    const phElms = this.xpath("./*[1]/p:nvPr/p:ph");
    if (phElms.length === 0) {
      return null;
    }
    return phElms[0];
  }

  /**
   * Integer value of placeholder idx attribute.
   *
   * Raises |ValueError| if shape is not a placeholder.
   */
  get phIdx(): number {
    const ph = this.ph;
    if (ph === null) {
      throw new ValueError("not a placeholder shape");
    }
    return ph.idx;
  }

  /**
   * Placeholder orientation, e.g. 'vert'.
   *
   * Raises |ValueError| if shape is not a placeholder.
   */
  get phOrient(): string {
    const ph = this.ph;
    if (ph === null) {
      throw new ValueError("not a placeholder shape");
    }
    return ph.orient;
  }

  /**
   * Placeholder size, e.g. ST_PlaceholderSize.HALF.
   *
   * Raises `ValueError` if shape is not a placeholder.
   */
  get phSz(): string {
    const ph = this.ph;
    if (ph === null) {
      throw new ValueError("not a placeholder shape");
    }
    return ph.sz;
  }

  /**
   * Placeholder type, e.g. ST_PlaceholderType.TITLE ('title').
   *
   * Raises `ValueError` if shape is not a placeholder.
   */
  get phType() {
    const ph = this.ph;
    if (ph === null) {
      throw new ValueError("not a placeholder shape");
    }
    return ph.type;
  }

  /** Float representing degrees this shape is rotated clockwise. */
  get rot(): number {
    const xfrm = this.xfrm;
    if ((xfrm === null || xfrm.rot === null)) {
      return 0.0;
    }
    return xfrm.rot;
  }

  set rot(value: number) {
    this.getOrAddXfrm().rot = value;
  }

  /** Integer id of this shape */
  get shapeId() {
    return this._nvXxPr.cNvPr.id;
  }

  /** Name of this shape */
  get shapeName() {
    return this._nvXxPr.cNvPr.name;
  }

  /** Child `p:txBody` element, None if not present. */
  get txBody() {
    return this.find(qn("p:txBody"));
  }

  get x(): number | null {
    return this._getXfrmAttr("x");
  }

  set x(value: any) {
    this._setXfrmAttr("x", value);
  }

  /**
   * The `a:xfrm` grandchild element or |None| if not found.
   *
   * This version works for `p:sp`, `p:cxnSp`, and `p:pic` elements, others will need to
   * override.
   */
  get xfrm() {
    return this.spPr.xfrm;
  }

  get y(): number | null {
    return this._getXfrmAttr("y");
  }

  set y(value: any) {
    this._setXfrmAttr("y", value);
  }

  /**
   * Required non-visual shape properties element for this shape. Actual
   * name depends on the shape type, e.g. `p:nvPicPr` for picture
   * shape.
   */
  get _nvXxPr() {
    return this.xpath("./*[1]")[0];
  }

  _getXfrmAttr(name: string): number | null {
    const xfrm = this.xfrm;
    if (xfrm === null) {
      return null;
    }
    return (xfrm as any)[name];
  }

  _setXfrmAttr(name: string, value: any): void {
    const xfrm = this.getOrAddXfrm();
    (xfrm as any)[name] = value;
  }
}

/** `p:nvPr` element. */
export class CT_ApplicationNonVisualDrawingProps extends BaseOxmlElement {
  declare getOrAddPh: (...args: any[]) => CT_Placeholder;
  declare ph: any;
  declare _addPh: (attrs?: Record<string, any>) => any;
  declare _insertPh: (child: any) => any;
  declare _newPh: () => any;
  declare _removePh: () => void;
  static {
    defineChildren(this, {
      ph: new ZeroOrOne("p:ph", ["a:audioCd", "a:wavAudioFile", "a:audioFile", "a:videoFile", "a:quickTimeFile", "p:custDataLst", "p:extLst"]),
    });
  }
}

/** Custom element class for <a:ln> element */
export class CT_LineProperties extends BaseOxmlElement {
  declare egLineFillProperties: any;
  declare _removeEgLineFillProperties: () => void;
  declare noFill: any;
  declare _newNoFill: () => any;
  declare _insertNoFill: (child: any) => any;
  declare _addNoFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToNoFill: () => any;
  declare solidFill: any;
  declare _newSolidFill: () => any;
  declare _insertSolidFill: (child: any) => any;
  declare _addSolidFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSolidFill: () => any;
  declare gradFill: any;
  declare _newGradFill: () => any;
  declare _insertGradFill: (child: any) => any;
  declare _addGradFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToGradFill: () => any;
  declare pattFill: any;
  declare _newPattFill: () => any;
  declare _insertPattFill: (child: any) => any;
  declare _addPattFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToPattFill: () => any;
  declare prstDash: any;
  declare getOrAddPrstDash: () => any;
  declare _addPrstDash: (attrs?: Record<string, any>) => any;
  declare _insertPrstDash: (child: any) => any;
  declare _newPrstDash: () => any;
  declare _removePrstDash: () => void;
  declare custDash: any;
  declare getOrAddCustDash: () => any;
  declare _addCustDash: (attrs?: Record<string, any>) => any;
  declare _insertCustDash: (child: any) => any;
  declare _newCustDash: () => any;
  declare _removeCustDash: () => void;
  declare w: any;
  static {
    const _tagSeq = ["a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:prstDash", "a:custDash", "a:round", "a:bevel", "a:miter", "a:headEnd", "a:tailEnd", "a:extLst"] as const;
    defineChildren(this, {
      egLineFillProperties: new ZeroOrOneChoice([new Choice("a:noFill"), new Choice("a:solidFill"), new Choice("a:gradFill"), new Choice("a:pattFill")], _tagSeq.slice(4)),
      prstDash: new ZeroOrOne("a:prstDash", _tagSeq.slice(5)),
      custDash: new ZeroOrOne("a:custDash", _tagSeq.slice(6)),
      w: new OptionalAttribute("w", ST_LineWidth, Emu(0)),
    });
  }

  /** Required to fulfill the interface used by dml.fill. */
  get egFillProperties() {
    return this.egLineFillProperties;
  }

  /**
   * Return value of `val` attribute of `a:prstDash` child.
   *
   * Return |None| if not present.
   */
  get prstDashVal() {
    const prstDash = this.prstDash;
    if (prstDash === null) {
      return null;
    }
    return prstDash.val;
  }

  set prstDashVal(val: any) {
    this._removeCustDash();
    const prstDash = this.getOrAddPrstDash();
    prstDash.val = val;
  }
}

/** `p:cNvPr` custom element class. */
export class CT_NonVisualDrawingProps extends BaseOxmlElement {
  declare getOrAddHlinkClick: (...args: any[]) => CT_Hyperlink;
  declare getOrAddHlinkHover: (...args: any[]) => CT_Hyperlink;
  declare hlinkClick: CT_Hyperlink | null;
  declare _addHlinkClick: (attrs?: Record<string, any>) => CT_Hyperlink;
  declare _insertHlinkClick: (child: CT_Hyperlink) => CT_Hyperlink;
  declare _newHlinkClick: () => CT_Hyperlink;
  declare _removeHlinkClick: () => void;
  declare hlinkHover: CT_Hyperlink | null;
  declare _addHlinkHover: (attrs?: Record<string, any>) => CT_Hyperlink;
  declare _insertHlinkHover: (child: CT_Hyperlink) => CT_Hyperlink;
  declare _newHlinkHover: () => CT_Hyperlink;
  declare _removeHlinkHover: () => void;
  declare id: any;
  declare name: any;
  static {
    const _tagSeq = ["a:hlinkClick", "a:hlinkHover", "a:extLst"] as const;
    defineChildren(this, {
      hlinkClick: new ZeroOrOne("a:hlinkClick", _tagSeq.slice(1)),
      hlinkHover: new ZeroOrOne("a:hlinkHover", _tagSeq.slice(2)),
      id: new RequiredAttribute("id", ST_DrawingElementId),
      name: new RequiredAttribute("name", XsdString),
    });
  }
}

/** `p:ph` custom element class. */
export class CT_Placeholder extends BaseOxmlElement {
  declare type: PP_PLACEHOLDER;
  declare orient: string;
  declare sz: string;
  declare idx: number;
  static {
    defineChildren(this, {
      type: new OptionalAttribute("type", PP_PLACEHOLDER, PP_PLACEHOLDER.OBJECT),
      orient: new OptionalAttribute("orient", ST_Direction, ST_Direction.HORZ),
      sz: new OptionalAttribute("sz", ST_PlaceholderSize, ST_PlaceholderSize.FULL),
      idx: new OptionalAttribute("idx", XsdUnsignedInt, 0),
    });
  }
}

/** Custom element class for <a:off> element. */
export class CT_Point2D extends BaseOxmlElement {
  declare x: number;
  declare y: number;
  static {
    defineChildren(this, {
      x: new RequiredAttribute("x", ST_Coordinate),
      y: new RequiredAttribute("y", ST_Coordinate),
    });
  }
}

/** Custom element class for <a:ext> element. */
export class CT_PositiveSize2D extends BaseOxmlElement {
  declare cx: any;
  declare cy: any;
  static {
    defineChildren(this, {
      cx: new RequiredAttribute("cx", ST_PositiveCoordinate),
      cy: new RequiredAttribute("cy", ST_PositiveCoordinate),
    });
  }
}

/**
 * Custom element class for `p:spPr` element.
 *
 * Shared by `p:sp`, `p:cxnSp`,  and `p:pic` elements as well as a few more obscure ones.
 */
export class CT_ShapeProperties extends BaseOxmlElement {
  declare getOrAddXfrm: (...args: any[]) => CT_Transform2D;
  declare getOrAddLn: (...args: any[]) => CT_LineProperties;
  declare _addPrstGeom: (...args: any[]) => CT_PresetGeometry2D;
  declare _removeCustGeom: (...args: any[]) => null;
  declare xfrm: CT_Transform2D | null;
  declare _addXfrm: (attrs?: Record<string, any>) => CT_Transform2D;
  declare _insertXfrm: (child: CT_Transform2D) => CT_Transform2D;
  declare _newXfrm: () => CT_Transform2D;
  declare _removeXfrm: () => void;
  declare custGeom: CT_CustomGeometry2D | null;
  declare getOrAddCustGeom: () => CT_CustomGeometry2D;
  declare _addCustGeom: (attrs?: Record<string, any>) => CT_CustomGeometry2D;
  declare _insertCustGeom: (child: CT_CustomGeometry2D) => CT_CustomGeometry2D;
  declare _newCustGeom: () => CT_CustomGeometry2D;
  declare prstGeom: CT_PresetGeometry2D | null;
  declare getOrAddPrstGeom: () => CT_PresetGeometry2D;
  declare _insertPrstGeom: (child: CT_PresetGeometry2D) => CT_PresetGeometry2D;
  declare _newPrstGeom: () => CT_PresetGeometry2D;
  declare _removePrstGeom: () => void;
  declare egFillProperties: any;
  declare _removeEgFillProperties: () => void;
  declare noFill: any;
  declare _newNoFill: () => any;
  declare _insertNoFill: (child: any) => any;
  declare _addNoFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToNoFill: () => any;
  declare solidFill: any;
  declare _newSolidFill: () => any;
  declare _insertSolidFill: (child: any) => any;
  declare _addSolidFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSolidFill: () => any;
  declare gradFill: any;
  declare _insertGradFill: (child: any) => any;
  declare _addGradFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToGradFill: () => any;
  declare blipFill: any;
  declare _newBlipFill: () => any;
  declare _insertBlipFill: (child: any) => any;
  declare _addBlipFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToBlipFill: () => any;
  declare pattFill: any;
  declare _newPattFill: () => any;
  declare _insertPattFill: (child: any) => any;
  declare _addPattFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToPattFill: () => any;
  declare grpFill: any;
  declare _newGrpFill: () => any;
  declare _insertGrpFill: (child: any) => any;
  declare _addGrpFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToGrpFill: () => any;
  declare ln: CT_LineProperties | null;
  declare _addLn: (attrs?: Record<string, any>) => CT_LineProperties;
  declare _insertLn: (child: CT_LineProperties) => CT_LineProperties;
  declare _newLn: () => CT_LineProperties;
  declare _removeLn: () => void;
  declare effectLst: any;
  declare getOrAddEffectLst: () => any;
  declare _addEffectLst: (attrs?: Record<string, any>) => any;
  declare _insertEffectLst: (child: any) => any;
  declare _newEffectLst: () => any;
  declare _removeEffectLst: () => void;
  static {
    const _tagSeq = ["a:xfrm", "a:custGeom", "a:prstGeom", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:ln", "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst"] as const;
    defineChildren(this, {
      xfrm: new ZeroOrOne("a:xfrm", _tagSeq.slice(1)),
      custGeom: new ZeroOrOne("a:custGeom", _tagSeq.slice(2)),
      prstGeom: new ZeroOrOne("a:prstGeom", _tagSeq.slice(3)),
      egFillProperties: new ZeroOrOneChoice([new Choice("a:noFill"), new Choice("a:solidFill"), new Choice("a:gradFill"), new Choice("a:blipFill"), new Choice("a:pattFill"), new Choice("a:grpFill")], _tagSeq.slice(9)),
      ln: new ZeroOrOne("a:ln", _tagSeq.slice(10)),
      effectLst: new ZeroOrOne("a:effectLst", _tagSeq.slice(11)),
    });
  }

  /** Shape width as an instance of Emu, or None if not present. */
  get cx() {
    const cxStrLst = this.xpath("./a:xfrm/a:ext/@cx");
    if (cxStrLst.length === 0) {
      return null;
    }
    return Emu(cxStrLst[0]);
  }

  /** Shape height as an instance of Emu, or None if not present. */
  get cy() {
    const cyStrLst = this.xpath("./a:xfrm/a:ext/@cy");
    if (cyStrLst.length === 0) {
      return null;
    }
    return Emu(cyStrLst[0]);
  }

  /**
   * Distance between the left edge of the slide and left edge of the shape.
   *
   * 0 if not present.
   */
  get x(): number | null {
    const xStrLst = this.xpath("./a:xfrm/a:off/@x");
    if (xStrLst.length === 0) {
      return null;
    }
    return Emu(xStrLst[0]);
  }

  /**
   * The offset of the top of the shape from the top of the slide, as an
   * instance of Emu. None if not present.
   */
  get y() {
    const yStrLst = this.xpath("./a:xfrm/a:off/@y");
    if (yStrLst.length === 0) {
      return null;
    }
    return Emu(yStrLst[0]);
  }

  _newGradFill() {
    return CT_GradientFillProperties.newGradFill();
  }
}

/**
 * `a:xfrm` custom element class.
 *
 * NOTE: this is a composite including CT_GroupTransform2D, which appears
 * with the `a:xfrm` tag in a group shape (including a slide `p:spTree`).
 */
export class CT_Transform2D extends BaseOxmlElement {
  declare off: CT_Point2D | null;
  declare getOrAddOff: () => CT_Point2D;
  declare _addOff: (attrs?: Record<string, any>) => CT_Point2D;
  declare _insertOff: (child: CT_Point2D) => CT_Point2D;
  declare _removeOff: () => void;
  declare ext: any;
  declare getOrAddExt: () => any;
  declare _addExt: (attrs?: Record<string, any>) => any;
  declare _insertExt: (child: any) => any;
  declare _removeExt: () => void;
  declare chOff: any;
  declare getOrAddChOff: () => any;
  declare _addChOff: (attrs?: Record<string, any>) => any;
  declare _insertChOff: (child: any) => any;
  declare _newChOff: () => any;
  declare _removeChOff: () => void;
  declare chExt: any;
  declare getOrAddChExt: () => any;
  declare _addChExt: (attrs?: Record<string, any>) => any;
  declare _insertChExt: (child: any) => any;
  declare _newChExt: () => any;
  declare _removeChExt: () => void;
  declare rot: number | null;
  declare flipH: any;
  declare flipV: any;
  static {
    const _tagSeq = ["a:off", "a:ext", "a:chOff", "a:chExt"] as const;
    defineChildren(this, {
      off: new ZeroOrOne("a:off", _tagSeq.slice(1)),
      ext: new ZeroOrOne("a:ext", _tagSeq.slice(2)),
      chOff: new ZeroOrOne("a:chOff", _tagSeq.slice(3)),
      chExt: new ZeroOrOne("a:chExt", _tagSeq.slice(4)),
      rot: new OptionalAttribute("rot", ST_Angle, 0.0),
      flipH: new OptionalAttribute("flipH", XsdBoolean, false),
      flipV: new OptionalAttribute("flipV", XsdBoolean, false),
    });
  }

  get x() {
    const off = this.off;
    if (off === null) {
      return null;
    }
    return off.x;
  }

  set x(value: any) {
    const off = this.getOrAddOff();
    off.x = value;
  }

  get y() {
    const off = this.off;
    if (off === null) {
      return null;
    }
    return off.y;
  }

  set y(value: any) {
    const off = this.getOrAddOff();
    off.y = value;
  }

  get cx() {
    const ext = this.ext;
    if (ext === null) {
      return null;
    }
    return ext.cx;
  }

  set cx(value: any) {
    const ext = this.getOrAddExt();
    ext.cx = value;
  }

  get cy() {
    const ext = this.ext;
    if (ext === null) {
      return null;
    }
    return ext.cy;
  }

  set cy(value: any) {
    const ext = this.getOrAddExt();
    ext.cy = value;
  }

  _newExt() {
    const ext = OxmlElement("a:ext");
    ext.cx = 0;
    ext.cy = 0;
    return ext;
  }

  _newOff() {
    const off = OxmlElement("a:off");
    off.x = 0;
    off.y = 0;
    return off;
  }
}
