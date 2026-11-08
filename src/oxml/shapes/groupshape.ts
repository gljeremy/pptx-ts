// lxml custom element classes for shape-tree-related XML elements. (`pptx.oxml.shapes.groupshape`)
import { defineChildren } from "../xmlchemy.js";
import { pyInt, pyPercent } from "../../util.js";
import { MSO_CONNECTOR_TYPE } from "../../enum/shapes.js";
import { parseXml } from "../parser.js";
import { nsdecls, qn } from "../ns.js";
import { CT_Shape } from "./autoshape.js";
import { CT_Connector } from "./connector.js";
import { CT_GraphicalObjectFrame } from "./graphfrm.js";
import { CT_Picture } from "./picture.js";
import { BaseShapeElement } from "./shared.js";
import { BaseOxmlElement, OneAndOnlyOne, ZeroOrOne } from "../xmlchemy.js";
import { Emu } from "../../util.js";
import type { PP_PLACEHOLDER } from "../../enum/shapes.js";
import type { ShapeElement } from "./index.js";
import type { CT_Transform2D } from "./shared.js";

/** Used for shape tree (`p:spTree`) as well as the group shape (`p:grpSp`) elements. */
export class CT_GroupShape extends BaseShapeElement {
  declare nvGrpSpPr: CT_GroupShapeNonVisual;
  declare grpSpPr: CT_GroupShapeProperties;
  static _shapeTags = [qn("p:sp"), qn("p:grpSp"), qn("p:graphicFrame"), qn("p:cxnSp"), qn("p:pic"), qn("p:contentPart")];
  static {
    defineChildren(this, {
      nvGrpSpPr: new OneAndOnlyOne("p:nvGrpSpPr"),
      grpSpPr: new OneAndOnlyOne("p:grpSpPr"),
    });
  }

  /** Return new `p:sp` appended to the group/shapetree with specified attributes. */
  addAutoshape(id_: number, name: string, prst: string, x: number, y: number, cx: number, cy: number): CT_Shape {
    const sp = CT_Shape.newAutoshapeSp(id_, name, prst, x, y, cx, cy);
    this.insertElementBefore(sp, "p:extLst");
    return sp;
  }

  /** Return new `p:cxnSp` appended to the group/shapetree with the specified attribues. */
  addCxnSp(id_: number, name: string, typeMember: MSO_CONNECTOR_TYPE, x: number, y: number, cx: number, cy: number, flipH: boolean, flipV: boolean): CT_Connector {
    const prst = MSO_CONNECTOR_TYPE.toXml(typeMember);
    const cxnSp = CT_Connector.newCxnSp(id_, name, prst, x, y, cx, cy, flipH, flipV);
    this.insertElementBefore(cxnSp, "p:extLst");
    return cxnSp;
  }

  /** Append a new freeform `p:sp` with specified position and size. */
  addFreeformSp(x: number, y: number, cx: number, cy: number): CT_Shape {
    const shapeId = this._nextShapeId;
    const name = `Freeform ${(shapeId - 1)}`;
    const sp = CT_Shape.newFreeformSp(shapeId, name, x, y, cx, cy);
    this.insertElementBefore(sp, "p:extLst");
    return sp;
  }

  /**
   * Return `p:grpSp` element newly appended to this shape tree.
   *
   * The element contains no sub-shapes, is positioned at (0, 0), and has
   * width and height of zero.
   */
  addGrpSp(): CT_GroupShape {
    const shapeId = this._nextShapeId;
    const name = `Group ${(shapeId - 1)}`;
    const grpSp = CT_GroupShape.newGrpSp(shapeId, name);
    this.insertElementBefore(grpSp, "p:extLst");
    return grpSp;
  }

  /** Append a `p:pic` shape to the group/shapetree having properties as specified in call. */
  addPic(id_: number, name: string, desc: string, rId: string, x: number, y: number, cx: number, cy: number): CT_Picture {
    const pic = CT_Picture.newPic(id_, name, desc, rId, x, y, cx, cy);
    this.insertElementBefore(pic, "p:extLst");
    return pic;
  }

  /** Append a newly-created placeholder `p:sp` shape having the specified properties. */
  addPlaceholder(id_: number, name: string, phType: PP_PLACEHOLDER, orient: string, sz: string, idx: number): CT_Shape {
    const sp = CT_Shape.newPlaceholderSp(id_, name, phType, orient, sz, idx);
    this.insertElementBefore(sp, "p:extLst");
    return sp;
  }

  /** Append a `p:graphicFrame` shape containing a table as specified in call. */
  addTable(id_: number, name: string, rows: number, cols: number, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const graphicFrame = CT_GraphicalObjectFrame.newTableGraphicFrame(id_, name, rows, cols, x, y, cx, cy);
    this.insertElementBefore(graphicFrame, "p:extLst");
    return graphicFrame;
  }

  /** Append a newly-created textbox `p:sp` shape having the specified position and size. */
  addTextbox(id_: number, name: string, x: number, y: number, cx: number, cy: number): CT_Shape {
    const sp = CT_Shape.newTextboxSp(id_, name, x, y, cx, cy);
    this.insertElementBefore(sp, "p:extLst");
    return sp;
  }

  /** Descendent `p:grpSpPr/a:xfrm/a:chExt` element. */
  get chExt() {
    return this.grpSpPr.getOrAddXfrm().getOrAddChExt();
  }

  /** Descendent `p:grpSpPr/a:xfrm/a:chOff` element. */
  get chOff() {
    return this.grpSpPr.getOrAddXfrm().getOrAddChOff();
  }

  /** Return the `a:xfrm` grandchild element, newly-added if not present. */
  getOrAddXfrm(): CT_Transform2D {
    return this.grpSpPr.getOrAddXfrm();
  }

  /** Generate each placeholder shape child element in document order. */
  *iterPhElms() {
    for (const e of this.iterShapeElms()) {
      if (e.hasPhElm) {
        yield e;
      }
    }
  }

  /**
   * Generate each child of this `p:spTree` element that corresponds to a shape.
   *
   * Items appear in XML document order.
   */
  *iterShapeElms(): Generator<ShapeElement> {
    for (const elm of this.iterchildren()) {
      if (CT_GroupShape._shapeTags.includes(elm.tag)) {
        yield elm as ShapeElement;
      }
    }
  }

  /**
   * Maximum int value assigned as @id in this slide.
   *
   * This is generally a shape-id, but ids can be assigned to other
   * objects so we just check all @id values anywhere in the document
   * (XML id-values have document scope).
   *
   * In practice, its minimum value is 1 because the spTree element itself
   * is always assigned id="1".
   */
  get maxShapeId(): number {
    const idStrLst = this.xpath("//@id");
    const usedIds = [...idStrLst].filter((idStr) => /^\d+$/.test(idStr)).map((idStr) => pyInt(idStr));
    return usedIds.length ? Math.max(...usedIds) : 0;
  }

  /** Return new "loose" `p:grpSp` element having `id_` and `name`. */
  static newGrpSp(id_: number, name: string): CT_GroupShape {
    const xml = pyPercent(`<p:grpSp ${nsdecls("a", "p", "r")}>
  <p:nvGrpSpPr>
    <p:cNvPr id="%d" name="%s"/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
  </p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="0" cy="0"/>
    </a:xfrm>
  </p:grpSpPr>
</p:grpSp>`, [id_, name]);
    const grpSp = parseXml(xml);
    return grpSp;
  }

  /**
   * Adjust x, y, cx, and cy to incorporate all contained shapes.
   *
   * This would typically be called when a contained shape is added,
   * removed, or its position or size updated.
   *
   * This method is recursive "upwards" since a change in a group shape
   * can change the position and size of its containing group.
   */
  recalculateExtents(): void {
    if (this.tag !== qn("p:grpSp")) {
      return;
    }
    const [x, y, cx, cy] = this._childExtents;
    this.chOff.x = x;
    this.x = x;
    this.chOff.y = y;
    this.y = y;
    this.chExt.cx = cx;
    this.cx = cx;
    this.chExt.cy = cy;
    this.cy = cy;
    (this.getparent() as CT_GroupShape).recalculateExtents();
  }

  /** The `a:xfrm` grandchild element or |None| if not found. */
  get xfrm(): CT_Transform2D | null {
    return this.grpSpPr.xfrm;
  }

  /**
   * (x, y, cx, cy) tuple representing net position and size.
   *
   * The values are formed as a composite of the contained child shapes.
   */
  get _childExtents(): [number, number, number, number] {
    const childShapeElms = [...this.iterShapeElms()];
    if (childShapeElms.length === 0) {
      return [Emu(0), Emu(0), Emu(0), Emu(0)];
    }
    const minX = Math.min(...[...childShapeElms].map((xSp) => xSp.x!));
    const minY = Math.min(...[...childShapeElms].map((xSp) => xSp.y!));
    const maxX = Math.max(...[...childShapeElms].map((xSp) => xSp.x! + xSp.cx!));
    const maxY = Math.max(...[...childShapeElms].map((xSp) => xSp.y! + xSp.cy!));
    const x = minX;
    const y = minY;
    const cx = (maxX - minX);
    const cy = (maxY - minY);
    return [x, y, cx, cy];
  }

  /**
   * Return unique shape id suitable for use with a new shape element.
   *
   * The returned id is the next available positive integer drawing object
   * id in shape tree, starting from 1 and making use of any gaps in
   * numbering. In practice, the minimum id is 2 because the spTree
   * element itself is always assigned id="1".
   */
  get _nextShapeId(): number {
    const idStrLst = this.xpath("//@id");
    const usedIds = [...idStrLst].filter((idStr) => /^\d+$/.test(idStr)).map((idStr) => pyInt(idStr));
    for (let n = 1; n < (usedIds.length + 2); n++) {
      if (!usedIds.includes(n)) {
        return n;
      }
    }
    throw new Error("unreachable: no free shape id");
  }
}

/** `p:nvGrpSpPr` element. */
export class CT_GroupShapeNonVisual extends BaseShapeElement {
  declare cNvPr: any;
  static {
    defineChildren(this, {
      cNvPr: new OneAndOnlyOne("p:cNvPr"),
    });
  }
}

/** p:grpSpPr element */
export class CT_GroupShapeProperties extends BaseOxmlElement {
  declare getOrAddXfrm: (...args: any[]) => CT_Transform2D;
  declare xfrm: CT_Transform2D | null;
  declare _addXfrm: (attrs?: Record<string, any>) => CT_Transform2D;
  declare _insertXfrm: (child: CT_Transform2D) => CT_Transform2D;
  declare _newXfrm: () => CT_Transform2D;
  declare _removeXfrm: () => void;
  declare effectLst: any;
  declare getOrAddEffectLst: () => any;
  declare _addEffectLst: (attrs?: Record<string, any>) => any;
  declare _insertEffectLst: (child: any) => any;
  declare _newEffectLst: () => any;
  declare _removeEffectLst: () => void;
  static {
    const _tagSeq = ["a:xfrm", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:effectLst", "a:effectDag", "a:scene3d", "a:extLst"] as const;
    defineChildren(this, {
      xfrm: new ZeroOrOne("a:xfrm", _tagSeq.slice(1)),
      effectLst: new ZeroOrOne("a:effectLst", _tagSeq.slice(8)),
    });
  }
}
