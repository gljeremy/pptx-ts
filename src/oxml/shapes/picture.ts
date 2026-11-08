// lxml custom element classes for picture-related XML elements. (`pptx.oxml.shapes.picture`)
import { defineChildren } from "../xmlchemy.js";
import { pyPercent } from "../../util.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { BaseShapeElement } from "./shared.js";
import { BaseOxmlElement, OneAndOnlyOne } from "../xmlchemy.js";
import type { CT_ShapeProperties } from "./shared.js";
import type { Length } from "../../util.js";

/**
 * `p:pic` element.
 *
 * Represents a picture shape (an image placement on a slide).
 */
export class CT_Picture extends BaseShapeElement {
  declare nvPicPr: any;
  declare blipFill: any;
  declare spPr: CT_ShapeProperties;
  static {
    defineChildren(this, {
      nvPicPr: new OneAndOnlyOne("p:nvPicPr"),
      blipFill: new OneAndOnlyOne("p:blipFill"),
      spPr: new OneAndOnlyOne("p:spPr"),
    });
  }

  /**
   * Value of `p:blipFill/a:blip/@r:embed`.
   *
   * Returns |None| if not present.
   */
  get blipRId(): string | null {
    const blip = this.blipFill.blip;
    if ((blip !== null && blip.rEmbed !== null)) {
      return blip.rEmbed;
    }
    return null;
  }

  /**
   * Set cropping values in `p:blipFill/a:srcRect` such that an image of
   * *image_size* will stretch to exactly fit *view_size* when its aspect
   * ratio is preserved.
   */
  cropToFit(imageSize: any, viewSize: any) {
    this.blipFill.crop(this._fillCropping(imageSize, viewSize));
  }

  /** Return the <a:ln> grandchild element, newly added if not present. */
  getOrAddLn() {
    return this.spPr.getOrAddLn();
  }

  /** ``<a:ln>`` grand-child element or |None| if not present */
  get ln() {
    return this.spPr.ln;
  }

  /**
   * Return a new `p:pic` placeholder element populated with the supplied
   * parameters.
   */
  static newPhPic(id_: any, name: any, desc: any, rId: any) {
    return parseXml(pyPercent(this._picPhTmpl(), [id_, name, desc, rId]));
  }

  /** Return new `<p:pic>` element tree configured with supplied parameters. */
  static newPic(shapeId: any, name: any, desc: any, rId: any, x: any, y: any, cx: any, cy: any) {
    return parseXml(pyPercent(this._picTmpl(), [shapeId, name, escape(desc), rId, x, y, cx, cy]));
  }

  /** Return a new `p:pic` populated with the specified video. */
  static newVideoPic(shapeId: number, shapeName: string, videoRId: string, mediaRId: string, posterFrameRId: string, x: number, y: number, cx: number, cy: number): CT_Picture {
    return parseXml(pyPercent(this._picVideoTmpl(), [shapeId, shapeName, videoRId, mediaRId, posterFrameRId, x, y, cx, cy]));
  }

  /** Value of `p:blipFill/a:srcRect/@b` or 0.0 if not present. */
  get srcRectB() {
    return this._srcRectX("b");
  }

  set srcRectB(value: any) {
    this.blipFill.getOrAddSrcRect().b = value;
  }

  /** Value of `p:blipFill/a:srcRect/@l` or 0.0 if not present. */
  get srcRectL() {
    return this._srcRectX("l");
  }

  set srcRectL(value: any) {
    this.blipFill.getOrAddSrcRect().l = value;
  }

  /** Value of `p:blipFill/a:srcRect/@r` or 0.0 if not present. */
  get srcRectR() {
    return this._srcRectX("r");
  }

  set srcRectR(value: any) {
    this.blipFill.getOrAddSrcRect().r = value;
  }

  /** Value of `p:blipFill/a:srcRect/@t` or 0.0 if not present. */
  get srcRectT() {
    return this._srcRectX("t");
  }

  set srcRectT(value: any) {
    this.blipFill.getOrAddSrcRect().t = value;
  }

  /**
   * Return a (left, top, right, bottom) 4-tuple containing the cropping
   * values required to display an image of *image_size* in *view_size*
   * when stretched proportionately. Each value is a percentage expressed
   * as a fraction of 1.0, e.g. 0.425 represents 42.5%. *image_size* and
   * *view_size* are each (width, height) pairs.
   */
  _fillCropping(imageSize: readonly [number, number], viewSize: readonly [number, number]): [number, number, number, number] {
    const aspectRatio = (width: number, height: number) => width / height;
    const arView = aspectRatio(viewSize[0], viewSize[1]);
    const arImage = aspectRatio(imageSize[0], imageSize[1]);
    if (arView < arImage) {
      const crop = (1.0 - arView / arImage) / 2.0;
      return [crop, 0.0, crop, 0.0];
    }
    if (arView > arImage) {
      const crop = (1.0 - arImage / arView) / 2.0;
      return [0.0, crop, 0.0, crop];
    }
    return [0.0, 0.0, 0.0, 0.0];
  }

  static _picPhTmpl() {
    return `<p:pic ${nsdecls("p", "a", "r")}>
  <p:nvPicPr>
    <p:cNvPr id="%d" name="%s" descr="%s"/>
    <p:cNvPicPr>
      <a:picLocks noGrp="1" noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="%s"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr/>
</p:pic>`;
  }

  static _picTmpl() {
    return `<p:pic ${nsdecls("a", "p", "r")}>
  <p:nvPicPr>
    <p:cNvPr id="%d" name="%s" descr="%s"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="%s"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="%d" y="%d"/>
      <a:ext cx="%d" cy="%d"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;
  }

  static _picVideoTmpl() {
    return `<p:pic ${nsdecls("a", "p", "r")}>
  <p:nvPicPr>
    <p:cNvPr id="%d" name="%s">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr>
      <a:videoFile r:link="%s"/>
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="%s"/>
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="%s"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="%d" y="%d"/>
      <a:ext cx="%d" cy="%d"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;
  }

  /** Value of `p:blipFill/a:srcRect/@{attr_name}` or 0.0 if not present. */
  _srcRectX(attrName: any) {
    const srcRect = this.blipFill.srcRect;
    if (srcRect === null) {
      return 0.0;
    }
    return srcRect[attrName];
  }
}

/**
 * ``<p:nvPicPr>`` element, containing non-visual properties for a picture
 * shape.
 */
export class CT_PictureNonVisual extends BaseOxmlElement {
  declare cNvPr: any;
  declare nvPr: any;
  static {
    defineChildren(this, {
      cNvPr: new OneAndOnlyOne("p:cNvPr"),
      nvPr: new OneAndOnlyOne("p:nvPr"),
    });
  }
}
