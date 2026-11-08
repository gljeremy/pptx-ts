// Shapes based on the `p:pic` element, including Picture and Movie. (`pptx.shapes.picture`)
import { ValueError } from "../errors.js";
import { lazy } from "../util.js";
import { LineFormat } from "../dml/line.js";
import { MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE } from "../enum/shapes.js";
import { BaseShape } from "./base.js";
import { ParentedElementProxy } from "../shared.js";
import type { CT_Picture } from "../oxml/shapes/picture.js";
import type { CT_LineProperties } from "../oxml/shapes/shared.js";
import type { ProvidesPart } from "../types.js";

/** Base class for shapes based on a `p:pic` element. */
export class _BasePicture extends BaseShape {
  declare _pic: any;

  constructor(pic: CT_Picture, parent: ProvidesPart) {
    super(pic, parent);
    this._pic = pic;
  }

  /**
   * |float| representing relative portion cropped from shape bottom.
   *
   * Read/write. 1.0 represents 100%. For example, 25% is represented by 0.25. Negative values
   * are valid as are values greater than 1.0.
   */
  get cropBottom(): number {
    return this._pic.srcRectB;
  }

  set cropBottom(value: number) {
    this._pic.srcRectB = value;
  }

  /**
   * |float| representing relative portion cropped from left of shape.
   *
   * Read/write. 1.0 represents 100%. A negative value extends the side beyond the image
   * boundary.
   */
  get cropLeft(): number {
    return this._pic.srcRectL;
  }

  set cropLeft(value: number) {
    this._pic.srcRectL = value;
  }

  /**
   * |float| representing relative portion cropped from right of shape.
   *
   * Read/write. 1.0 represents 100%.
   */
  get cropRight(): number {
    return this._pic.srcRectR;
  }

  set cropRight(value: number) {
    this._pic.srcRectR = value;
  }

  /**
   * |float| representing relative portion cropped from shape top.
   *
   * Read/write. 1.0 represents 100%.
   */
  get cropTop(): number {
    return this._pic.srcRectT;
  }

  set cropTop(value: number) {
    this._pic.srcRectT = value;
  }

  /**
   * Return the `a:ln` element for this `p:pic`-based image.
   *
   * The `a:ln` element contains the line format properties XML.
   */
  getOrAddLn() {
    return this._pic.getOrAddLn();
  }

  /** Provides access to properties of the picture outline, such as its color and width. */
  get line(): LineFormat {
    return lazy(this, "line", () => {
      return new LineFormat(this);
    });
  }

  /**
   * The `a:ln` element for this `p:pic`.
   *
   * Contains the line format properties such as line color and width. |None| if no `a:ln`
   * element is present.
   */
  get ln(): CT_LineProperties | null {
    return this._pic.ln;
  }
}

/**
 * A movie shape, one that places a video on a slide.
 *
 * Like |Picture|, a movie shape is based on the `p:pic` element. A movie is composed of a video
 * and a *poster frame*, the placeholder image that represents the video before it is played.
 */
export class Movie extends _BasePicture {

  /**
   * The |_MediaFormat| object for this movie.
   *
   * The |_MediaFormat| object provides access to formatting properties for the movie.
   */
  get mediaFormat(): _MediaFormat {
    return lazy(this, "mediaFormat", () => {
      return new _MediaFormat(this._pic, this);
    });
  }

  /**
   * Member of :ref:`PpMediaType` describing this shape.
   *
   * The return value is unconditionally `PP_MEDIA_TYPE.MOVIE` in this case.
   */
  get mediaType(): PP_MEDIA_TYPE {
    return PP_MEDIA_TYPE.MOVIE;
  }

  /**
   * Return |Image| object containing poster frame for this movie.
   *
   * Returns |None| if this movie has no poster frame (uncommon).
   */
  get posterFrame() {
    const [slidePart, rId] = [this.part, this._pic.blipRId];
    if (rId === null) {
      return null;
    }
    return slidePart.getImage(rId);
  }

  /**
   * Return member of :ref:`MsoShapeType` describing this shape.
   *
   * The return value is unconditionally `MSO_SHAPE_TYPE.MEDIA` in this
   * case.
   */
  get shapeType(): MSO_SHAPE_TYPE {
    return MSO_SHAPE_TYPE.MEDIA;
  }
}

/**
 * A picture shape, one that places an image on a slide.
 *
 * Based on the `p:pic` element.
 */
export class Picture extends _BasePicture {

  /**
   * Member of MSO_SHAPE indicating masking shape.
   *
   * A picture can be masked by any of the so-called "auto-shapes" available in PowerPoint,
   * such as an ellipse or triangle. When a picture is masked by a shape, the shape assumes the
   * same dimensions as the picture and the portion of the picture outside the shape boundaries
   * does not appear. Note the default value for a newly-inserted picture is
   * `MSO_AUTO_SHAPE_TYPE.RECTANGLE`, which performs no cropping because the extents of the
   * rectangle exactly correspond to the extents of the picture.
   *
   * The available shapes correspond to the members of :ref:`MsoAutoShapeType`.
   *
   * The return value can also be |None|, indicating the picture either has no geometry (not
   * expected) or has custom geometry, like a freeform shape. A picture with no geometry will
   * have no visible representation on the slide, although it can be selected. This is because
   * without geometry, there is no "inside-the-shape" for it to appear in.
   */
  get autoShapeType(): MSO_SHAPE | null {
    const prstGeom = this._pic.spPr.prstGeom;
    if (prstGeom === null) {
      return null;
    }
    return prstGeom.prst;
  }

  set autoShapeType(member: MSO_SHAPE) {
    let prstGeom;
    MSO_SHAPE.validate(member);
    const spPr = this._pic.spPr;
    prstGeom = spPr.prstGeom;
    if (prstGeom === null) {
      spPr._removeCustGeom();
      prstGeom = spPr._addPrstGeom();
    }
    prstGeom.prst = member;
  }

  /**
   * The |Image| object for this picture.
   *
   * Provides access to the properties and bytes of the image in this picture shape.
   */
  get image() {
    const [slidePart, rId] = [this.part, this._pic.blipRId];
    if (rId === null) {
      throw new ValueError("no embedded image");
    }
    return slidePart.getImage(rId);
  }

  /** Unconditionally `MSO_SHAPE_TYPE.PICTURE` in this case. */
  get shapeType(): MSO_SHAPE_TYPE {
    return MSO_SHAPE_TYPE.PICTURE;
  }
}

/**
 * Provides access to formatting properties for a Media object.
 *
 * Media format properties are things like start point, volume, and
 * compression type.
 */
export class _MediaFormat extends ParentedElementProxy {
}
