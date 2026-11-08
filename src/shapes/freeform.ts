// Objects related to construction of freeform shapes. (`pptx.shapes.freeform`)
import { IndexError, NotImplementedError } from "../errors.js";
import { PY_GETITEM } from "../oxml/xml.js";
import { lazy, pyInt, pyRound } from "../util.js";
import { Emu } from "../util.js";
import type { CT_Path2D, CT_Path2DClose, CT_Path2DLineTo, CT_Path2DMoveTo, CT_Shape } from "../oxml/shapes/autoshape.js";
import type { _BaseGroupShapes } from "./shapetree.js";
import type { Length } from "../util.js";
export type CT_DrawingOperation = CT_Path2DClose | CT_Path2DLineTo | CT_Path2DMoveTo;
export type DrawingOperation = _LineSegment | _MoveTo | _Close;

/**
 * Allows a freeform shape to be specified and created.
 *
 * The initial pen position is provided on construction. From there, drawing proceeds using
 * successive calls to draw line segments. The freeform shape may be closed by calling the
 * :meth:`close` method.
 *
 * A shape may have more than one contour, in which case overlapping areas are "subtracted". A
 * contour is a sequence of line segments beginning with a "move-to" operation. A move-to
 * operation is automatically inserted in each new freeform; additional move-to ops can be
 * inserted with the `.move_to()` method.
 */
export class FreeformBuilder implements Iterable<DrawingOperation> {
  declare _shapes: any;
  declare _startX: any;
  declare _startY: any;
  declare _xScale: any;
  declare _yScale: any;

  constructor(shapes: _BaseGroupShapes, startX: number, startY: number, xScale: number, yScale: number) {
    this._shapes = shapes;
    this._startX = startX;
    this._startY = startY;
    this._xScale = xScale;
    this._yScale = yScale;
  }

  at(idx: number): DrawingOperation {
    const op = this._drawingOperations.at(idx);
    if (op === undefined) throw new IndexError("list index out of range");
    return op;
  }

  [PY_GETITEM](idx: number): DrawingOperation {
    return this.at(idx);
  }

  *[Symbol.iterator](): Generator<DrawingOperation> {
    yield* this._drawingOperations;
  }

  get length(): number {
    return this._drawingOperations.length;
  }

  /**
   * Return a new |FreeformBuilder| object.
   *
   * The initial pen location is specified (in local coordinates) by
   * (`start_x`, `start_y`).
   */
  static new(shapes: _BaseGroupShapes, startX: number, startY: number, xScale: number, yScale: number) {
    return new this(shapes, Emu(pyInt(pyRound(startX))), Emu(pyInt(pyRound(startY))), xScale, yScale);
  }

  /**
   * Add a straight line segment to each point in `vertices`.
   *
   * `vertices` must be an iterable of (x, y) pairs (2-tuples). Each x and y value is rounded
   * to the nearest integer before use. The optional `close` parameter determines whether the
   * resulting contour is `closed` or left `open`.
   *
   * Returns this |FreeformBuilder| object so it can be used in chained calls.
   */
  addLineSegments(vertices: [number, number][], close: boolean = true) {
    let x, y;
    for (const [x, y] of vertices) {
      this._addLineSegment(x, y);
    }
    if (close) {
      this._addClose();
    }
    return this;
  }

  /**
   * Return new freeform shape positioned relative to specified offset.
   *
   * `origin_x` and `origin_y` locate the origin of the local coordinate system in slide
   * coordinates (EMU), perhaps most conveniently by use of a |Length| object.
   *
   * Note that this method may be called more than once to add multiple shapes of the same
   * geometry in different locations on the slide.
   */
  convertToShape(originX: number = Emu(0), originY: number = Emu(0)) {
    let drawingOperation;
    const sp = this._addFreeformSp(originX, originY);
    const path = this._startPath(sp);
    for (const drawingOperation of this) {
      drawingOperation.applyOperationTo(path);
    }
    return this._shapes._shapeFactory(sp);
  }

  /**
   * Move pen to (x, y) (local coordinates) without drawing line.
   *
   * Returns this |FreeformBuilder| object so it can be used in chained calls.
   */
  moveTo(x: number, y: number) {
    this._drawingOperations.push(_MoveTo.new(this, x, y));
    return this;
  }

  /**
   * Return x distance of shape origin from local coordinate origin.
   *
   * The returned integer represents the leftmost extent of the freeform shape, in local
   * coordinates. Note that the bounding box of the shape need not start at the local origin.
   */
  get shapeOffsetX(): number {
    let minX, drawingOperation;
    minX = this._startX;
    for (const drawingOperation of this) {
      if (drawingOperation instanceof _Close) {
        continue;
      }
      minX = Math.min(minX, drawingOperation.x);
    }
    return Emu(minX);
  }

  /**
   * Return y distance of shape origin from local coordinate origin.
   *
   * The returned integer represents the topmost extent of the freeform shape, in local
   * coordinates. Note that the bounding box of the shape need not start at the local origin.
   */
  get shapeOffsetY(): number {
    let minY, drawingOperation;
    minY = this._startY;
    for (const drawingOperation of this) {
      if (drawingOperation instanceof _Close) {
        continue;
      }
      minY = Math.min(minY, drawingOperation.y);
    }
    return Emu(minY);
  }

  /** Add a close |_Close| operation to the drawing sequence. */
  _addClose() {
    this._drawingOperations.push(_Close.new());
  }

  /**
   * Add a freeform `p:sp` element having no drawing elements.
   *
   * `origin_x` and `origin_y` are specified in slide coordinates, and represent the location
   * of the local coordinates origin on the slide.
   */
  _addFreeformSp(originX: number, originY: number) {
    const spTree = this._shapes._spTree;
    return spTree.addFreeformSp((originX + this._left), (originY + this._top), this._width, this._height);
  }

  /** Add a |_LineSegment| operation to the drawing sequence. */
  _addLineSegment(x: number, y: number): void {
    this._drawingOperations.push(_LineSegment.new(this, x, y));
  }

  /** Return the sequence of drawing operation objects for freeform. */
  get _drawingOperations(): DrawingOperation[] {
    return lazy(this, "_drawingOperations", () => {
      return [];
    });
  }

  /** Return width of this shape's path in local units. */
  get _dx(): number {
    let minX, maxX, drawingOperation;
    minX = this._startX;
    maxX = this._startX;
    for (const drawingOperation of this) {
      if (drawingOperation instanceof _Close) {
        continue;
      }
      minX = Math.min(minX, drawingOperation.x);
      maxX = Math.max(maxX, drawingOperation.x);
    }
    return Emu((maxX - minX));
  }

  /** Return integer height of this shape's path in local units. */
  get _dy(): number {
    let minY, maxY, drawingOperation;
    minY = this._startY;
    maxY = this._startY;
    for (const drawingOperation of this) {
      if (drawingOperation instanceof _Close) {
        continue;
      }
      minY = Math.min(minY, drawingOperation.y);
      maxY = Math.max(maxY, drawingOperation.y);
    }
    return Emu((maxY - minY));
  }

  /**
   * Return vertical size of this shape's path in slide coordinates.
   *
   * This value is based on the actual extents of the shape and does not include any
   * positioning offset.
   */
  get _height() {
    return pyInt(pyRound((this._dy * this._yScale)));
  }

  /**
   * Return leftmost extent of this shape's path in slide coordinates.
   *
   * Note that this value does not include any positioning offset; it assumes the drawing
   * (local) coordinate origin is at (0, 0) on the slide.
   */
  get _left() {
    return pyInt(pyRound((this.shapeOffsetX * this._xScale)));
  }

  /**
   * Translate local coordinates point to shape coordinates.
   *
   * Shape coordinates have the same unit as local coordinates, but are offset such that the
   * origin of the shape coordinate system (0, 0) is located at the top-left corner of the
   * shape bounding box.
   */
  _localToShape(localX: number, localY: number): [number, number] {
    return [Emu((localX - this.shapeOffsetX)), Emu((localY - this.shapeOffsetY))];
  }

  /**
   * Return a newly created `a:path` element added to `sp`.
   *
   * The returned `a:path` element has an `a:moveTo` element representing the shape starting
   * point as its only child.
   */
  _startPath(sp: CT_Shape): CT_Path2D {
    const path = sp.addPath(this._dx, this._dy);
    path.addMoveTo(...this._localToShape(this._startX, this._startY));
    return path;
  }

  /**
   * Return topmost extent of this shape's path in slide coordinates.
   *
   * Note that this value does not include any positioning offset; it assumes the drawing
   * (local) coordinate origin is located at slide coordinates (0, 0) (top-left corner of
   * slide).
   */
  get _top() {
    return pyInt(pyRound((this.shapeOffsetY * this._yScale)));
  }

  /**
   * Return width of this shape's path in slide coordinates.
   *
   * This value is based on the actual extents of the shape path and does not include any
   * positioning offset.
   */
  get _width() {
    return pyInt(pyRound((this._dx * this._xScale)));
  }
}

/**
 * Base class for freeform drawing operations.
 *
 * A drawing operation has at least one location (x, y) in local coordinates.
 */
export class _BaseDrawingOperation {
  declare _freeformBuilder: any;
  declare _x: any;
  declare _y: any;

  constructor(freeformBuilder: FreeformBuilder, x: number, y: number) {
    this._freeformBuilder = freeformBuilder;
    this._x = x;
    this._y = y;
  }

  /**
   * Add the XML element(s) implementing this operation to `path`.
   *
   * Must be implemented by each subclass.
   */
  applyOperationTo(path: CT_Path2D): CT_DrawingOperation {
    throw new NotImplementedError("must be implemented by each subclass");
  }

  /**
   * Return the horizontal (x) target location of this operation.
   *
   * The returned value is an integer in local coordinates.
   */
  get x(): number {
    return this._x;
  }

  /**
   * Return the vertical (y) target location of this operation.
   *
   * The returned value is an integer in local coordinates.
   */
  get y(): number {
    return this._y;
  }
}

/** Specifies adding a `<a:close/>` element to the current contour. */
export class _Close {

  /** Return a new _Close object. */
  static new(): _Close {
    return new this();
  }

  /** Add `a:close` element to `path`. */
  applyOperationTo(path: CT_Path2D): CT_Path2DClose {
    return path.addClose();
  }
}

/** Specifies a straight line segment ending at the specified point. */
export class _LineSegment extends _BaseDrawingOperation {

  /**
   * Return a new _LineSegment object ending at point *(x, y)*.
   *
   * Both `x` and `y` are rounded to the nearest integer before use.
   */
  static new(freeformBuilder: FreeformBuilder, x: number, y: number): _LineSegment {
    return new this(freeformBuilder, Emu(pyInt(pyRound(x))), Emu(pyInt(pyRound(y))));
  }

  /**
   * Add `a:lnTo` element to `path` for this line segment.
   *
   * Returns the `a:lnTo` element newly added to the path.
   */
  applyOperationTo(path: CT_Path2D): CT_Path2DLineTo {
    return path.addLnTo(Emu((this._x - this._freeformBuilder.shapeOffsetX)), Emu((this._y - this._freeformBuilder.shapeOffsetY)));
  }
}

/** Specifies a new pen position. */
export class _MoveTo extends _BaseDrawingOperation {

  /**
   * Return a new _MoveTo object for move to point `(x, y)`.
   *
   * Both `x` and `y` are rounded to the nearest integer before use.
   */
  static new(freeformBuilder: FreeformBuilder, x: number, y: number): _MoveTo {
    return new this(freeformBuilder, Emu(pyInt(pyRound(x))), Emu(pyInt(pyRound(y))));
  }

  /** Add `a:moveTo` element to `path` for this line segment. */
  applyOperationTo(path: CT_Path2D): CT_Path2DMoveTo {
    return path.addMoveTo(Emu((this._x - this._freeformBuilder.shapeOffsetX)), Emu((this._y - this._freeformBuilder.shapeOffsetY)));
  }
}
