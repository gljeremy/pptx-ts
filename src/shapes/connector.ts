// Connector (line) shape and related objects. (`pptx.shapes.connector`)
import { KeyError } from "../errors.js";
import { lazy, pyInt } from "../util.js";
import { LineFormat } from "../dml/line.js";
import { MSO_SHAPE_TYPE } from "../enum/shapes.js";
import { BaseShape } from "./base.js";
import { Emu } from "../util.js";

/**
 * Connector (line) shape.
 *
 * A connector is a linear shape having end-points that can be connected to
 * other objects (but not to other connectors). A connector can be straight,
 * have elbows, or can be curved.
 */
export class Connector extends BaseShape {

  /**
   * **EXPERIMENTAL** - *The current implementation only works properly
   * with rectangular shapes, such as pictures and rectangles. Use with
   * other shape types may cause unexpected visual alignment of the
   * connected end-point and could lead to a load error if cxn_pt_idx
   * exceeds the connection point count available on the connected shape.
   * That said, a quick test should reveal what to expect when using this
   * method with other shape types.*
   *
   * Connect the beginning of this connector to *shape* at the connection
   * point specified by *cxn_pt_idx*. Each shape has zero or more
   * connection points and they are identified by index, starting with 0.
   * Generally, the first connection point of a shape is at the top center
   * of its bounding box and numbering proceeds counter-clockwise from
   * there. However this is only a convention and may vary, especially
   * with non built-in shapes.
   */
  beginConnect(shape: any, cxnPtIdx: any) {
    this._connectBeginTo(shape, cxnPtIdx);
    this._moveBeginToCxn(shape, cxnPtIdx);
  }

  /**
   * Return the X-position of the begin point of this connector, in
   * English Metric Units (as a |Length| object).
   */
  get beginX() {
    const cxnSp = this._element;
    const [x, cx, flipH] = [cxnSp.x, cxnSp.cx, cxnSp.flipH];
    const beginX = (flipH ? (x + cx) : x);
    return Emu(beginX);
  }

  set beginX(value: any) {
    let dx;
    const cxnSp = this._element;
    const [x, cx, flipH, newX] = [cxnSp.x, cxnSp.cx, cxnSp.flipH, pyInt(value)];
    if (flipH) {
      const oldX = (x + cx);
      dx = Math.abs((newX - oldX));
      if (newX >= oldX) {
        cxnSp.cx = (cx + dx);
      } else if (dx <= cx) {
        cxnSp.cx = (cx - dx);
      } else {
        cxnSp.flipH = false;
        cxnSp.x = newX;
        cxnSp.cx = (dx - cx);
      }
    } else {
      dx = Math.abs((newX - x));
      if (newX <= x) {
        cxnSp.x = newX;
        cxnSp.cx = (cx + dx);
      } else if (dx <= cx) {
        cxnSp.x = newX;
        cxnSp.cx = (cx - dx);
      } else {
        cxnSp.flipH = true;
        cxnSp.x = (x + cx);
        cxnSp.cx = (dx - cx);
      }
    }
  }

  /**
   * Return the Y-position of the begin point of this connector, in
   * English Metric Units (as a |Length| object).
   */
  get beginY() {
    const cxnSp = this._element;
    const [y, cy, flipV] = [cxnSp.y, cxnSp.cy, cxnSp.flipV];
    const beginY = (flipV ? (y + cy) : y);
    return Emu(beginY);
  }

  set beginY(value: any) {
    let dy;
    const cxnSp = this._element;
    const [y, cy, flipV, newY] = [cxnSp.y, cxnSp.cy, cxnSp.flipV, pyInt(value)];
    if (flipV) {
      const oldY = (y + cy);
      dy = Math.abs((newY - oldY));
      if (newY >= oldY) {
        cxnSp.cy = (cy + dy);
      } else if (dy <= cy) {
        cxnSp.cy = (cy - dy);
      } else {
        cxnSp.flipV = false;
        cxnSp.y = newY;
        cxnSp.cy = (dy - cy);
      }
    } else {
      dy = Math.abs((newY - y));
      if (newY <= y) {
        cxnSp.y = newY;
        cxnSp.cy = (cy + dy);
      } else if (dy <= cy) {
        cxnSp.y = newY;
        cxnSp.cy = (cy - dy);
      } else {
        cxnSp.flipV = true;
        cxnSp.y = (y + cy);
        cxnSp.cy = (dy - cy);
      }
    }
  }

  /**
   * **EXPERIMENTAL** - *The current implementation only works properly
   * with rectangular shapes, such as pictures and rectangles. Use with
   * other shape types may cause unexpected visual alignment of the
   * connected end-point and could lead to a load error if cxn_pt_idx
   * exceeds the connection point count available on the connected shape.
   * That said, a quick test should reveal what to expect when using this
   * method with other shape types.*
   *
   * Connect the ending of this connector to *shape* at the connection
   * point specified by *cxn_pt_idx*.
   */
  endConnect(shape: any, cxnPtIdx: any) {
    this._connectEndTo(shape, cxnPtIdx);
    this._moveEndToCxn(shape, cxnPtIdx);
  }

  /**
   * Return the X-position of the end point of this connector, in English
   * Metric Units (as a |Length| object).
   */
  get endX() {
    const cxnSp = this._element;
    const [x, cx, flipH] = [cxnSp.x, cxnSp.cx, cxnSp.flipH];
    const endX = (flipH ? x : (x + cx));
    return Emu(endX);
  }

  set endX(value: any) {
    let dx;
    const cxnSp = this._element;
    const [x, cx, flipH, newX] = [cxnSp.x, cxnSp.cx, cxnSp.flipH, pyInt(value)];
    if (flipH) {
      dx = Math.abs((newX - x));
      if (newX <= x) {
        cxnSp.x = newX;
        cxnSp.cx = (cx + dx);
      } else if (dx <= cx) {
        cxnSp.x = newX;
        cxnSp.cx = (cx - dx);
      } else {
        cxnSp.flipH = false;
        cxnSp.x = (x + cx);
        cxnSp.cx = (dx - cx);
      }
    } else {
      const oldX = (x + cx);
      dx = Math.abs((newX - oldX));
      if (newX >= oldX) {
        cxnSp.cx = (cx + dx);
      } else if (dx <= cx) {
        cxnSp.cx = (cx - dx);
      } else {
        cxnSp.flipH = true;
        cxnSp.x = newX;
        cxnSp.cx = (dx - cx);
      }
    }
  }

  /**
   * Return the Y-position of the end point of this connector, in English
   * Metric Units (as a |Length| object).
   */
  get endY() {
    const cxnSp = this._element;
    const [y, cy, flipV] = [cxnSp.y, cxnSp.cy, cxnSp.flipV];
    const endY = (flipV ? y : (y + cy));
    return Emu(endY);
  }

  set endY(value: any) {
    let dy;
    const cxnSp = this._element;
    const [y, cy, flipV, newY] = [cxnSp.y, cxnSp.cy, cxnSp.flipV, pyInt(value)];
    if (flipV) {
      dy = Math.abs((newY - y));
      if (newY <= y) {
        cxnSp.y = newY;
        cxnSp.cy = (cy + dy);
      } else if (dy <= cy) {
        cxnSp.y = newY;
        cxnSp.cy = (cy - dy);
      } else {
        cxnSp.flipV = false;
        cxnSp.y = (y + cy);
        cxnSp.cy = (dy - cy);
      }
    } else {
      const oldY = (y + cy);
      dy = Math.abs((newY - oldY));
      if (newY >= oldY) {
        cxnSp.cy = (cy + dy);
      } else if (dy <= cy) {
        cxnSp.cy = (cy - dy);
      } else {
        cxnSp.flipV = true;
        cxnSp.y = newY;
        cxnSp.cy = (dy - cy);
      }
    }
  }

  /** Helper method required by |LineFormat|. */
  getOrAddLn() {
    return this._element.spPr.getOrAddLn();
  }

  /**
   * |LineFormat| instance for this connector.
   *
   * Provides access to line properties such as line color, width, and
   * line style.
   */
  get line() {
    return lazy(this, "line", () => {
      return new LineFormat(this);
    });
  }

  /**
   * Helper method required by |LineFormat|.
   *
   * The ``<a:ln>`` element containing the line format properties such as
   * line color and width. |None| if no `<a:ln>` element is present.
   */
  get ln() {
    return this._element.spPr.ln;
  }

  /**
   * Member of `MSO_SHAPE_TYPE` identifying the type of this shape.
   *
   * Unconditionally `MSO_SHAPE_TYPE.LINE` for a `Connector` object.
   */
  get shapeType() {
    return MSO_SHAPE_TYPE.LINE;
  }

  /**
   * Add or update a stCxn element for this connector that connects its
   * begin point to the connection point of *shape* specified by
   * *cxn_pt_idx*.
   */
  _connectBeginTo(shape: any, cxnPtIdx: any) {
    const cNvCxnSpPr = this._element.nvCxnSpPr.cNvCxnSpPr;
    const stCxn = cNvCxnSpPr.getOrAddStCxn();
    stCxn.id = shape.shapeId;
    stCxn.idx = cxnPtIdx;
  }

  /**
   * Add or update an endCxn element for this connector that connects its
   * end point to the connection point of *shape* specified by
   * *cxn_pt_idx*.
   */
  _connectEndTo(shape: any, cxnPtIdx: any) {
    const cNvCxnSpPr = this._element.nvCxnSpPr.cNvCxnSpPr;
    const endCxn = cNvCxnSpPr.getOrAddEndCxn();
    endCxn.id = shape.shapeId;
    endCxn.idx = cxnPtIdx;
  }

  /**
   * Move the begin point of this connector to coordinates of the
   * connection point of *shape* specified by *cxn_pt_idx*.
   */
  _moveBeginToCxn(shape: any, cxnPtIdx: any) {
    const [x, y, cx, cy] = [shape.left, shape.top, shape.width, shape.height];
    [this.beginX, this.beginY] = cxnPoint(x, y, cx, cy, cxnPtIdx);
  }

  /**
   * Move the end point of this connector to the coordinates of the
   * connection point of *shape* specified by *cxn_pt_idx*.
   */
  _moveEndToCxn(shape: any, cxnPtIdx: any) {
    const [x, y, cx, cy] = [shape.left, shape.top, shape.width, shape.height];
    [this.endX, this.endY] = cxnPoint(x, y, cx, cy, cxnPtIdx);
  }
}

/** Slide coordinates of connection point `cxnPtIdx` (0 top, 1 left, 2 bottom, 3 right) of a shape. */
function cxnPoint(x: number, y: number, cx: number, cy: number, cxnPtIdx: number): [number, number] {
  switch (cxnPtIdx) {
    case 0:
      return [Math.trunc(x + cx / 2), y];
    case 1:
      return [x, Math.trunc(y + cy / 2)];
    case 2:
      return [Math.trunc(x + cx / 2), y + cy];
    case 3:
      return [x + cx, Math.trunc(y + cy / 2)];
    default:
      throw new KeyError(String(cxnPtIdx));
  }
}
