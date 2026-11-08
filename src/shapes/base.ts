// Base shape-related objects such as BaseShape. (`pptx.shapes.base`)
import { ValueError, NotImplementedError } from "../errors.js";
import { lazy } from "../util.js";
import { ActionSetting } from "../action.js";
import { ShadowFormat } from "../dml/effect.js";
import { ElementProxy } from "../shared.js";
import type { MSO_SHAPE_TYPE, PP_PLACEHOLDER } from "../enum/shapes.js";
import type { ShapeElement } from "../oxml/shapes/index.js";
import type { CT_Placeholder } from "../oxml/shapes/shared.js";
import type { BaseSlidePart } from "../parts/slide.js";
import type { ProvidesPart } from "../types.js";
import type { Length } from "../util.js";

/**
 * Base class for shape objects.
 *
 * Subclasses include |Shape|, |Picture|, and |GraphicFrame|.
 */
export class BaseShape {
  declare _element: any;
  declare _parent: any;

  constructor(shapeElm: ShapeElement, parent: ProvidesPart) {
    this._element = shapeElm;
    this._parent = parent;
  }

  /**
   * |True| if this shape object proxies the same element as *other*.
   *
   * Equality for proxy objects is defined as referring to the same XML element, whether or not
   * they are the same proxy object instance.
   */
  equals(other: unknown): boolean {
    if (!(other instanceof BaseShape)) {
      return false;
    }
    return this._element === other._element;
  }


  /**
   * |ActionSetting| instance providing access to click behaviors.
   *
   * Click behaviors are hyperlink-like behaviors including jumping to a hyperlink (web page)
   * or to another slide in the presentation. The click action is that defined on the overall
   * shape, not a run of text within the shape. An |ActionSetting| object is always returned,
   * even when no click behavior is defined on the shape.
   */
  get clickAction(): ActionSetting {
    return lazy(this, "clickAction", () => {
      const cNvPr = this._element._nvXxPr.cNvPr;
      return new ActionSetting(cNvPr, this);
    });
  }

  /**
   * `lxml` element for this shape, e.g. a CT_Shape instance.
   *
   * Note that manipulating this element improperly can produce an invalid presentation file.
   * Make sure you know what you're doing if you use this to change the underlying XML.
   */
  get element(): ShapeElement {
    return this._element;
  }

  /**
   * |True| if this shape is a graphic frame containing a chart object.
   *
   * |False| otherwise. When |True|, the chart object can be accessed using the ``.chart``
   * property.
   */
  get hasChart(): boolean {
    return false;
  }

  /**
   * |True| if this shape is a graphic frame containing a table object.
   *
   * |False| otherwise. When |True|, the table object can be accessed using the ``.table``
   * property.
   */
  get hasTable(): boolean {
    return false;
  }

  /** |True| if this shape can contain text. */
  get hasTextFrame(): boolean {
    return false;
  }

  /** Read/write. Integer distance between top and bottom extents of shape in EMUs. */
  get height(): number {
    return this._element.cy;
  }

  set height(value: number) {
    this._element.cy = value;
  }

  /**
   * True if this shape is a placeholder.
   *
   * A shape is a placeholder if it has a <p:ph> element.
   */
  get isPlaceholder(): boolean {
    return this._element.hasPhElm;
  }

  /**
   * Integer distance of the left edge of this shape from the left edge of the slide.
   *
   * Read/write. Expressed in English Metric Units (EMU)
   */
  get left(): number {
    return this._element.x;
  }

  set left(value: number) {
    this._element.x = value;
  }

  /** Name of this shape, e.g. 'Picture 7'. */
  get name(): string {
    return this._element.shapeName;
  }

  set name(value: string) {
    this._element._nvXxPr.cNvPr.name = value;
  }

  /**
   * The package part containing this shape.
   *
   * A |BaseSlidePart| subclass in this case. Access to a slide part should only be required if
   * you are extending the behavior of |pp| API objects.
   */
  get part(): BaseSlidePart {
    return this._parent.part;
  }

  /**
   * Provides access to placeholder-specific properties such as placeholder type.
   *
   * Raises |ValueError| on access if the shape is not a placeholder.
   */
  get placeholderFormat(): _PlaceholderFormat {
    const ph = this._element.ph;
    if (ph === null) {
      throw new ValueError("shape is not a placeholder");
    }
    return new _PlaceholderFormat(ph);
  }

  /**
   * Degrees of clockwise rotation.
   *
   * Read/write float. Negative values can be assigned to indicate counter-clockwise rotation,
   * e.g. assigning -45.0 will change setting to 315.0.
   */
  get rotation(): number {
    return this._element.rot;
  }

  set rotation(value: number) {
    this._element.rot = value;
  }

  /**
   * |ShadowFormat| object providing access to shadow for this shape.
   *
   * A |ShadowFormat| object is always returned, even when no shadow is
   * explicitly defined on this shape (i.e. it inherits its shadow
   * behavior).
   */
  get shadow(): ShadowFormat {
    return lazy(this, "shadow", () => {
      return new ShadowFormat(this._element.spPr);
    });
  }

  /**
   * Read-only positive integer identifying this shape.
   *
   * The id of a shape is unique among all shapes on a slide.
   */
  get shapeId(): number {
    return this._element.shapeId;
  }

  /**
   * A member of MSO_SHAPE_TYPE classifying this shape by type.
   *
   * Like ``MSO_SHAPE_TYPE.CHART``. Must be implemented by subclasses.
   */
  get shapeType(): MSO_SHAPE_TYPE | null {
    throw new NotImplementedError(`${this.constructor.name} does not implement \`.shape_type\``);
  }

  /**
   * Distance from the top edge of the slide to the top edge of this shape.
   *
   * Read/write. Expressed in English Metric Units (EMU)
   */
  get top(): number {
    return this._element.y;
  }

  set top(value: number) {
    this._element.y = value;
  }

  /**
   * Distance between left and right extents of this shape.
   *
   * Read/write. Expressed in English Metric Units (EMU).
   */
  get width(): number {
    return this._element.cx;
  }

  set width(value: number) {
    this._element.cx = value;
  }
}

/**
 * Provides properties specific to placeholders, such as the placeholder type.
 *
 * Accessed via the :attr:`~.BaseShape.placeholder_format` property of a placeholder shape,
 */
export class _PlaceholderFormat extends ElementProxy {
  declare _ph: any;

  constructor(element: CT_Placeholder) {
    super(element);
    this._ph = element;
  }

  /** The `p:ph` element proxied by this object. */
  get element(): CT_Placeholder {
    return this._ph;
  }

  /** Integer placeholder 'idx' attribute. */
  get idx(): number {
    return this._ph.idx;
  }

  /**
   * Placeholder type.
   *
   * A member of the :ref:`PpPlaceholderType` enumeration, e.g. PP_PLACEHOLDER.CHART
   */
  get type(): PP_PLACEHOLDER {
    return this._ph.type;
  }
}
