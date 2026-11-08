// Autoshape-related objects such as Shape and Adjustment. (`pptx.shapes.autoshape`)
import { IndexError, KeyError, NotImplementedError, ValueError } from "../errors.js";
import { PY_GETITEM } from "../oxml/xml.js";
import { lazy, pyInt, pyRepr, xmlEscape } from "../util.js";
import { FillFormat } from "../dml/fill.js";
import { LineFormat } from "../dml/line.js";
import { MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE } from "../enum/shapes.js";
import { BaseShape } from "./base.js";
import { autoshapeTypes } from "../spec.js";
import { TextFrame } from "../text/text.js";
import type { CT_GeomGuide, CT_PresetGeometry2D, CT_Shape } from "../oxml/shapes/autoshape.js";
import type { AdjustmentValue } from "../spec.js";
import type { ProvidesPart } from "../types.js";

/**
 * An adjustment value for an autoshape.
 *
 * An adjustment value corresponds to the position of an adjustment handle on an auto shape.
 * Adjustment handles are the small yellow diamond-shaped handles that appear on certain auto
 * shapes and allow the outline of the shape to be adjusted. For example, a rounded rectangle has
 * an adjustment handle that allows the radius of its corner rounding to be adjusted.
 *
 * Values are |float| and generally range from 0.0 to 1.0, although the value can be negative or
 * greater than 1.0 in certain circumstances.
 */
export class Adjustment {
  declare name: any;
  declare defVal: any;
  declare actual: any;

  constructor(name: string, defVal: number, actual: number | null = null) {
    this.name = name;
    this.defVal = defVal;
    this.actual = actual;
  }

  /**
   * Read/write |float| representing normalized adjustment value for this adjustment.
   *
   * Actual values are a large-ish integer expressed in shape coordinates, nominally between 0
   * and 100,000. The effective value is normalized to a corresponding value nominally between
   * 0.0 and 1.0. Intuitively this represents the proportion of the width or height of the shape
   * at which the adjustment value is located from its starting point. For simple shapes such as
   * a rounded rectangle, this intuitive correspondence holds. For more complicated shapes and
   * at more extreme shape proportions (e.g. width is much greater than height), the value can
   * become negative or greater than 1.0.
   */
  get effectiveValue(): number {
    const rawValue = (this.actual !== null ? this.actual : this.defVal);
    return (this.constructor as typeof Adjustment)._normalize(rawValue);
  }

  set effectiveValue(value: number) {
    if (!(typeof value === "number")) {
      throw new ValueError(`adjustment value must be numeric, got ${pyRepr(value)}`);
    }
    this.actual = (this.constructor as typeof Adjustment)._denormalize(value);
  }

  /**
   * Return integer corresponding to normalized `raw_value` on unit basis of 100,000.
   *
   * See Adjustment.normalize for additional details.
   */
  static _denormalize(value: number): number {
    return pyInt((value * 100000.0));
  }

  /**
   * Return normalized value for `raw_value`.
   *
   * A normalized value is a |float| between 0.0 and 1.0 for nominal raw values between 0 and
   * 100,000. Raw values less than 0 and greater than 100,000 are valid and return values
   * calculated on the same unit basis of 100,000.
   */
  static _normalize(rawValue: number): number {
    return (rawValue / 100000.0);
  }

  /**
   * Denormalized effective value.
   *
   * Expressed in shape coordinates, this is suitable for using in the XML.
   */
  get val(): number {
    return (this.actual !== null ? this.actual : this.defVal);
  }
}

/**
 * Sequence of |Adjustment| instances for an auto shape.
 *
 * Each represents an available adjustment for a shape of its type. Supports `len()` and indexed
 * access, e.g. `shape.adjustments[1] = 0.15`.
 */
export class AdjustmentCollection {
  declare _adjustments_: any;
  declare _prstGeom: any;

  constructor(prstGeom: CT_PresetGeometry2D) {
    this._adjustments_ = this._initializedAdjustments(prstGeom);
    this._prstGeom = prstGeom;
  }

  /** Provides indexed access, (e.g. 'adjustments[9]'). */
  at(idx: number): number {
    return this._adjustment(idx).effectiveValue;
  }

  [PY_GETITEM](idx: number): number {
    return this.at(idx);
  }

  /** The Adjustment at `idx` (python list indexing: negative counts from the end). */
  _adjustment(idx: number): Adjustment {
    if (typeof idx !== "number" || !Number.isInteger(idx)) throw new TypeError(`list indices must be integers or slices, not ${typeof idx}`);
    const adj = this._adjustments_.at(idx);
    if (adj === undefined) throw new IndexError("list index out of range");
    return adj;
  }

  /**
   * Provides item assignment via an indexed expression, e.g. `adjustments[9] = 999.9`.
   *
   * Causes all adjustment values in collection to be written to the XML.
   */
  set(idx: number, value: number): void {
    this._adjustment(idx).effectiveValue = value;
    this._rewriteGuides();
  }

  /** Return an initialized list of adjustment values based on the contents of `prstGeom`. */
  _initializedAdjustments(prstGeom: CT_PresetGeometry2D | null): Adjustment[] {
    if (prstGeom === null) {
      return [];
    }
    const davs = AutoShapeType.defaultAdjustmentValues(prstGeom.prst);
    const adjustments = [...davs].map(([name, defVal]) => new Adjustment(name, defVal));
    (this.constructor as typeof AdjustmentCollection)._updateAdjustmentsWithActuals(adjustments, prstGeom.gdLst);
    return adjustments;
  }

  /**
   * Write `a:gd` elements to the XML, one for each adjustment value.
   *
   * Any existing guide elements are overwritten.
   */
  _rewriteGuides() {
    const guides = [...this._adjustments_].map((adj) => [adj.name, adj.val]);
    this._prstGeom.rewriteGuides(guides);
  }

  /**
   * Update |Adjustment| instances in `adjustments` with actual values held in `guides`.
   *
   * `guides` is a list of `a:gd` elements. Guides with a name that does not match an adjustment
   * object are skipped.
   */
  static _updateAdjustmentsWithActuals(adjustments: Adjustment[], guides: CT_GeomGuide[]): void {
    const adjustmentsByName = new Map(adjustments.map((adj) => [adj.name, adj]));
    for (const gd of guides) {
      const name = gd.name;
      const actual = pyInt(gd.fmla.slice(4));
      // guides with a name that does not match an adjustment are skipped
      const adjustment = adjustmentsByName.get(name);
      if (adjustment === undefined) continue;
      adjustment.actual = actual;
    }
  }

  /** Sequence of |Adjustment| objects contained in collection. */
  get _adjustments(): Adjustment[] {
    return [...this._adjustments_];
  }

  /** Implement built-in function len() */
  get length(): number {
    return this._adjustments_.length;
  }
}

/**
 * Provides access to metadata for an auto-shape of type identified by `autoshape_type_id`.
 *
 * Instances are cached, so no more than one instance for a particular auto shape type is in
 * memory.
 *
 * Instances provide the following attributes:
 *
 * .. attribute:: autoshape_type_id
 *
 *    Integer uniquely identifying this auto shape type. Corresponds to a
 *    value in `pptx.constants.MSO` like `MSO_SHAPE.ROUNDED_RECTANGLE`.
 *
 * .. attribute:: basename
 *
 *    Base part of shape name for auto shapes of this type, e.g. `Rounded
 *    Rectangle` becomes `Rounded Rectangle 99` when the distinguishing
 *    integer is added to the shape name.
 *
 * .. attribute:: prst
 *
 *    String identifier for this auto shape type used in the `a:prstGeom`
 *    element.
 */
export class AutoShapeType {
  declare _autoshapeTypeId: MSO_AUTO_SHAPE_TYPE;
  declare _basename: string;
  declare _loaded: boolean;
  /** One instance per autoshape type (python `__new__` caches them the same way). */
  static _instances = new Map<number, AutoShapeType>();

  /** Initialize attributes from constant values in `pptx.spec`; returns the cached instance after the first call. */
  constructor(autoshapeTypeId: MSO_AUTO_SHAPE_TYPE) {
    const key = Number(autoshapeTypeId);
    const cached = AutoShapeType._instances.get(key);
    if (cached !== undefined) return cached;
    const autoshapeType = autoshapeTypes.get(key);
    if (autoshapeType === undefined) {
      throw new KeyError(`no autoshape type with id '${autoshapeTypeId}' in pptx.spec.autoshape_types`);
    }
    this._autoshapeTypeId = autoshapeTypeId;
    this._basename = autoshapeType.basename;
    this._loaded = true;
    AutoShapeType._instances.set(key, this);
  }

  /** MSO_AUTO_SHAPE_TYPE enumeration member identifying this auto shape type. */
  get autoshapeTypeId(): MSO_AUTO_SHAPE_TYPE {
    return this._autoshapeTypeId;
  }

  /**
   * Base of shape name for this auto shape type.
   *
   * A shape name is like "Rounded Rectangle 7" and appears as an XML attribute for example at
   * `p:sp/p:nvSpPr/p:cNvPr{name}`. This basename value is the name less the distinguishing
   * integer. This value is escaped because at least one autoshape-type name includes double
   * quotes ('"No" Symbol').
   */
  get basename(): string {
    return xmlEscape(this._basename, { '"': "&quot;" });
  }

  /** Sequence of (name, value) pair adjustment value defaults for `prst` autoshape-type. */
  static defaultAdjustmentValues(prst: MSO_AUTO_SHAPE_TYPE): readonly AdjustmentValue[] {
    const spec = autoshapeTypes.get(Number(prst));
    if (spec === undefined) throw new KeyError(String(prst));
    return spec.avLst;
  }

  /**
   * Select auto shape type with matching `prst`.
   *
   * e.g. `MSO_SHAPE.RECTANGLE` corresponding to preset geometry keyword `"rect"`.
   */
  static idFromPrst(prst: string): MSO_AUTO_SHAPE_TYPE {
    return MSO_AUTO_SHAPE_TYPE.fromXml(prst);
  }

  /**
   * Preset geometry identifier string for this auto shape. Used in the
   * `prst` attribute of `a:prstGeom` element to specify the geometry
   * to be used in rendering the shape, for example `'roundRect'`.
   */
  get prst() {
    return MSO_AUTO_SHAPE_TYPE.toXml(this._autoshapeTypeId);
  }
}

/**
 * A shape that can appear on a slide.
 *
 * Corresponds to the `p:sp` element that can appear in any of the slide-type parts
 * (slide, slideLayout, slideMaster, notesPage, notesMaster, handoutMaster).
 */
export class Shape extends BaseShape {
  declare _sp: any;

  constructor(sp: CT_Shape, parent: ProvidesPart) {
    super(sp, parent);
    this._sp = sp;
  }

  /** Read-only reference to |AdjustmentCollection| instance for this shape. */
  get adjustments(): AdjustmentCollection {
    return lazy(this, "adjustments", () => {
      return new AdjustmentCollection(this._sp.prstGeom);
    });
  }

  /**
   * Enumeration value identifying the type of this auto shape.
   *
   * Like `MSO_SHAPE.ROUNDED_RECTANGLE`. Raises |ValueError| if this shape is not an auto shape.
   */
  get autoShapeType() {
    if (!this._sp.isAutoshape) {
      throw new ValueError("shape is not an auto shape");
    }
    return this._sp.prst;
  }

  /**
   * |FillFormat| instance for this shape.
   *
   * Provides access to fill properties such as fill color.
   */
  get fill() {
    return lazy(this, "fill", () => {
      return FillFormat.fromFillParent(this._sp.spPr);
    });
  }

  /** Return the `a:ln` element containing the line format properties XML for this shape. */
  getOrAddLn() {
    return this._sp.getOrAddLn();
  }

  /** |True| if this shape can contain text. Always |True| for an AutoShape. */
  get hasTextFrame(): boolean {
    return true;
  }

  /**
   * |LineFormat| instance for this shape.
   *
   * Provides access to line properties such as line color.
   */
  get line() {
    return lazy(this, "line", () => {
      return new LineFormat(this);
    });
  }

  /**
   * The `a:ln` element containing the line format properties such as line color and width.
   *
   * |None| if no `a:ln` element is present.
   */
  get ln() {
    return this._sp.ln;
  }

  /** Unique integer identifying the type of this shape, like `MSO_SHAPE_TYPE.TEXT_BOX`. */
  get shapeType(): MSO_SHAPE_TYPE {
    if (this.isPlaceholder) {
      return MSO_SHAPE_TYPE.PLACEHOLDER;
    }
    if (this._sp.hasCustomGeometry) {
      return MSO_SHAPE_TYPE.FREEFORM;
    }
    if (this._sp.isAutoshape) {
      return MSO_SHAPE_TYPE.AUTO_SHAPE;
    }
    if (this._sp.isTextbox) {
      return MSO_SHAPE_TYPE.TEXT_BOX;
    }
    throw new NotImplementedError("Shape instance of unrecognized shape type");
  }

  /**
   * Read/write. Text in shape as a single string.
   *
   * The returned string will contain a newline character (`"\n"`) separating each paragraph
   * and a vertical-tab (`"\v"`) character for each line break (soft carriage return) in the
   * shape's text.
   *
   * Assignment to `text` replaces any text previously contained in the shape, along with any
   * paragraph or font formatting applied to it. A newline character (`"\n"`) in the assigned
   * text causes a new paragraph to be started. A vertical-tab (`"\v"`) character in the
   * assigned text causes a line-break (soft carriage-return) to be inserted. (The vertical-tab
   * character appears in clipboard text copied from PowerPoint as its str encoding of
   * line-breaks.)
   */
  get text(): string {
    return this.textFrame.text;
  }

  set text(text: string) {
    this.textFrame.text = text;
  }

  /**
   * |TextFrame| instance for this shape.
   *
   * Contains the text of the shape and provides access to text formatting properties.
   */
  get textFrame() {
    const txBody = this._sp.getOrAddTxBody();
    return new TextFrame(txBody, this);
  }
}
