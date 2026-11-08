// DrawingML objects related to fill (`pptx.dml.fill`).
import { MSO_FILL, type MSO_FILL_TYPE, type MSO_PATTERN_TYPE } from "../enum/dml.js";
import { NotImplementedError, ValueError } from "../errors.js";
import {
  CT_BlipFillProperties,
  CT_GradientFillProperties,
  CT_GroupFillProperties,
  CT_NoFillProperties,
  CT_PatternFillProperties,
  CT_SolidColorFillProperties,
  type CT_GradientStop,
  type CT_GradientStopList,
} from "../oxml/dml/fill.js";
import { PY_GETITEM } from "../oxml/xml.js";
import type { BaseOxmlElement } from "../oxml/xmlchemy.js";
import { ElementProxy } from "../shared.js";
import { lazy } from "../util.js";
import { ColorFormat } from "./color.js";

export { CT_GradientFillProperties };

/**
 * Provides access to the current fill properties object and provides methods to change the
 * fill type.
 */
export class FillFormat {
  declare _xPr: any;
  declare _fill: _Fill;

  constructor(egFillPropertiesParent: BaseOxmlElement, fillObj: _Fill) {
    this._xPr = egFillPropertiesParent;
    this._fill = fillObj;
  }

  /**
   * A `FillFormat` instance initialized to the settings contained in
   * `egFillPropertiesParent`, which must be an element having EG_FillProperties in its child
   * element sequence in the XML schema.
   */
  static fromFillParent(egFillPropertiesParent: BaseOxmlElement): FillFormat {
    const fillElm = (egFillPropertiesParent as any).egFillProperties;
    const fill = new _Fill(fillElm);
    return new this(egFillPropertiesParent, fill);
  }

  /**
   * `ColorFormat` object representing background color.
   *
   * This property is only applicable to pattern fills and lines.
   */
  get backColor(): ColorFormat {
    return this._fill.backColor;
  }

  /** Sets the fill type to noFill, i.e. transparent. */
  background(): void {
    const noFill = this._xPr.getOrChangeToNoFill();
    this._fill = new _NoFill(noFill);
  }

  /** `ColorFormat` instance representing the foreground color of this fill. */
  get foreColor(): ColorFormat {
    return this._fill.foreColor;
  }

  /**
   * Sets the fill type to gradient.
   *
   * If the fill is not already a gradient, a default gradient is added. The default gradient
   * corresponds to the default in the built-in PowerPoint "White" template. This gradient is
   * linear at angle 90-degrees (upward), with two stops. The first stop is Accent-1 with tint
   * 100%, shade 100%, and satMod 130%. The second stop is Accent-1 with tint 50%, shade 100%,
   * and satMod 350%.
   */
  gradient(): void {
    const gradFill = this._xPr.getOrChangeToGradFill();
    this._fill = new _GradFill(gradFill);
  }

  /**
   * Angle in float degrees of line of a linear gradient.
   *
   * Read/Write. May be null, indicating the angle should be inherited from the style hierarchy.
   * An angle of 0.0 corresponds to a left-to-right gradient. Increasing angles represent
   * counter-clockwise rotation of the line, for example 90.0 represents a bottom-to-top
   * gradient. Raises TypeError when the fill type is not MSO_FILL_TYPE.GRADIENT. Raises
   * ValueError for a non-linear gradient (e.g. a radial gradient).
   */
  get gradientAngle(): number | null {
    if (this.type !== MSO_FILL.GRADIENT) throw new TypeError("Fill is not of type MSO_FILL_TYPE.GRADIENT");
    return (this._fill as _GradFill).gradientAngle;
  }
  set gradientAngle(value: number) {
    if (this.type !== MSO_FILL.GRADIENT) throw new TypeError("Fill is not of type MSO_FILL_TYPE.GRADIENT");
    (this._fill as _GradFill).gradientAngle = value;
  }

  /**
   * `_GradientStops` object providing access to stops of this gradient.
   *
   * Raises TypeError when fill is not gradient (call `fill.gradient()` first). Each stop
   * represents a color between which the gradient smoothly transitions.
   */
  get gradientStops(): _GradientStops {
    if (this.type !== MSO_FILL.GRADIENT) throw new TypeError("Fill is not of type MSO_FILL_TYPE.GRADIENT");
    return (this._fill as _GradFill).gradientStops;
  }

  /**
   * Return member of `MSO_PATTERN_TYPE` indicating fill pattern.
   *
   * Raises TypeError when fill is not patterned (call `fill.patterned()` first). Returns null
   * if no pattern has been set; PowerPoint may display the default `PERCENT_5` pattern in this
   * case. Assigning null will remove any explicit pattern setting, although relying on the
   * default behavior is discouraged and may produce rendering differences across client
   * applications.
   */
  get pattern(): MSO_PATTERN_TYPE | null {
    return this._fill.pattern;
  }
  set pattern(patternType: MSO_PATTERN_TYPE | null) {
    this._fill.pattern = patternType;
  }

  /**
   * Selects the pattern fill type.
   *
   * Note that calling this method does not by itself set a foreground or background color of
   * the pattern. Rather it enables subsequent assignments to properties like foreColor to set
   * the pattern and colors.
   */
  patterned(): void {
    const pattFill = this._xPr.getOrChangeToPattFill();
    this._fill = new _PattFill(pattFill);
  }

  /**
   * Sets the fill type to solid, i.e. a solid color. Note that calling this method does not set
   * a color or by itself cause the shape to appear with a solid color fill; rather it enables
   * subsequent assignments to properties like foreColor to set the color.
   */
  solid(): void {
    const solidFill = this._xPr.getOrChangeToSolidFill();
    this._fill = new _SolidFill(solidFill);
  }

  /** The type of this fill, e.g. `MSO_FILL_TYPE.SOLID`. */
  get type(): MSO_FILL_TYPE | null {
    return this._fill.type;
  }
}

/**
 * Object factory for fill object of class matching fill element, such as `_SolidFill` for
 * `<a:solidFill>`; also serves as the base class for all fill classes. `new _Fill(xFill)`
 * returns an instance of the matching subclass (python's `_Fill.__new__`).
 */
export class _Fill {
  constructor(xFill?: any) {
    if (new.target === _Fill) {
      const fillCls = fillClsFor(xFill);
      if (fillCls !== _Fill) return new fillCls(xFill);
    }
  }

  get backColor(): ColorFormat {
    throw new TypeError(`fill type ${this.constructor.name} has no background color, call .patterned() first`);
  }

  get foreColor(): ColorFormat {
    throw new TypeError(`fill type ${this.constructor.name} has no foreground color, call .solid() or .patterned() first`);
  }

  get pattern(): MSO_PATTERN_TYPE | null {
    throw new TypeError(`fill type ${this.constructor.name} has no pattern, call .patterned() first`);
  }
  set pattern(_patternType: MSO_PATTERN_TYPE | null) {
    throw new TypeError(`fill type ${this.constructor.name} has no pattern, call .patterned() first`);
  }

  get type(): MSO_FILL_TYPE | null {
    throw new NotImplementedError(`.type property must be implemented on ${this.constructor.name}`);
  }
}

function fillClsFor(xFill: unknown): new (xFill: any) => _Fill {
  if (xFill === null || xFill === undefined) return _NoneFill;
  if (xFill instanceof CT_BlipFillProperties) return _BlipFill;
  if (xFill instanceof CT_GradientFillProperties) return _GradFill;
  if (xFill instanceof CT_GroupFillProperties) return _GrpFill;
  if (xFill instanceof CT_NoFillProperties) return _NoFill;
  if (xFill instanceof CT_PatternFillProperties) return _PattFill;
  if (xFill instanceof CT_SolidColorFillProperties) return _SolidFill;
  return _Fill;
}

export class _BlipFill extends _Fill {
  get type(): MSO_FILL_TYPE {
    return MSO_FILL.PICTURE;
  }
}

/** Proxies an `a:gradFill` element. */
export class _GradFill extends _Fill {
  declare _element: CT_GradientFillProperties;
  declare _gradFill: CT_GradientFillProperties;

  constructor(gradFill: CT_GradientFillProperties) {
    super();
    this._element = this._gradFill = gradFill;
  }

  /**
   * Angle in float degrees of line of a linear gradient.
   *
   * Read/Write. May be null, indicating the angle is inherited from the style hierarchy. An
   * angle of 0.0 corresponds to a left-to-right gradient. Increasing angles represent
   * counter-clockwise rotation of the line, for example 90.0 represents a bottom-to-top
   * gradient. Raises ValueError for a non-linear gradient (e.g. a radial gradient).
   */
  get gradientAngle(): number | null {
    // case 1: gradient path is explicit, but not linear
    const path = this._gradFill.path;
    if (path !== null) throw new ValueError("not a linear gradient");
    // case 2: gradient path is inherited (no a:lin OR a:path)
    const lin = this._gradFill.lin;
    if (lin === null) return null;
    // case 3: gradient path is explicitly linear. The angle is stored in XML as a clockwise
    // angle, whereas the UI reports it as counter-clockwise from horizontal-pointing-right.
    // Since the UI is consistent with trigonometry conventions, we respect that in the API.
    const clockwiseAngle = lin.ang;
    return clockwiseAngle === 0.0 ? 0.0 : 360.0 - clockwiseAngle;
  }
  set gradientAngle(value: number) {
    const lin = this._gradFill.lin;
    if (lin === null) throw new ValueError("not a linear gradient");
    lin.ang = 360.0 - value;
  }

  /**
   * `_GradientStops` object providing access to gradient colors.
   *
   * Each stop represents a color between which the gradient smoothly transitions.
   */
  get gradientStops(): _GradientStops {
    return lazy(this, "gradientStops", () => new _GradientStops(this._gradFill.getOrAddGsLst()));
  }

  get type(): MSO_FILL_TYPE {
    return MSO_FILL.GRADIENT;
  }
}

export class _GrpFill extends _Fill {
  get type(): MSO_FILL_TYPE {
    return MSO_FILL.GROUP;
  }
}

export class _NoFill extends _Fill {
  get type(): MSO_FILL_TYPE {
    return MSO_FILL.BACKGROUND;
  }
}

export class _NoneFill extends _Fill {
  get type(): null {
    return null;
  }
}

/** Provides access to patterned fill properties. */
export class _PattFill extends _Fill {
  declare _element: CT_PatternFillProperties;
  declare _pattFill: CT_PatternFillProperties;

  constructor(pattFill: CT_PatternFillProperties) {
    super();
    this._element = this._pattFill = pattFill;
  }

  /** `ColorFormat` object providing access to background color. */
  get backColor(): ColorFormat {
    return lazy(this, "backColor", () => ColorFormat.fromColorchoiceParent(this._pattFill.getOrAddBgClr()));
  }

  /** `ColorFormat` object providing access to foreground color. */
  get foreColor(): ColorFormat {
    return lazy(this, "foreColor", () => ColorFormat.fromColorchoiceParent(this._pattFill.getOrAddFgClr()));
  }

  /**
   * Return member of `MSO_PATTERN_TYPE` indicating fill pattern.
   *
   * Returns null if no pattern has been set; PowerPoint may display the default `PERCENT_5`
   * pattern in this case. Assigning null will remove any explicit pattern setting.
   */
  get pattern(): MSO_PATTERN_TYPE | null {
    return this._pattFill.prst;
  }
  set pattern(patternType: MSO_PATTERN_TYPE | null) {
    this._pattFill.prst = patternType;
  }

  get type(): MSO_FILL_TYPE {
    return MSO_FILL.PATTERNED;
  }
}

/** Provides access to fill properties such as color for solid fills. */
export class _SolidFill extends _Fill {
  declare _solidFill: CT_SolidColorFillProperties;

  constructor(solidFill: CT_SolidColorFillProperties) {
    super();
    this._solidFill = solidFill;
  }

  /** Color of this fill. */
  get foreColor(): ColorFormat {
    return lazy(this, "foreColor", () => ColorFormat.fromColorchoiceParent(this._solidFill));
  }

  get type(): MSO_FILL_TYPE {
    return MSO_FILL.SOLID;
  }
}

/**
 * Collection of `GradientStop` objects defining gradient colors.
 *
 * A gradient must have a minimum of two stops, but can have as many more than that as
 * required to achieve the desired effect (three is perhaps most common). Stops are sequenced
 * in the order they are transitioned through.
 */
export class _GradientStops implements Iterable<_GradientStop> {
  declare _gsLst: CT_GradientStopList;

  constructor(gsLst: CT_GradientStopList) {
    this._gsLst = gsLst;
  }

  at(idx: number): _GradientStop {
    const gs = this._gsLst.gsLst.at(idx);
    if (gs === undefined) throw new RangeError("list index out of range");
    return new _GradientStop(gs);
  }

  [PY_GETITEM](idx: number): _GradientStop {
    return this.at(idx);
  }

  *[Symbol.iterator](): Generator<_GradientStop> {
    for (let i = 0; i < this.length; i++) yield this.at(i);
  }

  get length(): number {
    return this._gsLst.length;
  }
}

/**
 * Gradient stop object defining a single stop in a gradient.
 *
 * A gradient stop defines a color and a position.
 */
export class _GradientStop extends ElementProxy {
  declare _gs: CT_GradientStop;

  constructor(gs: CT_GradientStop) {
    super(gs);
    this._gs = gs;
  }

  /** `ColorFormat` object for controlling color of this gradient stop. */
  get color(): ColorFormat {
    return lazy(this, "color", () => ColorFormat.fromColorchoiceParent(this._gs));
  }

  /**
   * Location of stop in gradient path as float between 0.0 and 1.0.
   *
   * The value represents a percentage, where 0.0 (0%) represents the start of the path and 1.0
   * (100%) represents the end of the path. For a linear gradient, these would represent
   * opposing extents of the filled area.
   */
  get position(): number {
    return this._gs.pos;
  }
  set position(value: number) {
    this._gs.pos = Number(value);
  }
}
