// DrawingML objects related to color, ColorFormat being the most prominent (`pptx.dml.color`).
import { MSO_COLOR_TYPE, MSO_THEME_COLOR } from "../enum/dml.js";
import { AttributeError, NotImplementedError, ValueError } from "../errors.js";
import { CT_HslColor, CT_PresetColor, CT_SchemeColor, CT_ScRgbColor, CT_SRgbColor, CT_SystemColor } from "../oxml/dml/color.js";
import { PY_GETITEM, PY_LEN } from "../oxml/xml.js";

/**
 * Provides access to color settings such as RGB color, theme color, and luminance
 * adjustments.
 */
export class ColorFormat {
  declare _xFill: any;
  declare _color: _Color;

  constructor(egColorChoiceParent: any, color: _Color) {
    this._xFill = egColorChoiceParent;
    this._color = color;
  }

  /**
   * Read/write float value between -1.0 and 1.0 indicating the brightness adjustment for this
   * color, e.g. -0.25 is 25% darker and 0.4 is 40% lighter. 0 means no brightness adjustment.
   */
  get brightness(): number {
    return this._color.brightness;
  }
  set brightness(value: number) {
    this._validateBrightnessValue(value);
    this._color.brightness = value;
  }

  static fromColorchoiceParent(egColorChoiceParent: any): ColorFormat {
    const xClr = egColorChoiceParent.egColorChoice;
    const color = new _Color(xClr);
    return new this(egColorChoiceParent, color);
  }

  /**
   * `RGBColor` value of this color, or null if no RGB color is explicitly defined for this font.
   * Setting this value to an `RGBColor` instance causes its type to change to
   * MSO_COLOR_TYPE.RGB. If the color was a theme color with a brightness adjustment, the
   * brightness adjustment is removed when changing it to an RGB color.
   */
  get rgb(): RGBColor | null {
    return this._color.rgb;
  }
  set rgb(rgb: RGBColor) {
    if (!(rgb instanceof RGBColor)) throw new ValueError("assigned value must be type RGBColor");
    // change to rgb color format if not already
    if (!(this._color instanceof _SRgbColor)) {
      const srgbClr = this._xFill.getOrChangeToSrgbClr();
      this._color = new _SRgbColor(srgbClr);
    }
    // the _SRgbColor instance does the setting
    this._color.rgb = rgb;
  }

  /**
   * Theme color value of this color.
   *
   * Value is a member of `MSO_THEME_COLOR`, e.g. `MSO_THEME_COLOR.ACCENT_1`. Raises
   * AttributeError on access if the color is not type `MSO_COLOR_TYPE.SCHEME`. Assigning a
   * member of `MSO_THEME_COLOR` causes the color's type to change to `MSO_COLOR_TYPE.SCHEME`.
   */
  get themeColor(): MSO_THEME_COLOR {
    return this._color.themeColor;
  }
  set themeColor(msoThemeColorIdx: MSO_THEME_COLOR) {
    // change to theme color format if not already
    if (!(this._color instanceof _SchemeColor)) {
      const schemeClr = this._xFill.getOrChangeToSchemeClr();
      this._color = new _SchemeColor(schemeClr);
    }
    this._color.themeColor = msoThemeColorIdx;
  }

  /**
   * Read-only. A value from `MSO_COLOR_TYPE`, either RGB or SCHEME, corresponding to the way
   * this color is defined, or null if no color is defined at the level of this font.
   */
  get type(): MSO_COLOR_TYPE | null {
    return this._color.colorType;
  }

  _validateBrightnessValue(value: number): void {
    if (value < -1.0 || value > 1.0) throw new ValueError("brightness must be number in range -1.0 to 1.0");
    if (this._color instanceof _NoneColor)
      throw new ValueError("can't set brightness when color.type is None. Set color.rgb or .theme_color first.");
  }
}

/**
 * Object factory for color object of the appropriate type, also the base class for all color
 * type classes such as SRgbColor. `new _Color(xClr)` returns an instance of the subclass for
 * the type of `xClr` (python's `_Color.__new__`).
 */
export class _Color {
  declare _xClr: any;

  constructor(xClr: any) {
    if (new.target === _Color) {
      const colorCls = colorClsFor(xClr);
      return new colorCls(xClr);
    }
    this._xClr = xClr;
  }

  get brightness(): number {
    const [lumMod, lumOff] = [this._xClr.lumMod, this._xClr.lumOff];
    // a tint is lighter, a shade is darker; only tints have a lumOff child
    if (lumOff !== null) return lumOff.val;
    // which leaves shades, if lumMod is present
    if (lumMod !== null) return lumMod.val - 1.0;
    // there's no brightness adjustment if no lum{Mod|Off} elements
    return 0;
  }
  set brightness(value: number) {
    if (value > 0) this._tint(value);
    else if (value < 0) this._shade(value);
    else this._xClr.clearLum();
  }

  get colorType(): MSO_COLOR_TYPE | null {
    throw new NotImplementedError(`.color_type property must be implemented on ${this.constructor.name}`);
  }

  /** Raises AttributeError on access unless overridden by subclass. */
  get rgb(): RGBColor | null {
    throw new AttributeError(`no .rgb property on color type '${this.constructor.name}'`);
  }
  set rgb(_rgb: RGBColor) {
    throw new AttributeError(`can't set .rgb on color type '${this.constructor.name}'`);
  }

  /** `MSO_THEME_COLOR.NOT_THEME_COLOR` unless overridden by subclass. */
  get themeColor(): MSO_THEME_COLOR {
    return MSO_THEME_COLOR.NOT_THEME_COLOR;
  }
  set themeColor(_value: MSO_THEME_COLOR) {
    throw new AttributeError(`can't set .theme_color on color type '${this.constructor.name}'`);
  }

  _shade(value: number): void {
    const lumModVal = 1.0 - Math.abs(value);
    const colorElm = this._xClr.clearLum();
    colorElm.addLumMod(lumModVal);
  }

  _tint(value: number): void {
    const lumOffVal = value;
    const lumModVal = 1.0 - lumOffVal;
    const colorElm = this._xClr.clearLum();
    colorElm.addLumMod(lumModVal);
    colorElm.addLumOff(lumOffVal);
  }
}

export class _HslColor extends _Color {
  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.HSL;
  }
}

export class _NoneColor extends _Color {
  get colorType(): null {
    return null;
  }

  /** Raise AttributeError on attempt to access .themeColor when no color choice is present. */
  get themeColor(): MSO_THEME_COLOR {
    throw new AttributeError(`no .theme_color property on color type '${this.constructor.name}'`);
  }
}

export class _PrstColor extends _Color {
  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.PRESET;
  }
}

export class _SchemeColor extends _Color {
  declare _schemeClr: CT_SchemeColor;

  constructor(schemeClr: CT_SchemeColor) {
    super(schemeClr);
    this._schemeClr = schemeClr;
  }

  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.SCHEME;
  }

  /**
   * Theme color value of this color, one of those defined in the MSO_THEME_COLOR enumeration,
   * e.g. MSO_THEME_COLOR.ACCENT_1. Setting this to a value in MSO_THEME_COLOR causes the
   * color's type to change to `MSO_COLOR_TYPE.SCHEME`.
   */
  get themeColor(): MSO_THEME_COLOR {
    return this._schemeClr.val;
  }
  set themeColor(msoThemeColorIdx: MSO_THEME_COLOR) {
    this._schemeClr.val = msoThemeColorIdx;
  }
}

export class _ScRgbColor extends _Color {
  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.SCRGB;
  }
}

export class _SRgbColor extends _Color {
  declare _srgbClr: CT_SRgbColor;

  constructor(srgbClr: CT_SRgbColor) {
    super(srgbClr);
    this._srgbClr = srgbClr;
  }

  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.RGB;
  }

  /** `RGBColor` value of this color, from the required `val` attribute of `<a:srgbClr>`. */
  get rgb(): RGBColor {
    return RGBColor.fromString(this._srgbClr.val);
  }
  set rgb(rgb: RGBColor) {
    this._srgbClr.val = String(rgb);
  }
}

export class _SysColor extends _Color {
  get colorType(): MSO_COLOR_TYPE {
    return MSO_COLOR_TYPE.SYSTEM;
  }
}

/** The `_Color` subclass for the class of `xClr` (null for no color). */
function colorClsFor(xClr: unknown): new (xClr: any) => _Color {
  if (xClr === null || xClr === undefined) return _NoneColor;
  const ctor = (xClr as object).constructor;
  if (ctor === CT_HslColor) return _HslColor;
  if (ctor === CT_PresetColor) return _PrstColor;
  if (ctor === CT_SchemeColor) return _SchemeColor;
  if (ctor === CT_ScRgbColor) return _ScRgbColor;
  if (ctor === CT_SRgbColor) return _SRgbColor;
  if (ctor === CT_SystemColor) return _SysColor;
  throw new Error(`KeyError: ${ctor?.name}`);
}

/**
 * Immutable value object defining a particular RGB color.
 *
 * Behaves like python's 3-tuple: `color[0]`, `[...color]`, `.length`; `String(color)` is the
 * hex string, e.g. '3C2F80'.
 */
export class RGBColor implements Iterable<number> {
  declare readonly 0: number;
  declare readonly 1: number;
  declare readonly 2: number;

  constructor(r: number, g: number, b: number) {
    for (const val of [r, g, b]) {
      if (!Number.isInteger(val) || val < 0 || val > 255) throw new ValueError("RGBColor() takes three integer values 0-255");
    }
    Object.defineProperties(this, { 0: { value: r, enumerable: true }, 1: { value: g, enumerable: true }, 2: { value: b, enumerable: true } });
    Object.freeze(this);
  }

  get length(): 3 {
    return 3;
  }

  *[Symbol.iterator](): Generator<number> {
    yield this[0];
    yield this[1];
    yield this[2];
  }

  [PY_GETITEM](idx: number): number {
    const v = [this[0], this[1], this[2]].at(idx);
    if (v === undefined) throw new RangeError("tuple index out of range");
    return v;
  }

  [PY_LEN](): number {
    return 3;
  }

  /** Equal to another RGBColor or a [r, g, b] array with the same values (python tuple `==`). */
  equals(other: unknown): boolean {
    if (!(other instanceof RGBColor) && !Array.isArray(other)) return false;
    const o = [...(other as Iterable<number>)];
    return o.length === 3 && o[0] === this[0] && o[1] === this[1] && o[2] === this[2];
  }

  /** The hex string rgb value, like '3C2F80'. */
  toString(): string {
    return [this[0], this[1], this[2]].map((x) => x.toString(16).toUpperCase().padStart(2, "0")).join("");
  }

  __repr__(): string {
    return `(${this[0]}, ${this[1]}, ${this[2]})`;
  }

  /** A new instance from an RGB color hex string like `'3C2F80'`. */
  static fromString(rgbHexStr: string): RGBColor {
    const hex = (s: string) => {
      if (!/^[0-9a-fA-F]+$/.test(s)) throw new ValueError(`invalid literal for int() with base 16: '${s}'`);
      return parseInt(s, 16);
    };
    return new this(hex(rgbHexStr.slice(0, 2)), hex(rgbHexStr.slice(2, 4)), hex(rgbHexStr.slice(4)));
  }
}
