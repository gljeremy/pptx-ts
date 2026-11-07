// Simple-type classes (`pptx.oxml.simpletypes`): validation and format translation for
// values stored in XML element attributes. Names follow the XML schema simple types.
//
// JS has a single number type, so "integral" means `Number.isInteger(value)`; booleans
// are not numbers here (python accepts True where an int is expected).
import { ValueError } from "../errors.js";
import { InvalidXmlError } from "../exc.js";
import { Centipoints, Emu, pyFloat, pyFloatStr, pyInt, pyMod, pyRound, pyTypeName } from "../util.js";

const isNum = (v: unknown): v is number => typeof v === "number" || v instanceof Number;
const isStr = (v: unknown): v is string => typeof v === "string" || v instanceof String;

export class BaseSimpleType {
  static fromXml(xmlValue: string): any {
    return this.convertFromXml(xmlValue);
  }

  static toXml(value: any): string {
    this.validate(value);
    return this.convertToXml(value);
  }

  static convertFromXml(strValue: string): any {
    return strValue;
  }

  static convertToXml(value: any): string {
    return String(value);
  }

  static validate(_value: any): void {}

  static validateFloat(value: any): void {
    if (!isNum(value)) throw new TypeError(`value must be a number, got ${pyTypeName(value)}`);
  }

  static validateInt(value: any): void {
    if (!isNum(value) || !Number.isInteger(Number(value)))
      throw new TypeError(`value must be an integral type, got ${pyTypeName(value)}`);
  }

  static validateFloatInRange(value: any, minInclusive: number, maxInclusive: number): void {
    this.validateFloat(value);
    if (value < minInclusive || value > maxInclusive)
      throw new ValueError(`value must be in range ${fmt(minInclusive)} to ${fmt(maxInclusive)} inclusive, got ${fmt(value)}`);
  }

  static validateIntInRange(value: any, minInclusive: number, maxInclusive: number): void {
    this.validateInt(value);
    if (value < minInclusive || value > maxInclusive)
      throw new ValueError(`value must be in range ${minInclusive} to ${maxInclusive} inclusive, got ${Number(value)}`);
  }

  static validateString(value: any): string {
    if (isStr(value)) return String(value);
    throw new TypeError(`value must be a string, got ${pyTypeName(value)}`);
  }
}

/** python's "%s" of a number */
const fmt = (v: any) => (typeof v === "number" && !Number.isInteger(v) ? pyFloatStr(v) : String(v));

export class BaseFloatType extends BaseSimpleType {
  static convertFromXml(strValue: string): number {
    return pyFloat(strValue);
  }
  static convertToXml(value: any): string {
    return pyFloatStr(Number(value));
  }
  static validate(value: any): void {
    if (!isNum(value)) throw new TypeError(`value must be a number, got ${pyTypeName(value)}`);
  }
}

export class BaseIntType extends BaseSimpleType {
  static convertFromPercentLiteral(strValue: string): number {
    return pyInt(strValue.replace(/%/g, ""));
  }
  static convertFromXml(strValue: string): number {
    return pyInt(strValue);
  }
  static convertToXml(value: any): string {
    return String(Number(value));
  }
  static validate(value: any): void {
    this.validateInt(value);
  }
}

export class BaseStringType extends BaseSimpleType {
  static convertFromXml(strValue: string): string {
    return strValue;
  }
  static convertToXml(value: any): string {
    return String(value);
  }
  static validate(value: any): void {
    this.validateString(value);
  }
}

export class BaseStringEnumerationType extends BaseStringType {
  static _members: readonly string[] = [];
  static validate(value: any): void {
    this.validateString(value);
    if (!this._members.includes(String(value)))
      throw new ValueError(`must be one of (${this._members.map((m) => `'${m}'`).join(", ")}), got '${value}'`);
  }
}

export class XsdAnyUri extends BaseStringType {}

export class XsdBoolean extends BaseSimpleType {
  static convertFromXml(strValue: string): boolean {
    if (!["1", "0", "true", "false"].includes(strValue))
      throw new InvalidXmlError(`value must be one of '1', '0', 'true' or 'false', got '${strValue}'`);
    return strValue === "1" || strValue === "true";
  }
  static convertToXml(value: any): string {
    return value ? "1" : "0";
  }
  static validate(value: any): void {
    if (value !== true && value !== false)
      throw new TypeError(`only True or False (and possibly None) may be assigned, got '${value}'`);
  }
}

export class XsdDouble extends BaseFloatType {}

export class XsdId extends BaseStringType {}

export class XsdInt extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, -2147483648, 2147483647);
  }
}

export class XsdLong extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, -9223372036854775808, 9223372036854775807);
  }
}

export class XsdString extends BaseStringType {}

export class XsdStringEnumeration extends BaseStringEnumerationType {}

export class XsdToken extends BaseStringType {}

export class XsdTokenEnumeration extends BaseStringEnumerationType {}

export class XsdUnsignedByte extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 255);
  }
}

export class XsdUnsignedInt extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 4294967295);
  }
}

export class XsdUnsignedShort extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 65535);
  }
}

/** `rot` attribute on `<a:xfrm>`: 60000ths of a degree. */
export class ST_Angle extends XsdInt {
  static DEGREE_INCREMENTS = 60000;
  static THREE_SIXTY = 360 * 60000;

  static convertFromXml(strValue: string): number {
    const rot = pyMod(pyInt(strValue), this.THREE_SIXTY);
    return rot / this.DEGREE_INCREMENTS;
  }
  /** Signed angle float like -42.42 to int 60000 per degree, normalized to positive. */
  static convertToXml(value: any): string {
    const rot = pyMod(Math.trunc(pyRound(Number(value) * this.DEGREE_INCREMENTS)), this.THREE_SIXTY);
    return String(rot);
  }
  static validate(value: any): void {
    BaseFloatType.validate(value);
  }
}

export class ST_AxisUnit extends XsdDouble {
  static validate(value: any): void {
    super.validate(value);
    if (value <= 0.0) throw new ValueError(`must be positive numeric value, got ${fmt(value)}`);
  }
}

export class ST_BarDir extends XsdStringEnumeration {
  static BAR = "bar";
  static COL = "col";
  static _members = ["bar", "col"];
}

export class ST_BubbleScale extends BaseIntType {
  static convertFromXml(strValue: string): number {
    if (strValue.includes("%")) return this.convertFromPercentLiteral(strValue);
    return super.convertFromXml(strValue);
  }
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 300);
  }
}

export class ST_ContentType extends XsdString {}

export class ST_CoordinateUnqualified extends XsdLong {
  static validate(value: any): void {
    this.validateIntInRange(value, -27273042329600, 27273042316900);
  }
}

export class ST_UniversalMeasure extends BaseSimpleType {
  static convertFromXml(strValue: string): number {
    const floatPart = strValue.slice(0, -2);
    const unitsPart = strValue.slice(-2);
    const quantity = pyFloat(floatPart);
    const multiplier = ({ mm: 36000, cm: 360000, in: 914400, pt: 12700, pc: 152400, pi: 152400 } as Record<string, number>)[unitsPart];
    if (multiplier === undefined) throw new Error(`KeyError: '${unitsPart}'`);
    return Emu(Math.trunc(pyRound(quantity * multiplier)));
  }
}

export class ST_Coordinate extends BaseSimpleType {
  static convertFromXml(strValue: string): number {
    if (/[imp]/.test(strValue)) return ST_UniversalMeasure.convertFromXml(strValue);
    return Emu(pyInt(strValue));
  }
  static convertToXml(value: any): string {
    return String(Number(value));
  }
  static validate(value: any): void {
    ST_CoordinateUnqualified.validate(value);
  }
}

export class ST_Coordinate32Unqualified extends XsdInt {
  static convertFromXml(strValue: string): number {
    return Emu(pyInt(strValue));
  }
}

/** xsd:union of ST_Coordinate32Unqualified, ST_UniversalMeasure */
export class ST_Coordinate32 extends BaseSimpleType {
  static convertFromXml(strValue: string): number {
    if (/[imp]/.test(strValue)) return ST_UniversalMeasure.convertFromXml(strValue);
    return ST_Coordinate32Unqualified.convertFromXml(strValue);
  }
  static convertToXml(value: any): string {
    return ST_Coordinate32Unqualified.convertToXml(value);
  }
  static validate(value: any): void {
    ST_Coordinate32Unqualified.validate(value);
  }
}

export class ST_Direction extends XsdTokenEnumeration {
  static HORZ = "horz";
  static VERT = "vert";
  static _members = ["horz", "vert"];
}

export class ST_DrawingElementId extends XsdUnsignedInt {}

export class ST_Extension extends XsdString {}

export class ST_GapAmount extends BaseIntType {
  static convertFromXml(strValue: string): number {
    if (strValue.includes("%")) return this.convertFromPercentLiteral(strValue);
    return super.convertFromXml(strValue);
  }
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 500);
  }
}

/** <c:grouping val=""> (also used as ST_BarGrouping). */
export class ST_Grouping extends XsdStringEnumeration {
  static CLUSTERED = "clustered";
  static PERCENT_STACKED = "percentStacked";
  static STACKED = "stacked";
  static STANDARD = "standard";
  static _members = ["clustered", "percentStacked", "stacked", "standard"];
}

export class ST_HexColorRGB extends BaseStringType {
  /** Keep alpha characters all uppercase just for consistency. */
  static convertToXml(value: any): string {
    return String(value).toUpperCase();
  }
  static validate(value: any): void {
    const s = this.validateString(value);
    if (s.length !== 6) throw new ValueError(`RGB string must be six characters long, got '${s}'`);
    if (!/^\s*[+-]?[0-9a-fA-F]+(_[0-9a-fA-F]+)*\s*$/.test(s)) throw new ValueError(`RGB string must be valid hex string, got '${s}'`);
  }
}

export class ST_LayoutMode extends XsdStringEnumeration {
  static EDGE = "edge";
  static FACTOR = "factor";
  static _members = ["edge", "factor"];
}

/** Unsigned integer 0..1000 with optional '%' suffix. */
export class ST_LblOffset extends XsdUnsignedShort {
  static convertFromXml(strValue: string): number {
    if (strValue.endsWith("%")) return this.convertFromPercentLiteral(strValue);
    return pyInt(strValue);
  }
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 1000);
  }
}

export class ST_LineWidth extends XsdInt {
  static convertFromXml(strValue: string): number {
    return Emu(pyInt(strValue));
  }
  static validate(value: any): void {
    super.validate(value);
    if (value < 0 || value > 20116800)
      throw new ValueError(`value must be in range 0-20116800 inclusive (0-1584 points), got ${Number(value)}`);
  }
}

export class ST_MarkerSize extends XsdUnsignedByte {
  static validate(value: any): void {
    this.validateIntInRange(value, 2, 72);
  }
}

export class ST_Orientation extends XsdStringEnumeration {
  static MAX_MIN = "maxMin";
  static MIN_MAX = "minMax";
  static _members = ["maxMin", "minMax"];
}

export class ST_Overlap extends BaseIntType {
  static convertFromXml(strValue: string): number {
    if (strValue.includes("%")) return this.convertFromPercentLiteral(strValue);
    return super.convertFromXml(strValue);
  }
  static validate(value: any): void {
    this.validateIntInRange(value, -100, 100);
  }
}

/** Integer 1000ths of a percent ("42000") or a float with '%' suffix ("42.0%"). */
export class ST_Percentage extends BaseIntType {
  static convertFromXml(strValue: string): number {
    if (strValue.includes("%")) return this._convertFromPercentLiteral(strValue);
    return pyInt(strValue) / 100000.0;
  }
  static convertToXml(value: any): string {
    return String(Math.trunc(pyRound(Number(value) * 100000.0)));
  }
  static validate(value: any): void {
    this.validateFloatInRange(value, -21474.83648, 21474.83647);
  }
  static _convertFromPercentLiteral(strValue: string): number {
    return pyFloat(strValue.slice(0, -1)) / 100.0;
  }
}

export class ST_PlaceholderSize extends XsdTokenEnumeration {
  static FULL = "full";
  static HALF = "half";
  static QUARTER = "quarter";
  static _members = ["full", "half", "quarter"];
}

export class ST_PositiveCoordinate extends XsdLong {
  static convertFromXml(strValue: string): number {
    return Emu(super.convertFromXml(strValue));
  }
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 27273042316900);
  }
}

/** 60000ths of a degree, constrained to positive angles less than 360 degrees. */
export class ST_PositiveFixedAngle extends ST_Angle {
  static convertToXml(degrees: any): string {
    let d = Number(degrees);
    if (d < 0.0) {
      d = pyMod(d, -360);
      d += 360;
    } else if (d > 0.0) d = pyMod(d, 360);
    return String(Math.trunc(pyRound(d * this.DEGREE_INCREMENTS)));
  }
}

/** A float between 0.0 and 1.0 stored as 1000ths of a percent. */
export class ST_PositiveFixedPercentage extends ST_Percentage {
  static validate(value: any): void {
    this.validateFloatInRange(value, 0.0, 1.0);
  }
}

export class ST_RelationshipId extends XsdString {}

export class ST_SlideId extends XsdUnsignedInt {
  static validate(value: any): void {
    this.validateIntInRange(value, 256, 2147483647);
  }
}

export class ST_SlideSizeCoordinate extends BaseIntType {
  static convertFromXml(strValue: string): number {
    return Emu(pyInt(strValue));
  }
  static validate(value: any): void {
    this.validateInt(value);
    if (value < 914400 || value > 51206400)
      throw new ValueError(`value must be in range(914400, 51206400) (1-56 inches), got ${Number(value)}`);
  }
}

export class ST_Style extends XsdUnsignedByte {
  static validate(value: any): void {
    this.validateIntInRange(value, 1, 48);
  }
}

/** `TargetMode` of a Relationship element: 'External' or 'Internal'. */
export class ST_TargetMode extends XsdString {
  static validate(value: any): void {
    this.validateString(value);
    if (value !== "External" && value !== "Internal")
      throw new ValueError(`must be one of 'Internal' or 'External', got '${value}'`);
  }
}

/** `fontScale` attribute of `<a:normAutofit>`, as a float percent. */
export class ST_TextFontScalePercentOrPercentString extends BaseFloatType {
  static convertFromXml(strValue: string): number {
    if (strValue.endsWith("%")) return pyFloat(strValue.slice(0, -1));
    return pyInt(strValue) / 1000.0;
  }
  static convertToXml(value: any): string {
    return String(Math.trunc(Number(value) * 1000.0));
  }
  static validate(value: any): void {
    BaseFloatType.validate(value);
    if (value < 1.0 || value > 100.0) throw new ValueError(`value must be in range 1.0..100.0 (percent), got ${fmt(value)}`);
  }
}

export class ST_TextFontSize extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, 100, 400000);
  }
}

export class ST_TextIndentLevelType extends BaseIntType {
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 8);
  }
}

export class ST_TextSpacingPercentOrPercentString extends BaseFloatType {
  static convertFromXml(strValue: string): number {
    if (strValue.endsWith("%")) return this._convertFromPercentLiteral(strValue);
    return pyInt(strValue) / 100000.0;
  }
  static _convertFromPercentLiteral(strValue: string): number {
    return pyFloat(strValue.slice(0, -1)) / 100.0;
  }
  /** 1.75 -> '175000' */
  static convertToXml(value: any): string {
    return String(Math.trunc(pyRound(Number(value) * 100000.0)));
  }
  static validate(value: any): void {
    this.validateFloatInRange(value, 0.0, 132.0);
  }
}

export class ST_TextSpacingPoint extends BaseIntType {
  /** Reads string integer centipoints, returns a length in EMU. */
  static convertFromXml(strValue: string): number {
    return Centipoints(pyInt(strValue));
  }
  static convertToXml(value: any): string {
    return String(Math.floor(Emu(value) / 127));
  }
  static validate(value: any): void {
    this.validateIntInRange(value, 0, 20116800);
  }
}

export class ST_TextTypeface extends XsdString {}

export class ST_TextWrappingType extends XsdTokenEnumeration {
  static NONE = "none";
  static SQUARE = "square";
  static _members = ["none", "square"];
}
