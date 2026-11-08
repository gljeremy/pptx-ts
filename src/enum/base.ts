// Base classes for enumerations (`pptx.enum.base`).
//
// python-pptx's enums are `int` enums: a member compares equal to its integer value, has
// `.name`/`.value`, and the enum class is callable to look a member up by value
// (`PP_ACTION_TYPE(10)`). Here a member is a `Number` object (so `+member`, `member == 6`
// and arithmetic work) and an enum is a callable object made by `defineEnum()`, with its
// members as properties and the class-level methods (`fromXml`, `toXml`, `validate`).
import { ValueError } from "../errors.js";

/** A member of a `BaseEnum`: a Number with a name and docstring. */
export class BaseEnum extends Number {
  declare readonly name: string;
  declare readonly __doc__: string;
  /** The enum this member belongs to. */
  declare readonly __enum__: EnumClass<any>;

  /** The integer value of this member (the MS API value). */
  get value(): number {
    return this.valueOf();
  }

  /** The symbolic name and value of this member, e.g. 'MIDDLE (3)'. */
  toString(): string {
    return `${this.name} (${this.value})`;
  }

  /** Python repr, e.g. '<PP_ACTION_TYPE.END_SHOW: 6>'. */
  __repr__(): string {
    return `<${this.__enum__.__name__}.${this.name}: ${this.value}>`;
  }

  toJSON(): number {
    return this.value;
  }
}

/** A member of a `BaseXmlEnum`: also maps to an XML attribute value. */
export class BaseXmlEnum extends BaseEnum {
  /** XML attribute value for this member; "" or null for "return-value only" members. */
  declare readonly xmlValue: string | null;
}

type MemberDef = readonly [value: number, docstr: string] | readonly [value: number, xmlValue: string | null, docstr: string];

/** The callable enum object: `E(value)` looks up a member, members are properties of `E`. */
export interface EnumClass<M extends BaseEnum> extends Iterable<M> {
  (value: number | M): M;
  readonly __name__: string;
  readonly __members__: Record<string, M>;
  readonly prototype: M;
  /** Members in definition order. */
  readonly members: readonly M[];
  /** True when `value` is a member, or the int value of one. */
  has(value: unknown): boolean;
  /** Raise ValueError if `value` is not an assignable value. */
  validate(value: unknown): void;
}

/** Class-level methods of an enum mapping XML attribute values. */
export interface XmlEnumClass<M extends BaseXmlEnum> extends EnumClass<M> {
  /** Member corresponding to XML attribute value `xmlValue`; ValueError when there's none. */
  fromXml(xmlValue: string): M;
  /** XML value of `value` (a member or its int value); ValueError when it has none. */
  toXml(value: number | M): string;
}

export type EnumOf<Defs, M extends BaseEnum> = { readonly [K in keyof Defs]: M };

/**
 * Define an enum named `name` whose members are `defs` ({NAME: [value, docstr]} for a
 * `BaseEnum`, {NAME: [value, xmlValue, docstr]} for a `BaseXmlEnum`).
 */
export function defineEnum<const D extends Record<string, readonly [number, string]>>(name: string, base: typeof BaseEnum, defs: D): EnumClass<BaseEnum> & EnumOf<D, BaseEnum>;
export function defineEnum<const D extends Record<string, readonly [number, string | null, string]>>(
  name: string,
  base: typeof BaseXmlEnum,
  defs: D,
): XmlEnumClass<BaseXmlEnum> & EnumOf<D, BaseXmlEnum>;
export function defineEnum(name: string, base: typeof BaseEnum, defs: Record<string, MemberDef>): any {
  const byValue = new Map<number, BaseEnum>();
  const members: BaseEnum[] = [];
  const memberMap: Record<string, BaseEnum> = {};

  const E: any = function (value: unknown) {
    if (value instanceof BaseEnum && value.__enum__ === E) return value;
    const n = value instanceof Number ? value.valueOf() : value;
    const m = typeof n === "number" ? byValue.get(n) : undefined;
    if (m === undefined) throw new ValueError(`${typeof value === "string" ? `'${value}'` : String(n)} is not a valid ${name}`);
    return m;
  };
  // members are instances of the enum (python `isinstance(member, PP_ACTION_TYPE)`)
  E.prototype = Object.create(base.prototype, { constructor: { value: E, writable: true, configurable: true } });
  Object.defineProperty(E, "name", { value: name });
  E.__name__ = name;
  E.__members__ = memberMap;
  E.members = members;

  const isXml = base === BaseXmlEnum || base.prototype instanceof BaseXmlEnum;
  for (const [memberName, def] of Object.entries(defs)) {
    const value = def[0];
    const xmlValue = isXml ? (def[1] as string | null) : undefined;
    const docstr = String(isXml ? def[2] : def[1]).trim();
    // python Enum: a later member with an existing value is an alias of the first one
    const existing = byValue.get(value);
    if (existing !== undefined) {
      memberMap[memberName] = existing;
      E[memberName] = existing;
      continue;
    }
    const m: any = Reflect.construct(Number, [value], E);
    Object.defineProperty(m, "name", { value: memberName, enumerable: true });
    Object.defineProperty(m, "__doc__", { value: docstr });
    Object.defineProperty(m, "__enum__", { value: E });
    if (isXml) Object.defineProperty(m, "xmlValue", { value: xmlValue, enumerable: true });
    Object.freeze(m);
    byValue.set(value, m);
    members.push(m);
    memberMap[memberName] = m;
    E[memberName] = m;
  }

  E[Symbol.iterator] = () => members[Symbol.iterator]();
  E.has = (value: unknown) => (value instanceof BaseEnum ? value.__enum__ === E : typeof value === "number" && byValue.has(value));
  E.validate = (value: unknown) => {
    if (!(value instanceof BaseEnum && value.__enum__ === E)) throw new ValueError(`${String(value)} not a member of ${name} enumeration`);
  };
  if (isXml) {
    E.fromXml = (xmlValue: string) => {
      // the empty string never maps to a member
      const m = xmlValue ? members.find((m: any) => m.xmlValue === xmlValue) : undefined;
      if (m === undefined) throw new ValueError(`${name} has no XML mapping for ${pyRepr(xmlValue)}`);
      return m;
    };
    E.toXml = (value: number | BaseEnum) => {
      const m: any = E(value);
      if (!m.xmlValue) throw new ValueError(`${name}.${m.name} has no XML representation`);
      return m.xmlValue;
    };
  }
  return Object.freeze(E);
}

function pyRepr(v: unknown): string {
  if (typeof v !== "string") return String(v);
  return v.includes("'") && !v.includes('"') ? `"${v}"` : `'${v.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}'`;
}
