// Declarative definition of custom element classes (`pptx.oxml.xmlchemy`).
//
// python-pptx uses a metaclass to turn class attributes such as
// `solidFill = ZeroOrOne("a:solidFill", successors=...)` into generated
// properties and methods. Here a class calls `defineChildren(this, {...})` in a
// static block (or right after its definition) to the same effect. Generated
// names are the python names in camelCase: `get_or_add_spPr` -> `getOrAddSpPr`,
// `_add_p` -> `_addP`, `p_lst` -> `pLst`.
import { InvalidXmlError } from "../exc.js";
import { NamespacePrefixedTag, _nsmap, qn } from "./ns.js";
import { makeelement } from "./parser.js";
import { setDefaultElementClass, tostring, XmlElement } from "./xml.js";
import { xpath as evalXpath } from "./xpath.js";

/** Transforms values to and from the XML string value of an attribute. */
export interface AttributeType {
  fromXml(xmlValue: string): any;
  toXml(value: any): string;
  readonly name?: string;
}

/** python name -> camelCase name used in pptx-ts (`_add_spPr` -> `_addSpPr`). */
export function camel(name: string): string {
  const m = /^_*/.exec(name)![0];
  const rest = name.slice(m.length);
  return m + rest.replace(/_+([A-Za-z0-9])/g, (_s, c: string) => c.toUpperCase());
}

/** A "loose" element having the tag `nsptagStr` (e.g. 'a:tbl'), of its custom class if one is registered. */
export function OxmlElement(nsptagStr: string, nsmap?: Record<string, string> | null): any {
  const nsptag = new NamespacePrefixedTag(nsptagStr);
  return makeelement(nsptag.clarkName, nsmap ?? nsptag.nsmap);
}

/** Human-readable XML for tests: pretty printed, no XML declaration. */
export function serializeForReading(element: XmlElement): XmlString {
  return new XmlString(tostring(element, { prettyPrint: true }));
}

const XML_ELM_LINE = /^( *<\/?[\w:]+)(.*?)(\/?>)([^<]*<\/[\w:]+>)?$/;

/** A str whose equality comparison ignores attribute order, line by line; useful for tests. */
export class XmlString extends String {
  equals(other: unknown): boolean {
    if (typeof other !== "string" && !(other instanceof String)) return false;
    const lines = splitlines(this.valueOf());
    const linesOther = splitlines(String(other));
    if (lines.length !== linesOther.length) return false;
    for (let i = 0; i < lines.length; i++) if (!this._eqElmStrs(lines[i], linesOther[i])) return false;
    return true;
  }

  _attrSeq(attrs: string): string[] {
    return attrs.trim().split(/\s+/).filter(Boolean).sort();
  }

  _eqElmStrs(line: string, line2: string): boolean {
    const [front, attrs, close, text] = this._parseLine(line);
    const [front2, attrs2, close2, text2] = this._parseLine(line2);
    if (front !== front2) return false;
    const a = this._attrSeq(attrs);
    const b = this._attrSeq(attrs2);
    if (a.length !== b.length || a.some((x, i) => x !== b[i])) return false;
    return close === close2 && text === text2;
  }

  _parseLine(line: string): [string, string, string, string | undefined] {
    const m = XML_ELM_LINE.exec(line);
    if (!m) throw new Error("ValueError: `line` does not match pattern for an XML element");
    return [m[1], m[2], m[3], m[4]];
  }
}

function splitlines(s: string): string[] {
  const lines = s.split(/\r\n|\r|\n/);
  if (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

// -- descriptors ---------------------------------------------------------------------------------

type Cls = { prototype: any; name: string };

function addToClass(cls: Cls, name: string, fn: (...a: any[]) => any): void {
  // a method already defined by hand on the class (or a base) wins
  if (name in cls.prototype) return;
  Object.defineProperty(cls.prototype, name, { value: fn, writable: true, configurable: true });
}

function defineGetter(cls: Cls, name: string, get: (this: any) => any, set?: (this: any, v: any) => void): void {
  Object.defineProperty(cls.prototype, name, { get, set, configurable: true, enumerable: false });
}

abstract class BaseAttribute {
  _elementCls!: Cls;
  _propName!: string;
  _attrName: string;
  _simpleType: AttributeType;

  constructor(attrName: string, simpleType: AttributeType) {
    this._attrName = attrName;
    this._simpleType = simpleType;
  }

  populateClassMembers(cls: Cls, propName: string): void {
    this._elementCls = cls;
    this._propName = propName;
    defineGetter(cls, propName, this._getter(), this._setter());
  }

  get _clarkName(): string {
    return this._attrName.includes(":") ? qn(this._attrName) : this._attrName;
  }

  abstract _getter(): (this: any) => any;
  abstract _setter(): (this: any, v: any) => void;
}

/** An optional attribute: reads as `default` when absent; assigning the default (or null) removes it. */
export class OptionalAttribute extends BaseAttribute {
  _default: any;

  constructor(attrName: string, simpleType: AttributeType, dflt: any = null) {
    super(attrName, simpleType);
    this._default = dflt;
  }

  _getter() {
    const clark = this._clarkName;
    const st = this._simpleType;
    const dflt = this._default;
    return function (this: XmlElement) {
      const v = this.get(clark);
      return v === null ? dflt : st.fromXml(v);
    };
  }

  _setter() {
    const clark = this._clarkName;
    const st = this._simpleType;
    const dflt = this._default;
    return function (this: XmlElement, value: any) {
      // setting an attribute to its default value removes it
      if (value === undefined) value = null;
      if (value === dflt || (dflt !== null && typeof dflt === "object" && typeof dflt.equals === "function" && dflt.equals(value))) {
        this.attrib.delete(clark);
        return;
      }
      this.set(clark, st.toXml(value));
    };
  }
}

/** A required attribute: InvalidXmlError when missing on read; never removed on assignment. */
export class RequiredAttribute extends BaseAttribute {
  _getter() {
    const clark = this._clarkName;
    const st = this._simpleType;
    const name = this._attrName;
    return function (this: XmlElement) {
      const v = this.get(clark);
      if (v === null) throw new InvalidXmlError(`required '${name}' attribute not present on element ${this.tag}`);
      return st.fromXml(v);
    };
  }

  _setter() {
    const clark = this._clarkName;
    const st = this._simpleType;
    return function (this: XmlElement, value: any) {
      this.set(clark, st.toXml(value));
    };
  }
}

abstract class BaseChildElement {
  _elementCls!: Cls;
  _propName!: string;
  _nsptagname: string;
  _successors: readonly string[];

  constructor(nsptagname: string, successors: readonly string[] = []) {
    this._nsptagname = nsptagname;
    this._successors = successors;
  }

  populateClassMembers(cls: Cls, propName: string): void {
    this._elementCls = cls;
    this._propName = propName;
  }

  get _addMethodName() {
    return camel(`_add_${this._propName}`);
  }
  get _insertMethodName() {
    return camel(`_insert_${this._propName}`);
  }
  get _newMethodName() {
    return camel(`_new_${this._propName}`);
  }
  get _removeMethodName() {
    return camel(`_remove_${this._propName}`);
  }

  _addAdder(): void {
    const newName = this._newMethodName;
    const insertName = this._insertMethodName;
    addToClass(this._elementCls, this._addMethodName, function (this: any, attrs?: Record<string, any>) {
      const child = this[newName]();
      if (attrs) for (const [k, v] of Object.entries(attrs)) child[k] = v;
      this[insertName](child);
      return child;
    });
  }

  _addCreator(): void {
    const tag = this._nsptagname;
    addToClass(this._elementCls, this._newMethodName, function () {
      return OxmlElement(tag);
    });
  }

  _addGetter(): void {
    defineGetter(this._elementCls, this._propName, this._getter());
  }

  _getter(): (this: any) => any {
    const clark = qn(this._nsptagname);
    return function (this: XmlElement) {
      return this.find(clark);
    };
  }

  _addInserter(): void {
    const successors = this._successors;
    addToClass(this._elementCls, this._insertMethodName, function (this: BaseOxmlElement, child: XmlElement) {
      this.insertElementBefore(child, ...successors);
      return child;
    });
  }

  _addListGetter(): void {
    const clark = qn(this._nsptagname);
    defineGetter(this._elementCls, camel(`${this._propName}_lst`), function (this: XmlElement) {
      return this.findall(clark);
    });
  }
}

/** A child element belonging to a group, only one of which may appear as a child. */
export class Choice extends BaseChildElement {
  _groupPropName!: string;

  constructor(nsptagname: string) {
    super(nsptagname, []);
  }

  get nsptagname(): string {
    return this._nsptagname;
  }

  populateChoice(cls: Cls, groupPropName: string, successors: readonly string[]): void {
    this._elementCls = cls;
    this._groupPropName = groupPropName;
    this._successors = successors;
    const tag = this._nsptagname;
    this._propName = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;
    this._addGetter();
    this._addCreator();
    this._addInserter();
    this._addAdder();
    this._addGetOrChangeToMethod();
  }

  _addGetOrChangeToMethod(): void {
    const prop = this._propName;
    const removeGroup = camel(`_remove_${this._groupPropName}`);
    const add = this._addMethodName;
    addToClass(this._elementCls, camel(`get_or_change_to_${prop}`), function (this: any) {
      const child = this[prop];
      if (child !== null) return child;
      this[removeGroup]();
      return this[add]();
    });
  }
}

/** A required child element. */
export class OneAndOnlyOne extends BaseChildElement {
  constructor(nsptagname: string) {
    super(nsptagname, []);
  }

  populateClassMembers(cls: Cls, propName: string): void {
    super.populateClassMembers(cls, propName);
    this._addGetter();
  }

  _getter() {
    const clark = qn(this._nsptagname);
    const tag = this._nsptagname;
    return function (this: XmlElement) {
      const child = this.find(clark);
      if (child === null) throw new InvalidXmlError(`required \`\`<${tag}>\`\` child element not present`);
      return child;
    };
  }
}

/** A repeating child element that must appear at least once. */
export class OneOrMore extends BaseChildElement {
  populateClassMembers(cls: Cls, propName: string): void {
    super.populateClassMembers(cls, propName);
    this._addListGetter();
    this._addCreator();
    this._addInserter();
    this._addAdder();
    const add = this._addMethodName;
    addToClass(cls, camel(`add_${propName}`), function (this: any) {
      return this[add]();
    });
  }
}

/** An optional repeating child element. */
export class ZeroOrMore extends BaseChildElement {
  populateClassMembers(cls: Cls, propName: string): void {
    super.populateClassMembers(cls, propName);
    this._addListGetter();
    this._addCreator();
    this._addInserter();
    this._addAdder();
  }
}

/** An optional child element. */
export class ZeroOrOne extends BaseChildElement {
  populateClassMembers(cls: Cls, propName: string): void {
    super.populateClassMembers(cls, propName);
    this._addGetter();
    this._addCreator();
    this._addInserter();
    this._addAdder();
    const prop = propName;
    const add = this._addMethodName;
    addToClass(cls, camel(`get_or_add_${propName}`), function (this: any) {
      let child = this[prop];
      if (child === null) child = this[add]();
      return child;
    });
    const tag = this._nsptagname;
    addToClass(cls, this._removeMethodName, function (this: BaseOxmlElement) {
      this.removeAll(tag);
    });
  }
}

/** An `EG_*` element group where at most one of its members may appear as a child. */
export class ZeroOrOneChoice extends BaseChildElement {
  _choices: Choice[];

  constructor(choices: Iterable<Choice>, successors: Iterable<string> = []) {
    super("", [...successors]);
    this._choices = [...choices];
  }

  populateClassMembers(cls: Cls, propName: string): void {
    super.populateClassMembers(cls, propName);
    const tags = this._memberNsptagnames;
    defineGetter(cls, propName, function (this: BaseOxmlElement) {
      return this.firstChildFoundIn(...tags);
    });
    for (const choice of this._choices) choice.populateChoice(cls, propName, this._successors);
    addToClass(cls, camel(`_remove_${propName}`), function (this: BaseOxmlElement) {
      for (const t of tags) this.removeAll(t);
    });
  }

  get _memberNsptagnames(): string[] {
    return this._choices.map((c) => c.nsptagname);
  }
}

type Member = BaseAttribute | BaseChildElement;

/**
 * Populate `cls` from its declarative members (python's MetaOxmlElement). Members are
 * processed in order, so earlier definitions take precedence just as in python.
 */
export function defineChildren(cls: Cls, members: Record<string, Member>): void {
  for (const [name, m] of Object.entries(members)) m.populateClassMembers(cls, name);
}
export const define = defineChildren;

// -- the base element class ----------------------------------------------------------------------

/** Effective base class for all custom element classes. */
export class BaseOxmlElement extends XmlElement {
  toString(): string {
    return `<${this.constructor.name} '<${this._nsptag}>' at 0x${idOf(this).toString(16)}>`;
  }

  /** First child with tag in `tagnames` (e.g. "a:p"), or null if not found. */
  firstChildFoundIn(...tagnames: string[]): any {
    for (const t of tagnames) {
      const child = this.find(qn(t));
      if (child !== null) return child;
    }
    return null;
  }

  insertElementBefore(elm: XmlElement, ...tagnames: string[]): any {
    const successor = this.firstChildFoundIn(...tagnames);
    if (successor !== null) successor.addprevious(elm);
    else this.append(elm);
    return elm;
  }

  /** Remove child elements with tagname (e.g. "a:p") in `tagnames`. */
  removeAll(...tagnames: string[]): void {
    for (const t of tagnames) for (const child of this.findall(qn(t))) this.remove(child);
  }

  /** XML string for this element, pretty printed and without an XML declaration. */
  get xml(): XmlString {
    return serializeForReading(this);
  }

  /** XPath evaluation with the standard Open XML namespace prefixes. */
  xpath(xpathStr: string): any {
    return evalXpath(this, xpathStr, _nsmap);
  }

  get _nsptag(): NamespacePrefixedTag {
    return NamespacePrefixedTag.fromClarkName(this.tag);
  }
}

const ids = new WeakMap<object, number>();
let nextId = 0x10000;
function idOf(o: object): number {
  let v = ids.get(o);
  if (v === undefined) ids.set(o, (v = nextId += 0x40));
  return v;
}

// unregistered tags parse as BaseOxmlElement too, so `.xpath()`/`.xml` work everywhere
setDefaultElementClass(BaseOxmlElement);
