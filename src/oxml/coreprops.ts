// Custom element classes for core properties-related XML elements (`pptx.oxml.coreprops`).
//
// python's naive `datetime` values are JS `Date`s holding that wall-clock time in UTC: a
// `<dcterms:created>2003-12-31T10:14:55Z</dcterms:created>` reads as
// `new Date(Date.UTC(2003, 11, 31, 10, 14, 55))`, and assigned dates are written from their
// UTC fields.
import { ValueError } from "../errors.js";
import { nsdecls, qn } from "./ns.js";
import { parseXml } from "./parser.js";
import type { XmlElement } from "./xml.js";
import { BaseOxmlElement, defineChildren, ZeroOrOne } from "./xmlchemy.js";

const PROPS = [
  ["category", "cp:category"],
  ["contentStatus", "cp:contentStatus"],
  ["created", "dcterms:created"],
  ["creator", "dc:creator"],
  ["description", "dc:description"],
  ["identifier", "dc:identifier"],
  ["keywords", "cp:keywords"],
  ["language", "dc:language"],
  ["lastModifiedBy", "cp:lastModifiedBy"],
  ["lastPrinted", "cp:lastPrinted"],
  ["modified", "dcterms:modified"],
  ["revision", "cp:revision"],
  ["subject", "dc:subject"],
  ["title", "dc:title"],
  ["version", "cp:version"],
] as const;

type PropName = (typeof PROPS)[number][0];

/**
 * `cp:coreProperties` element.
 *
 * The root element of the Core Properties part stored as `/docProps/core.xml`. Implements many
 * of the Dublin Core document metadata elements. String elements resolve to an empty string ('')
 * if the element is not present in the XML. String elements are limited in length to 255 unicode
 * characters.
 */
export class CT_CoreProperties extends BaseOxmlElement {
  declare category: XmlElement | null;
  declare contentStatus: XmlElement | null;
  declare created: XmlElement | null;
  declare creator: XmlElement | null;
  declare description: XmlElement | null;
  declare identifier: XmlElement | null;
  declare keywords: XmlElement | null;
  declare language: XmlElement | null;
  declare lastModifiedBy: XmlElement | null;
  declare lastPrinted: XmlElement | null;
  declare modified: XmlElement | null;
  declare revision: XmlElement | null;
  declare subject: XmlElement | null;
  declare title: XmlElement | null;
  declare version: XmlElement | null;
  declare getOrAddRevision: () => XmlElement;
  static {
    defineChildren(this, Object.fromEntries(PROPS.map(([name, tag]) => [name, new ZeroOrOne(tag, [])])));
  }

  static readonly _corePropertiesTmpl = `<cp:coreProperties ${nsdecls("cp", "dc", "dcterms")}/>\n`;

  /** A new `cp:coreProperties` element. */
  static newCoreProperties(): CT_CoreProperties {
    return parseXml(CT_CoreProperties._corePropertiesTmpl) as CT_CoreProperties;
  }

  get authorText(): string {
    return this._textOfElement("creator");
  }
  set authorText(value: string) {
    this._setElementText("creator", value);
  }

  get categoryText(): string {
    return this._textOfElement("category");
  }
  set categoryText(value: string) {
    this._setElementText("category", value);
  }

  get commentsText(): string {
    return this._textOfElement("description");
  }
  set commentsText(value: string) {
    this._setElementText("description", value);
  }

  get contentStatusText(): string {
    return this._textOfElement("contentStatus");
  }
  set contentStatusText(value: string) {
    this._setElementText("contentStatus", value);
  }

  get createdDatetime(): Date | null {
    return this._datetimeOfElement("created");
  }
  set createdDatetime(value: Date) {
    this._setElementDatetime("created", value);
  }

  get identifierText(): string {
    return this._textOfElement("identifier");
  }
  set identifierText(value: string) {
    this._setElementText("identifier", value);
  }

  get keywordsText(): string {
    return this._textOfElement("keywords");
  }
  set keywordsText(value: string) {
    this._setElementText("keywords", value);
  }

  get languageText(): string {
    return this._textOfElement("language");
  }
  set languageText(value: string) {
    this._setElementText("language", value);
  }

  get lastModifiedByText(): string {
    return this._textOfElement("lastModifiedBy");
  }
  set lastModifiedByText(value: string) {
    this._setElementText("lastModifiedBy", value);
  }

  get lastPrintedDatetime(): Date | null {
    return this._datetimeOfElement("lastPrinted");
  }
  set lastPrintedDatetime(value: Date) {
    this._setElementDatetime("lastPrinted", value);
  }

  get modifiedDatetime(): Date | null {
    return this._datetimeOfElement("modified");
  }
  set modifiedDatetime(value: Date) {
    this._setElementDatetime("modified", value);
  }

  /** Integer value of revision property. */
  get revisionNumber(): number {
    const revisionStr = this.revision?.text ?? null;
    if (revisionStr === null) return 0;
    // non-integer revision strings resolve to 0, as do negative integers
    const s = revisionStr.trim().replace(/_/g, "");
    if (!/^[+-]?\d+$/.test(s)) return 0;
    const revision = parseInt(s, 10);
    return revision < 0 ? 0 : revision;
  }
  /** Set revision property to string value of integer `value`. */
  set revisionNumber(value: number) {
    if (typeof value !== "number" || !Number.isInteger(value) || value < 1)
      throw new ValueError(`revision property requires positive int, got '${value}'`);
    this.getOrAddRevision().text = String(value);
  }

  get subjectText(): string {
    return this._textOfElement("subject");
  }
  set subjectText(value: string) {
    this._setElementText("subject", value);
  }

  get titleText(): string {
    return this._textOfElement("title");
  }
  set titleText(value: string) {
    this._setElementText("title", value);
  }

  get versionText(): string {
    return this._textOfElement("version");
  }
  set versionText(value: string) {
    this._setElementText("version", value);
  }

  _datetimeOfElement(propertyName: PropName): Date | null {
    const datetimeStr = this[propertyName]?.text ?? null;
    if (datetimeStr === null) return null;
    try {
      return CT_CoreProperties._parseW3CDTFToDatetime(datetimeStr);
    } catch (e) {
      // invalid datetime strings are ignored
      if (e instanceof ValueError) return null;
      throw e;
    }
  }

  /** Element returned by the 'getOrAdd' method for `propName`. */
  _getOrAdd(propName: PropName): XmlElement {
    const name = `getOrAdd${propName[0].toUpperCase()}${propName.slice(1)}`;
    return (this as any)[name]();
  }

  /**
   * `datetime` offset by the offset specified in `offsetStr`, a string like '-07:00'.
   * The offset is subtracted, converting a local time to UTC.
   */
  static _offsetDt(datetime: Date, offsetStr: string): Date {
    const match = /^([+-])(\d\d):(\d\d)/.exec(offsetStr);
    if (match === null) throw new ValueError(`'${offsetStr}' is not a valid offset string`);
    const [, sign, hoursStr, minutesStr] = match;
    const signFactor = sign === "+" ? -1 : 1;
    const ms = (parseInt(hoursStr, 10) * 60 + parseInt(minutesStr, 10)) * 60000 * signFactor;
    return new Date(datetime.getTime() + ms);
  }

  /**
   * Parse a W3CDTF date: yyyy ('2003'), yyyy-mm ('2003-12'), yyyy-mm-dd ('2003-12-31'),
   * UTC ('2003-12-31T10:14:55Z') or with a numeric timezone ('2003-12-31T10:14:55-08:00').
   */
  static _parseW3CDTFToDatetime(w3cdtfStr: string): Date {
    const parseablePart = w3cdtfStr.slice(0, 19);
    const offsetStr = w3cdtfStr.slice(19);
    const timestamp = strptimeW3C(parseablePart);
    if (timestamp === null) throw new ValueError(`could not parse W3CDTF datetime string '${w3cdtfStr}'`);
    if (offsetStr.length === 6) return CT_CoreProperties._offsetDt(timestamp, offsetStr);
    return timestamp;
  }

  /** Set date/time value of child element having `propName` to `value`. */
  _setElementDatetime(propName: PropName, value: Date): void {
    if (!(value instanceof Date) || Number.isNaN(value.getTime()))
      throw new ValueError(`property requires <type 'datetime.datetime'> object, got ${typeName(value)}`);
    const element = this._getOrAdd(propName);
    element.text = formatW3C(value);
    if (propName === "created" || propName === "modified") {
      // These two require an explicit 'xsi:type="dcterms:W3CDTF"' attribute. Setting (then
      // removing) a dummy xsi attribute on the root declares the xsi namespace there rather
      // than on each child element that references it.
      this.set(qn("xsi:foo"), "bar");
      element.set(qn("xsi:type"), "dcterms:W3CDTF");
      this.attrib.delete(qn("xsi:foo"));
    }
  }

  /** Set string value of `propName` property to `value`. */
  _setElementText(propName: PropName, value: unknown): void {
    const s = String(value);
    if ([...s].length > 255) throw new ValueError(`exceeded 255 char limit for property, got:\n\n'${s}'`);
    this._getOrAdd(propName).text = s;
  }

  _textOfElement(propertyName: PropName): string {
    return this[propertyName]?.text ?? "";
  }
}

const pad = (n: number, w = 2) => String(n).padStart(w, "0");

/** python `dt.strftime("%Y-%m-%dT%H:%M:%SZ")` of a naive (UTC-held) datetime. */
function formatW3C(d: Date): string {
  return (
    `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`
  );
}

/** python's strptime with templates "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%Y-%m", "%Y" (last match wins). */
function strptimeW3C(s: string): Date | null {
  const n = "(\\d{1,2})";
  const patterns: [RegExp, number][] = [
    [new RegExp(`^(\\d{4})-${n}-${n}T${n}:${n}:${n}$`), 6],
    [new RegExp(`^(\\d{4})-${n}-${n}$`), 3],
    [new RegExp(`^(\\d{4})-${n}$`), 2],
    [/^(\d{4})$/, 1],
  ];
  let result: Date | null = null;
  for (const [re] of patterns) {
    const m = re.exec(s);
    if (!m) continue;
    const [y, mo = 1, d = 1, h = 0, mi = 0, sec = 0] = m.slice(1).map((x) => parseInt(x, 10));
    if (mo < 1 || mo > 12 || d < 1 || d > daysIn(y, mo) || h > 23 || mi > 59 || sec > 61) continue;
    const date = new Date(Date.UTC(y, mo - 1, d, h, mi, Math.min(sec, 59)));
    date.setUTCFullYear(y);
    result = date;
  }
  return result;
}

function daysIn(y: number, m: number): number {
  const leap = y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0);
  return m === 2 ? (leap ? 29 : 28) : [31, 0, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}

function typeName(v: unknown): string {
  if (v === null || v === undefined) return "<class 'NoneType'>";
  if (typeof v === "string") return "<class 'str'>";
  if (typeof v === "number") return Number.isInteger(v) ? "<class 'int'>" : "<class 'float'>";
  return `<class '${(v as object).constructor?.name ?? typeof v}'>`;
}
