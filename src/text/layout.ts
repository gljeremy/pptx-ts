// Objects related to layout of rendered text, such as TextFitter (`pptx.text.layout`).
import { readFileBytes } from "../fileio.js";
import { PY_GETITEM, PY_LEN } from "../oxml/xml.js";
import { TrueTypeFont } from "./ttf.js";

/** A font file: a path (Node) or the font file's bytes. */
export type FontFile = string | Uint8Array;

/** Value object that knows how to fit text into given rectangular extents. */
export class TextFitter {
  declare readonly 0: _LineSource;
  declare readonly 1: number;
  declare readonly 2: number;
  declare readonly 3: FontFile;

  constructor(lineSource: _LineSource, extents: readonly [number, number], fontFile: FontFile) {
    const [width, height] = extents;
    Object.defineProperties(this, { 0: { value: lineSource }, 1: { value: width }, 2: { value: height }, 3: { value: fontFile } });
  }

  /**
   * The largest whole-number point size less than or equal to `maxSize` that allows `text` to
   * fit completely within `extents` when rendered using the font defined in `fontFile`.
   */
  static bestFitFontSize(text: string, extents: readonly [number, number], maxSize: number, fontFile: FontFile): number {
    const lineSource = new _LineSource(text);
    const textFitter = new this(lineSource, extents, fontFile);
    return textFitter._bestFitFontSize(maxSize);
  }

  /** The largest whole-number point size less than or equal to `maxSize` that this fitter can fit. */
  _bestFitFontSize(maxSize: number): number {
    const predicate = this._fitsInsidePredicate;
    const sizes = _BinarySearchTree.fromOrderedSequence(Array.from({ length: Math.trunc(maxSize) }, (_, i) => i + 1));
    return sizes.findMax(predicate);
  }

  /**
   * [line, remainder] where `line` is the longest line in `lineSource` that will fit in this
   * fitter's width and `remainder` is a `_LineSource` containing the text after the break.
   */
  _breakLine(lineSource: _LineSource, pointSize: number): _Line {
    const lines = _BinarySearchTree.fromOrderedSequence(lineSource);
    const predicate = this._fitsInWidthPredicate(pointSize);
    return lines.findMax(predicate);
  }

  /** A function returning true if a line of text fits this fitter's width at `pointSize`. */
  _fitsInWidthPredicate(pointSize: number): (line: _Line) => boolean {
    return (line: _Line) => {
      const cx = _renderedSize(line.text, pointSize, this._fontFile)[0];
      return cx <= this._width;
    };
  }

  /**
   * A function taking an integer point size, returning true if the text in this fitter can be
   * wrapped to fit entirely within its extents at that point size.
   */
  get _fitsInsidePredicate(): (pointSize: number) => boolean {
    return (pointSize: number) => {
      const textLines = this._wrapLines(this._lineSource, pointSize);
      const cy = _renderedSize("Ty", pointSize, this._fontFile)[1];
      return cy * textLines.length <= this._height;
    };
  }

  get _fontFile(): FontFile {
    return this[3];
  }

  get _height(): number {
    return this[2];
  }

  get _lineSource(): _LineSource {
    return this[0];
  }

  get _width(): number {
    return this[1];
  }

  /** The text in `lineSource` wrapped within this fitter when rendered at `pointSize`. */
  _wrapLines(lineSource: _LineSource, pointSize: number): string[] {
    const [text, remainder] = this._breakLine(lineSource, pointSize);
    const lines = [text];
    if (remainder.truthy) lines.push(...this._wrapLines(remainder, pointSize));
    return lines;
  }

  [PY_GETITEM](i: number): any {
    return [this[0], this[1], this[2], this[3]].at(i);
  }
}

/** A node in a binary search tree. Uniform for root, subtree root, and leaf nodes. */
export class _BinarySearchTree<T> {
  declare _value: T;
  declare _lesser: _BinarySearchTree<T> | null;
  declare _greater: _BinarySearchTree<T> | null;

  constructor(value: T) {
    this._value = value;
    this._lesser = null;
    this._greater = null;
  }

  /** The largest item in or under this node that satisfies `predicate`. */
  findMax(predicate: (v: T) => boolean, max: any = null): any {
    let nextNode: _BinarySearchTree<T> | null;
    if (predicate(this.value)) {
      max = this.value;
      nextNode = this._greater;
    } else {
      nextNode = this._lesser;
    }
    if (nextNode === null) return max;
    return nextNode.findMax(predicate, max);
  }

  /** The root of a balanced binary search tree populated with the values in `iseq`. */
  static fromOrderedSequence<T>(iseq: Iterable<T>): _BinarySearchTree<T> {
    const seq = [...iseq];
    // optimize for usually all fits by making longest first
    const bst = new this<T>(seq.pop() as T);
    bst._insertFromOrderedSequence(seq);
    return bst;
  }

  /** Insert a node containing `value` such that this stays a binary search tree. */
  insert(value: T): void {
    const side = lessThan(value, this.value) ? "_lesser" : "_greater";
    const child = this[side];
    if (child === null) this[side] = new _BinarySearchTree(value);
    else child.insert(value);
  }

  /** A string representation of the tree rooted in this node, for debugging. */
  tree(level = 0, prefix = ""): string {
    let text = `${prefix}${(this.value as any).text}\n`;
    prefix = `${"    ".repeat(level)}└── `;
    if (this._lesser) text += this._lesser.tree(level + 1, prefix);
    if (this._greater) text += this._greater.tree(level + 1, prefix);
    return text;
  }

  /** The value object contained in this node. */
  get value(): T {
    return this._value;
  }

  /** [medialValue, greaterValues, lesserValues] obtained by bisecting `seq`. */
  static _bisect<T>(seq: T[]): [T | T[], T[], T[]] {
    if (seq.length === 0) return [[], [], []];
    const midIdx = Math.trunc(seq.length / 2);
    return [seq[midIdx], seq.slice(midIdx + 1), seq.slice(0, midIdx)];
  }

  /** Insert the values in `seq` into this tree such that a balanced tree is produced. */
  _insertFromOrderedSequence(seq: T[]): void {
    if (seq.length === 0) return;
    const [mid, greater, lesser] = _BinarySearchTree._bisect(seq);
    this.insert(mid as T);
    this._insertFromOrderedSequence(greater);
    this._insertFromOrderedSequence(lesser);
  }
}

/** python `a < b` for the values a search tree holds (numbers or `_Line`s). */
function lessThan(a: any, b: any): boolean {
  return a instanceof _Line ? a.lt(b) : a < b;
}

/**
 * Generates all the possible even-word line breaks in a string of text, each as a `_Line`
 * holding the text before the break and the text after it as a `_LineSource`. It is "truthy"
 * (python `bool()`) when it contains text other than whitespace.
 */
export class _LineSource implements Iterable<_Line> {
  declare _text: string;

  constructor(text: string) {
    this._text = text;
  }

  /** python `bool(line_source)`: false for the empty string or whitespace only. */
  get truthy(): boolean {
    return this._text.trim() !== "";
  }

  __bool__(): boolean {
    return this.truthy;
  }

  equals(other: unknown): boolean {
    return other instanceof _LineSource && this._text === other._text;
  }

  /** Generate a `_Line` for each possible even-word line break in this line source. */
  *[Symbol.iterator](): Generator<_Line> {
    const words = this._text.trim() === "" ? [] : this._text.trim().split(/\s+/);
    for (let idx = 1; idx <= words.length; idx++) {
      const lineText = words.slice(0, idx).join(" ");
      const remainderText = words.slice(idx).join(" ");
      yield new _Line(lineText, new _LineSource(remainderText));
    }
  }

  __repr__(): string {
    return `<_LineSource('${this._text}')>`;
  }
}

/**
 * A candidate line broken at an even word boundary from a string of text, and a `_LineSource`
 * containing the text that remains after the line is broken at this spot.
 */
export class _Line implements Iterable<any> {
  declare readonly 0: string;
  declare readonly 1: _LineSource;

  constructor(text: string, remainder: _LineSource) {
    Object.defineProperties(this, { 0: { value: text }, 1: { value: remainder } });
  }

  /** python `>`: compares by length of text. */
  gt(other: _Line): boolean {
    return this.text.length > other.text.length;
  }

  /** python `<` (note: upstream defines it as "not greater than"). */
  lt(other: _Line): boolean {
    return !this.gt(other);
  }

  get length(): number {
    return this.text.length;
  }

  [PY_LEN](): number {
    return this.length;
  }

  [PY_GETITEM](i: number): any {
    return [this[0], this[1]].at(i);
  }

  *[Symbol.iterator](): Generator<any> {
    yield this[0];
    yield this[1];
  }

  /** python tuple `==`: same text and an equal remainder. */
  equals(other: unknown): boolean {
    const o = other as any;
    if (o === null || o === undefined) return false;
    const [t, r] = other instanceof _Line ? [other[0], other[1]] : Array.isArray(o) ? o : [undefined, undefined];
    return t === this[0] && (r === this[1] || (r instanceof _LineSource && r.equals(this[1])));
  }

  __repr__(): string {
    return `'${this.text}' => '${this.remainder}'`;
  }

  get remainder(): _LineSource {
    return this[1];
  }

  get text(): string {
    return this[0];
  }
}

/** A memoizing cache of parsed font files (python's cache of Pillow `ImageFont`s). */
export class _Fonts {
  static fonts = new Map<FontFile, TrueTypeFont>();

  /** The parsed font for `fontPath`; the point size doesn't change what's parsed. */
  static font(fontPath: FontFile, _pointSize: number): TrueTypeFont {
    let f = this.fonts.get(fontPath);
    if (f === undefined) {
      f = new TrueTypeFont(typeof fontPath === "string" ? readFileBytes(fontPath) : fontPath);
      this.fonts.set(fontPath, f);
    }
    return f;
  }
}

/**
 * [width, height] of `text` in English Metric Units (EMU) when rendered at `pointSize` in the
 * font defined in `fontFile` (python-pptx measures with Pillow; see src/text/ttf.ts).
 */
export function _renderedSize(text: string, pointSize: number, fontFile: FontFile): [number, number] {
  const emuPerInch = 914400;
  const pxPerInch = 72.0;
  const font = _Fonts.font(fontFile, pointSize);
  const [left, top, right, bottom] = font.getbbox(text, pointSize);
  const [pxWidth, pxHeight] = [right - left, bottom - top];
  return [Math.trunc((pxWidth / pxPerInch) * emuPerInch), Math.trunc((pxHeight / pxPerInch) * emuPerInch)];
}
