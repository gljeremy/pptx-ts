// Objects related to system font file lookup (`pptx.text.fonts`).
//
// `FontFiles.find()` locates an installed .ttf/.otf file by family name and style, for
// `TextFrame.fitText()`. Font directories are searched on macOS and Windows (as python-pptx
// does) and also on Linux; searching needs Node's filesystem, so in browsers pass the font
// file to `fitText()` explicitly.
import { BytesIO, nodeFs, readFileBytes, type ReadableFile } from "../fileio.js";
import { KeyError } from "../errors.js";
import { lazy } from "../util.js";
import { calcsize, unpackFrom } from "./struct.js";

/** OSError: the operating system's font directories are unknown. */
export class OSError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OSError";
  }
}

type FontKey = string;

/**
 * Key for `FontFiles._fontFiles`: python's (family_name, is_bold, is_italic) tuple key, written
 * as `String(tuple)` like other tuple-keyed python mappings in pptx-ts ("Arial,true,false").
 */
export const fontKey = (familyName: string | null, isBold: boolean, isItalic: boolean): FontKey => String([familyName, isBold, isItalic]);

/** python `mapping[key]` for a Map or a dict-like with a raising `get()`: KeyError when missing. */
function getItem<V>(m: Map<string, V> | { get(k: string): V }, key: string): V {
  if (m instanceof Map) {
    if (!m.has(key)) throw new KeyError(key);
    return m.get(key)!;
  }
  return m.get(key);
}

/** python `mapping.get(key)`: undefined when missing. */
function getOrNone<V>(m: Map<string, V> | { get(k: string): V; has(k: string): boolean }, key: string): V | undefined {
  return m.has(key) ? m.get(key) : undefined;
}

const isKeyError = (e: unknown) => e instanceof KeyError || (e instanceof Error && e.name === "KeyError");

/** The platform python-pptx dispatches on (python `sys.platform`). */
export const sys = {
  get platform(): string {
    return (globalThis as any).process?.platform ?? "browser";
  },
};

/** The filesystem helpers this module uses (python `os`). */
export const os = {
  get environ(): Record<string, string | undefined> {
    return (globalThis as any).process?.env ?? {};
  },
  /** python `os.walk(top)`: [root, dirs, files] for `top` and every directory below it. */
  *walk(top: string): Generator<[string, string[], string[]]> {
    const fs = nodeFs() as any;
    if (!fs) return;
    let entries: any[];
    try {
      entries = fs.readdirSync(top, { withFileTypes: true });
    } catch {
      return;
    }
    const dirs = entries.filter((e) => e.isDirectory()).map((e) => e.name as string);
    const files = entries.filter((e) => !e.isDirectory()).map((e) => e.name as string);
    yield [top, dirs, files];
    for (const d of dirs) yield* os.walk(joinPath(top, d));
  },
};

function joinPath(a: string, b: string): string {
  return /[\\/]$/.test(a) ? a + b : `${a}${a.includes("\\") && !a.includes("/") ? "\\" : "/"}${b}`;
}

/** A class-based singleton serving as a lazy cache for system font details. */
export class FontFiles {
  static _fontFiles: Map<FontKey, string> | null = null;

  /**
   * The path to the font file for `familyName` with the given bold and italic styles; KeyError
   * if no installed font matches.
   */
  static find(familyName: string, isBold: boolean, isItalic: boolean): string {
    if (this._fontFiles === null) this._fontFiles = this._installedFonts();
    return getItem(this._fontFiles, fontKey(familyName, isBold, isItalic));
  }

  /** {fontKey(family, bold, italic): path} for each font file in the font directories. */
  static _installedFonts(): Map<FontKey, string> {
    const fonts = new Map<FontKey, string>();
    for (const d of this._fontDirectories()) {
      for (const [key, path] of this._iterFontFilesIn(d)) fonts.set(String(key), path);
    }
    return fonts;
  }

  /** The font directories on this computer; OSError when the platform has none known. */
  static _fontDirectories(): string[] {
    const platform = sys.platform;
    if (platform.startsWith("darwin")) return this._osXFontDirectories();
    if (platform.startsWith("win32")) return this._windowsFontDirectories();
    if (platform.startsWith("linux")) return this._linuxFontDirectories();
    throw new OSError("unsupported operating system");
  }

  /** Generate [fontKey, path] for each .otf/.ttf file in `directory` and its subdirectories. */
  static *_iterFontFilesIn(directory: string): Generator<[[string | null, boolean, boolean], string]> {
    for (const [root, , files] of os.walk(directory)) {
      for (const filename of files) {
        const ext = filename.includes(".") ? filename.slice(filename.lastIndexOf(".")).toLowerCase() : "";
        if (ext !== ".otf" && ext !== ".ttf") continue;
        const path = joinPath(root, filename);
        const f = _Font.open(path).__enter__();
        try {
          yield [[f.familyName, f.isBold, f.isItalic], path];
        } finally {
          f.__exit__();
        }
      }
    }
  }

  /** Font directories on macOS. */
  static _osXFontDirectories(): string[] {
    const dirs = ["/Library/Fonts", "/Network/Library/Fonts", "/System/Library/Fonts"];
    const home = os.environ.HOME;
    if (home !== undefined && home !== null) dirs.push(joinPath(joinPath(home, "Library"), "Fonts"), joinPath(home, ".fonts"));
    return dirs;
  }

  /** Font directories on Windows (the system directory and the per-user one). */
  static _windowsFontDirectories(): string[] {
    return ["C:\\Windows\\Fonts"];
  }

  /** Font directories on Linux (pptx-ts addition; python-pptx raises OSError there). */
  static _linuxFontDirectories(): string[] {
    const dirs = ["/usr/share/fonts", "/usr/local/share/fonts"];
    const home = os.environ.HOME;
    if (home) dirs.push(joinPath(home, ".fonts"), joinPath(joinPath(joinPath(home, ".local"), "share"), "fonts"));
    return dirs;
  }
}

/**
 * A wrapper around an OTF/TTF font file stream that knows how to parse it for its name and
 * style characteristics, e.g. bold and italic.
 */
export class _Font {
  declare _stream: _Stream;

  constructor(stream: _Stream) {
    this._stream = stream;
  }

  /** Context-manager entry (python `with _Font.open(path) as f:`). */
  __enter__(): this {
    return this;
  }

  /** Context-manager exit: closes the underlying stream. */
  __exit__(..._args: unknown[]): void {
    this._stream.close();
  }

  /** True if this font is marked as a bold style of its font family. */
  get isBold(): boolean {
    try {
      return (getItem(this._tables, "head") as _HeadTable).isBold;
    } catch (e) {
      // some files don't have a head table
      if (isKeyError(e)) return false;
      throw e;
    }
  }

  /** True if this font is marked as an italic style of its font family. */
  get isItalic(): boolean {
    try {
      return (getItem(this._tables, "head") as _HeadTable).isItalic;
    } catch (e) {
      if (isKeyError(e)) return false;
      throw e;
    }
  }

  /** A _Font instance loaded from the font file at `fontFilePath`. */
  static open(fontFilePath: string): _Font {
    return new this(_Stream.open(fontFilePath));
  }

  /**
   * The name of the typeface family for this font, e.g. 'Arial'. The full typeface name
   * includes optional style names, such as 'Regular' or 'Bold Italic'. This attribute is only
   * the common base name shared by all fonts in the family.
   */
  get familyName(): string | null {
    return (getItem(this._tables, "name") as _NameTable).familyName;
  }

  /** The 5-tuple of fields from the start of the font file. */
  get _fields(): any[] {
    // sfnt_version, tbl_count, search_range, entry_selector, range_shift
    return lazy(this, "_fields", () => this._stream.readFields(">4sHHHH", 0));
  }

  /** Generate [tag, offset, length] for each of the tables in this font file. */
  *_iterTableRecords(): Generator<[string, number, number]> {
    const count = this._tableCount;
    const bufr = this._stream.read(12, count * 16);
    const tmpl = ">4sLLL";
    for (let i = 0; i < count; i++) {
      const [tag, , off, len] = unpackFrom(tmpl, bufr, i * 16);
      yield [new TextDecoder().decode(tag), off, len];
    }
  }

  /** {tag: table} for the tables in this font file. */
  get _tables(): Map<string, _BaseTable> {
    return lazy(this, "_tables", () => new Map([...this._iterTableRecords()].map(([tag, off, len]) => [tag, _TableFactory(tag, this._stream, off, len)])));
  }

  /** The number of tables in this OpenType font file. */
  get _tableCount(): number {
    return this._fields[1];
  }
}

/** A wrapper providing a few convenience methods for reading the bytes of a font file. */
export class _Stream {
  declare _file: ReadableFile & { seek(offset: number): unknown; close?(): void };

  constructor(file: ReadableFile & { seek(offset: number): unknown; close?(): void }) {
    this._file = file;
  }

  /** A _Stream over the font file at `path` (Node). */
  static open(path: string): _Stream {
    return new this(new BytesIO(readFileBytes(path)));
  }

  /** Close the wrapped file; using the stream after closing raises an exception. */
  close(): void {
    this._file.close?.();
  }

  /** `length` bytes from this stream starting at `offset`. */
  read(offset: number, length: number): Uint8Array {
    this._file.seek(offset);
    return this._file.read(length);
  }

  /** The values read from this stream at `offset` as described by struct `template`. */
  readFields(template: string | Uint8Array, offset = 0): any[] {
    this._file.seek(offset);
    const bufr = this._file.read(calcsize(template));
    return unpackFrom(template, bufr);
  }
}

/** Base class for OpenType font file table objects. */
export class _BaseTable {
  declare _tag: string;
  declare _stream: _Stream;
  declare _offset: number;
  declare _length: number;

  constructor(tag: string, stream: _Stream, offset: number, length: number) {
    this._tag = tag;
    this._stream = stream;
    this._offset = offset;
    this._length = length;
  }
}

/** The 'head' table: header information for the font, including its bold and italic style. */
export class _HeadTable extends _BaseTable {
  /** True if this font is marked as having emboldened characters. */
  get isBold(): boolean {
    return Boolean(this._macStyle & 1);
  }

  /** True if this font is marked as having italicized characters. */
  get isItalic(): boolean {
    return Boolean(this._macStyle & 2);
  }

  /** The fields of this head table, read from the font file. */
  get _fields(): any[] {
    return lazy(this, "_fields", () => this._stream.readFields(">4s4sLLHHqqhhhhHHHHH", this._offset));
  }

  /** The integer value of the macStyle field, whose bits mark bold and italic. */
  get _macStyle(): number {
    return this._fields[12];
  }
}

/** The 'name' table: the name-related strings for the font. */
export class _NameTable extends _BaseTable {
  /** The name of the typeface family for this font, e.g. 'Arial'. */
  get familyName(): string | null {
    // keys for Unicode, Mac, and Windows family name, respectively
    for (const key of ["0,1", "1,1", "3,1"]) {
      const value = getOrNone(this._names, key);
      if (value !== undefined && value !== null) return value;
    }
    return null;
  }

  /** The unicode text of `rawName`, or null for a platform/encoding not supported. */
  _decodeName(rawName: Uint8Array, platformId: number, encodingId: number): string | null {
    if (platformId === 1) {
      // reject non-Roman Mac font names
      if (encodingId !== 0) return null;
      return new TextDecoder("macintosh").decode(rawName);
    }
    if (platformId === 0 || platformId === 3) return decodeUtf16be(rawName);
    return null;
  }

  /** Generate [[platformId, nameId], name] for each name in this table it can decode. */
  *_iterNames(): Generator<[[number, number], string]> {
    const [, count, stringsOffset] = this._tableHeader;
    const tableBytes = this._tableBytes;
    for (let idx = 0; idx < count; idx++) {
      const [platformId, nameId, name] = this._readName(tableBytes, idx, stringsOffset);
      if (name === null) continue;
      yield [[platformId, nameId], name];
    }
  }

  /** The [platformId, encodingId, languageId, nameId, length, nameStrOffset] of a name record. */
  _nameHeader(bufr: Uint8Array, idx: number): number[] {
    return unpackFrom(">HHHHHH", bufr, 6 + idx * 12);
  }

  /** The `length` bytes of the encoded string at `strOffset` in the strings area. */
  _rawNameString(bufr: Uint8Array, stringsOffset: number, strOffset: number, length: number): Uint8Array {
    return unpackFrom(`${length}s`, bufr, stringsOffset + strOffset)[0];
  }

  /** [platformId, nameId, name] for the name at `idx` in `bufr`. */
  _readName(bufr: Uint8Array, idx: number, stringsOffset: number): [number, number, string | null] {
    const [platformId, encId, , nameId, length, strOffset] = this._nameHeader(bufr, idx);
    const name = this._readNameText(bufr, platformId, encId, stringsOffset, strOffset, length);
    return [platformId, nameId, name];
  }

  /** The name string at `nameStrOffset`, or null if decoding its format is not supported. */
  _readNameText(bufr: Uint8Array, platformId: number, encodingId: number, stringsOffset: number, nameStrOffset: number, length: number): string | null {
    const rawName = this._rawNameString(bufr, stringsOffset, nameStrOffset, length);
    return this._decodeName(rawName, platformId, encodingId);
  }

  /** The binary contents of this name table. */
  get _tableBytes(): Uint8Array {
    return lazy(this, "_tableBytes", () => this._stream.read(this._offset, this._length));
  }

  /** [format, count, stringsOffset] from the start of this table. */
  get _tableHeader(): number[] {
    return unpackFrom(">HHH", this._tableBytes);
  }

  /** {"platformId,nameId": name} for the names in this table. */
  get _names(): Map<string, string> {
    // python's tuple keys, as strings
    return lazy(this, "_names", () => new Map([...this._iterNames()].map(([k, v]) => [String(k), v])));
  }
}

function decodeUtf16be(b: Uint8Array): string {
  let s = "";
  for (let i = 0; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]);
  return s;
}

/** The table object for `tag` in a font file, loaded from `length` bytes at `offset`. */
export function _TableFactory(tag: string, stream: _Stream, offset: number, length: number): _BaseTable {
  const TableClass = tag === "head" ? _HeadTable : tag === "name" ? _NameTable : _BaseTable;
  return new TableClass(tag, stream, offset, length);
}
