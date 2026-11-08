// A minimal XLSX writer standing in for XlsxWriter (https://github.com/jmcnamara/XlsxWriter,
// SPDX-License-Identifier: BSD-2-Clause, Copyright 2013-2024 John McNamara), which python-pptx
// uses to write the workbook embedded in each chart. It implements the part of XlsxWriter's API
// that `pptx.chart.xlsx` calls -- `Workbook(file, {in_memory: true})`, `add_worksheet()`,
// `add_format({num_format})`, `close()`, and on the worksheet `write()`, `write_column()` and
// `set_column()` -- and writes the same XML as XlsxWriter 3.2 for it: style indexes assigned in
// order of first use, identical formats sharing one, the same shared-string table, row spans,
// dimension, column widths and `%.16G` numbers. Only docProps/core.xml's timestamps (and the
// zip container's bytes) differ.
//
// Deliberate differences from XlsxWriter's `write()`: strings that start with "=" (formulas),
// that look like URLs (hyperlinks), or that are wrapped in "<r>...</r>" (raw rich text) are
// written as plain text. A chart label like "=Revenue" would otherwise become a formula Excel
// shows as #NAME?, and a "<r>" label could make the workbook invalid.
import { zipSync, type Zippable } from "fflate";
import { encodeUtf8 } from "../oxml/xml.js";
import { pyFloatStr } from "../util.js";
import { THEME_XML } from "./theme.js";

const XLS_ROWMAX = 1048576;
const XLS_COLMAX = 16384;
const XLS_STRMAX = 32767;
const DEFAULT_COL_WIDTH = 8.43;

/** Anything with a python file-like `write(bytes)`, such as pptx-ts's `BytesIO`. */
export interface WritableFile {
  write(data: Uint8Array): unknown;
}

/** A cell format. Only `num_format` matters to chart workbooks. */
export class Format {
  /** index in cellXfs, assigned on first use (XlsxWriter `_get_xf_index()`) */
  _xfIndex: number | null;
  readonly numFormat: string | number;
  _numFormatIndex = 0;

  constructor(
    private readonly _workbook: Workbook,
    properties: { num_format?: string | number; xf_index?: number } = {},
  ) {
    this.numFormat = properties.num_format ?? "General";
    this._xfIndex = properties.xf_index ?? null;
  }

  /** XlsxWriter's format key: formats that differ only in object identity share an xf. */
  get _key(): string {
    return String(this.numFormat);
  }

  _getXfIndex(): number {
    if (this._xfIndex !== null) return this._xfIndex;
    const indices = this._workbook._xfFormatIndices;
    const existing = indices.get(this._key);
    if (existing !== undefined) return existing;
    // +1: Excel has an implicit "General" format at index 0
    const index = 1 + indices.size;
    indices.set(this._key, index);
    this._xfIndex = index;
    return index;
  }
}

type Cell =
  | { kind: "string"; index: number; format: Format | null }
  | { kind: "number"; value: number; format: Format | null }
  | { kind: "boolean"; value: boolean; format: Format | null }
  | { kind: "blank"; format: Format };

/** A worksheet: a sparse table of cells plus column settings. */
export class Worksheet {
  readonly name: string;
  readonly _table = new Map<number, Map<number, Cell>>();
  readonly _colInfo = new Map<number, { width: number | null; format: Format | null }>();
  _dimRowMin: number | null = null;
  _dimRowMax: number | null = null;
  _dimColMin: number | null = null;
  _dimColMax: number | null = null;

  constructor(
    private readonly _workbook: Workbook,
    name: string,
  ) {
    this.name = name;
  }

  /**
   * Write `token` at (`row`, `col`) by its type, as XlsxWriter's `write()`: null or "" is a
   * blank (written only when it has a format), a number a number, a `Date` a date number, and
   * anything else its string. Returns 0, or -1 when the cell is outside the sheet.
   */
  write(row: number, col: number, token: unknown, format: Format | null = null): number {
    if (token === null || token === undefined) return this._writeBlank(row, col, format);
    if (typeof token === "boolean") return this._writeBoolean(row, col, token, format);
    if (typeof token === "number" || token instanceof Number) return this._writeNumber(row, col, Number(token), format);
    if (token instanceof Date) return this._writeNumber(row, col, excelDateNumber(token), format);
    const s = String(token);
    if (s === "") return this._writeBlank(row, col, format);
    return this._writeString(row, col, s, format);
  }

  /** Write each item of `data` down the column from (`row`, `col`). */
  writeColumn(row: number, col: number, data: Iterable<unknown>, format: Format | null = null): number {
    for (const token of data) {
      const error = this.write(row, col, token, format);
      if (error) return error;
      row += 1;
    }
    return 0;
  }

  /** Set the width (in character units) and format of columns `first`..`last`. */
  setColumn(first: number, last: number, width: number | null = null, format: Format | null = null): number {
    if (first > last) [first, last] = [last, first];
    if (first < 0 || last >= XLS_COLMAX) return -1;
    for (let col = first; col <= last; col++) this._colInfo.set(col, { width, format });
    return 0;
  }

  _checkDimensions(row: number, col: number): boolean {
    if (row < 0 || col < 0 || row >= XLS_ROWMAX || col >= XLS_COLMAX) return false;
    if (this._dimRowMin === null || row < this._dimRowMin) this._dimRowMin = row;
    if (this._dimRowMax === null || row > this._dimRowMax) this._dimRowMax = row;
    if (this._dimColMin === null || col < this._dimColMin) this._dimColMin = col;
    if (this._dimColMax === null || col > this._dimColMax) this._dimColMax = col;
    return true;
  }

  _setCell(row: number, col: number, cell: Cell): void {
    let r = this._table.get(row);
    if (r === undefined) this._table.set(row, (r = new Map()));
    r.set(col, cell);
  }

  _writeString(row: number, col: number, s: string, format: Format | null): number {
    if (!this._checkDimensions(row, col)) return -1;
    let error = 0;
    if (s.length > XLS_STRMAX) {
      s = s.slice(0, XLS_STRMAX);
      error = -2;
    }
    this._setCell(row, col, { kind: "string", index: this._workbook._strTable.indexOf(s), format });
    return error;
  }

  _writeNumber(row: number, col: number, value: number, format: Format | null): number {
    if (!Number.isFinite(value)) {
      throw new TypeError("NAN/INF not supported in write_number() without 'nan_inf_to_errors' Workbook() option");
    }
    if (!this._checkDimensions(row, col)) return -1;
    this._setCell(row, col, { kind: "number", value, format });
    return 0;
  }

  _writeBoolean(row: number, col: number, value: boolean, format: Format | null): number {
    if (!this._checkDimensions(row, col)) return -1;
    this._setCell(row, col, { kind: "boolean", value, format });
    return 0;
  }

  _writeBlank(row: number, col: number, format: Format | null): number {
    if (format === null) return 0;
    if (!this._checkDimensions(row, col)) return -1;
    this._setCell(row, col, { kind: "blank", format });
    return 0;
  }

  /** xl/worksheets/sheetN.xml */
  _xml(selected: boolean): string {
    const parts: string[] = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      `<dimension ref="${this._dimensionRef()}"/>`,
      `<sheetViews><sheetView ${selected ? 'tabSelected="1" ' : ""}workbookViewId="0"/></sheetViews>`,
      '<sheetFormatPr defaultRowHeight="15"/>',
    ];
    this._writeCols(parts);
    if (this._dimRowMin === null) parts.push("<sheetData/>");
    else {
      parts.push("<sheetData>");
      this._writeRows(parts);
      parts.push("</sheetData>");
    }
    parts.push('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
    return parts.join("");
  }

  _dimensionRef(): string {
    const [r0, r1, c0, c1] = [this._dimRowMin, this._dimRowMax, this._dimColMin, this._dimColMax];
    if (r0 === null || c0 === null) return "A1";
    if (r0 === r1 && c0 === c1) return cellRef(r0, c0);
    return `${cellRef(r0, c0)}:${cellRef(r1!, c1!)}`;
  }

  /** `<cols>`: adjacent columns with the same settings share one `<col>` range. */
  _writeCols(parts: string[]): void {
    if (this._colInfo.size === 0) return;
    parts.push("<cols>");
    const cols = [...this._colInfo.keys()].sort((a, b) => a - b);
    const same = (a: { width: number | null; format: Format | null }, b: typeof a) => a.width === b.width && a.format === b.format;
    let first = cols[0];
    let last = first;
    let prev = this._colInfo.get(first)!;
    for (const col of cols.slice(1)) {
      const info = this._colInfo.get(col)!;
      if (col === last + 1 && same(info, prev)) last = col;
      else {
        parts.push(colInfoXml(first, last, prev));
        first = last = col;
        prev = info;
      }
    }
    parts.push(colInfoXml(first, last, prev));
    parts.push("</cols>");
  }

  _writeRows(parts: string[]): void {
    const spans = this._calculateSpans();
    for (let row = this._dimRowMin!; row <= this._dimRowMax!; row++) {
      const cells = this._table.get(row);
      if (cells === undefined || cells.size === 0) continue;
      const span = spans.get(Math.floor(row / 16));
      parts.push(`<row r="${row + 1}"${span !== undefined ? ` spans="${span}"` : ""}>`);
      for (let col = this._dimColMin!; col <= this._dimColMax!; col++) {
        const cell = cells.get(col);
        if (cell !== undefined) parts.push(this._cellXml(row, col, cell));
      }
      parts.push("</row>");
    }
  }

  /** "min:max" (1-based) column span of each block of 16 rows, keyed by block. */
  _calculateSpans(): Map<number, string> {
    const spans = new Map<number, string>();
    let min: number | null = null;
    let max = 0;
    for (let row = this._dimRowMin!; row <= this._dimRowMax!; row++) {
      const cells = this._table.get(row);
      if (cells !== undefined) {
        for (let col = this._dimColMin!; col <= this._dimColMax!; col++) {
          if (!cells.has(col)) continue;
          if (min === null) min = max = col;
          else {
            if (col < min) min = col;
            if (col > max) max = col;
          }
        }
      }
      if ((row + 1) % 16 === 0 || row === this._dimRowMax) {
        if (min !== null) {
          spans.set(Math.floor(row / 16), `${min + 1}:${max + 1}`);
          min = null;
        }
      }
    }
    return spans;
  }

  _cellXml(row: number, col: number, cell: Cell): string {
    let attrs = ` r="${cellRef(row, col)}"`;
    if (cell.format !== null) attrs += ` s="${cell.format._getXfIndex()}"`;
    else {
      const colFormat = this._colInfo.get(col)?.format;
      if (colFormat) attrs += ` s="${colFormat._getXfIndex()}"`;
    }
    switch (cell.kind) {
      case "string":
        return `<c${attrs} t="s"><v>${cell.index}</v></c>`;
      case "number":
        return `<c${attrs}><v>${formatG16(cell.value)}</v></c>`;
      case "boolean":
        return `<c${attrs} t="b"><v>${cell.value ? 1 : 0}</v></c>`;
      case "blank":
        return `<c${attrs}/>`;
    }
  }
}

/** Shared-string table (XlsxWriter `SharedStringTable`). */
class SharedStrings {
  count = 0;
  readonly _index = new Map<string, number>();

  indexOf(s: string): number {
    this.count += 1;
    let i = this._index.get(s);
    if (i === undefined) this._index.set(s, (i = this._index.size));
    return i;
  }

  get uniqueCount(): number {
    return this._index.size;
  }

  _xml(): string {
    const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${ns}" count="${this.count}" uniqueCount="${this.uniqueCount}">`;
    for (const s of this._index.keys()) {
      const t = escapeControlCharacters(s);
      const preserve = /^\s/.test(t) || /\s$/.test(t) ? ' xml:space="preserve"' : "";
      xml += `<si><t${preserve}>${escapeData(t)}</t></si>`;
    }
    return xml + "</sst>";
  }
}

/** An XLSX workbook with XlsxWriter's API (camelCased) for what python-pptx uses. */
export class Workbook {
  readonly _file: WritableFile;
  readonly _formats: Format[] = [];
  readonly _xfFormatIndices = new Map<string, number>();
  readonly _strTable = new SharedStrings();
  readonly _worksheets: Worksheet[] = [];
  /** creation time written to docProps/core.xml (XlsxWriter uses the current time) */
  created: Date = new Date();
  _closed = false;

  /** `options` are accepted for API compatibility (`{in_memory: true}`); nothing uses them. */
  constructor(file: WritableFile, _options: Record<string, unknown> = {}) {
    this._file = file;
    // the default format, xf index 0
    this._formats.push(new Format(this, { xf_index: 0 }));
  }

  addWorksheet(name?: string): Worksheet {
    const sheet = new Worksheet(this, name ?? `Sheet${this._worksheets.length + 1}`);
    this._worksheets.push(sheet);
    return sheet;
  }

  addFormat(properties: { num_format?: string | number } = {}): Format {
    const format = new Format(this, properties);
    this._formats.push(format);
    return format;
  }

  /** Write the workbook to the file given to the constructor. */
  close(): void {
    if (this._closed) return;
    this._closed = true;
    if (this._worksheets.length === 0) this.addWorksheet();
    const files: Record<string, string> = {};
    const hasStrings = this._strTable.count > 0;
    // worksheets first: writing their cells assigns the style indexes the styles part lists
    const sheets = this._worksheets.map((ws, i) => ws._xml(i === 0));
    files["[Content_Types].xml"] = this._contentTypesXml(hasStrings);
    files["_rels/.rels"] = ROOT_RELS_XML;
    files["xl/_rels/workbook.xml.rels"] = this._workbookRelsXml(hasStrings);
    sheets.forEach((xml, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = xml));
    files["xl/workbook.xml"] = this._workbookXml();
    if (hasStrings) files["xl/sharedStrings.xml"] = this._strTable._xml();
    files["xl/styles.xml"] = this._stylesXml();
    files["xl/theme/theme1.xml"] = THEME_XML;
    files["docProps/core.xml"] = this._coreXml();
    files["docProps/app.xml"] = this._appXml();
    const zippable: Zippable = {};
    for (const [name, xml] of Object.entries(files)) zippable[name] = [encodeUtf8(xml), { level: 6 }];
    this._file.write(zipSync(zippable, { mtime: new Date(1980, 0, 1) }));
  }

  _contentTypesXml(hasStrings: boolean): string {
    const ct = "application/vnd.openxmlformats-officedocument";
    const sheets = this._worksheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${ct}.spreadsheetml.worksheet+xml"/>`).join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/docProps/app.xml" ContentType="${ct}.extended-properties+xml"/>` +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      `<Override PartName="/xl/styles.xml" ContentType="${ct}.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/xl/theme/theme1.xml" ContentType="${ct}.theme+xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="${ct}.spreadsheetml.sheet.main+xml"/>` +
      sheets +
      (hasStrings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${ct}.spreadsheetml.sharedStrings+xml"/>` : "") +
      "</Types>"
    );
  }

  _workbookRelsXml(hasStrings: boolean): string {
    const rt = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const n = this._worksheets.length;
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    this._worksheets.forEach((_, i) => (xml += `<Relationship Id="rId${i + 1}" Type="${rt}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`));
    xml += `<Relationship Id="rId${n + 1}" Type="${rt}/theme" Target="theme/theme1.xml"/>`;
    xml += `<Relationship Id="rId${n + 2}" Type="${rt}/styles" Target="styles.xml"/>`;
    if (hasStrings) xml += `<Relationship Id="rId${n + 3}" Type="${rt}/sharedStrings" Target="sharedStrings.xml"/>`;
    return xml + "</Relationships>";
  }

  _workbookXml(): string {
    const sheets = this._worksheets.map((ws, i) => `<sheet name="${escapeAttr(ws.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr defaultThemeVersion="124226"/>' +
      '<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>' +
      `<sheets>${sheets}</sheets><calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>`
    );
  }

  /** xl/styles.xml: the formats that were used in a cell, in xf-index order. */
  _stylesXml(): string {
    const xfFormats: Format[] = [];
    for (const f of this._formats) if (f._xfIndex !== null) xfFormats[f._xfIndex] = f;
    // number formats: "General" is 0 and "0" is 1; other strings get ids from 164 in order
    const unique = new Map<string, number>();
    const numFormats: string[] = [];
    for (const f of xfFormats) {
      if (f === undefined) continue;
      const nf = f.numFormat;
      if (typeof nf === "number") f._numFormatIndex = nf === 0 ? 1 : Math.trunc(nf);
      else if (nf === "0") f._numFormatIndex = 1;
      else if (nf === "General") f._numFormatIndex = 0;
      else if (unique.has(nf)) f._numFormatIndex = unique.get(nf)!;
      else {
        const index = 164 + unique.size;
        unique.set(nf, index);
        f._numFormatIndex = index;
        if (f._xfIndex) numFormats.push(nf);
      }
    }
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    if (numFormats.length) {
      xml += `<numFmts count="${numFormats.length}">`;
      numFormats.forEach((code, i) => (xml += `<numFmt numFmtId="${164 + i}" formatCode="${escapeAttr(code)}"/>`));
      xml += "</numFmts>";
    }
    xml +=
      '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    const used = xfFormats.filter((f) => f !== undefined);
    xml += `<cellXfs count="${used.length}">`;
    for (const f of used) {
      xml += `<xf numFmtId="${f._numFormatIndex}" fontId="0" fillId="0" borderId="0" xfId="0"${f._numFormatIndex > 0 ? ' applyNumberFormat="1"' : ""}/>`;
    }
    xml +=
      "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
    return xml;
  }

  _coreXml(): string {
    const t = this.created.toISOString().replace(/\.\d{3}Z$/, "Z");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified></cp:coreProperties>`
    );
  }

  _appXml(): string {
    const n = this._worksheets.length;
    const names = this._worksheets.map((ws) => `<vt:lpstr>${escapeData(ws.name)}</vt:lpstr>`).join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>" +
      `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${n}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
      `<TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${names}</vt:vector></TitlesOfParts>` +
      "<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>"
    );
  }
}

const ROOT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
  "</Relationships>";

/** `<col>` for columns `min`..`max` (0-based), with XlsxWriter's character-to-pixel width. */
function colInfoXml(min: number, max: number, info: { width: number | null; format: Format | null }): string {
  let width = info.width;
  let customWidth = 1;
  if (width === null) {
    width = DEFAULT_COL_WIDTH;
    customWidth = 0;
  } else if (width === DEFAULT_COL_WIDTH) customWidth = 0;
  if (width > 0) {
    const maxDigitWidth = 7;
    const padding = 5;
    width =
      width < 1
        ? Math.trunc((Math.trunc(width * (maxDigitWidth + padding) + 0.5) / maxDigitWidth) * 256) / 256
        : Math.trunc(((Math.trunc(width * maxDigitWidth + 0.5) + padding) / maxDigitWidth) * 256) / 256;
  }
  // python writes the float with str(): "10.7109375", "9.140625"
  let xml = `<col min="${min + 1}" max="${max + 1}" width="${pyFloatStr(width)}"`;
  const xf = info.format ? info.format._getXfIndex() : 0;
  if (xf) xml += ` style="${xf}"`;
  if (customWidth) xml += ' customWidth="1"';
  return xml + "/>";
}

/** "A1"-style reference for 0-based `row`, `col`. */
export function cellRef(row: number, col: number): string {
  let name = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `${name}${row + 1}`;
}

/**
 * Excel serial date number of a naive date held as a UTC `Date` (XlsxWriter
 * `datetime_to_excel_datetime`, 1900 epoch): days since 1899-12-31 plus the fraction of the day,
 * counting Excel's phantom 1900-02-29. A time on 1900-01-01 counts from day 0, as python's
 * `datetime` does there; a midnight value is taken as a `date`.
 */
export function excelDateNumber(d: Date): number {
  const epoch = Date.UTC(1899, 11, 31);
  let t = (d.getTime() - epoch) / 86400000;
  const midnight = d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && d.getUTCMilliseconds() === 0;
  if (!midnight && d.getUTCFullYear() === 1900 && d.getUTCMonth() === 0 && d.getUTCDate() === 1) t -= 1;
  if (t > 59) t += 1;
  return t;
}

/** python `"%.16G" % x`. */
export function formatG16(x: number): string {
  return formatG(x, 16);
}

/** python `"%.{p}G" % x` (`"%g"` when `p` is 6): exponent form outside 1e-4 <= |x| < 10^p. */
export function formatG(x: number, p = 6): string {
  if (x === 0) return Object.is(x, -0) ? "-0" : "0";
  const exp = Math.floor(Math.log10(Math.abs(Number(x.toPrecision(p)))));
  if (exp < -4 || exp >= p) {
    const [m, e] = x.toExponential(p - 1).split("e");
    const mant = m.includes(".") ? m.replace(/\.?0+$/, "") : m;
    const n = Number(e);
    return `${mant}E${n < 0 ? "-" : "+"}${String(Math.abs(n)).padStart(2, "0")}`;
  }
  const s = x.toFixed(Math.max(0, p - 1 - exp));
  return s.includes(".") ? s.replace(/\.?0+$/, "") : s;
}

function escapeData(s: string): string {
  return /[&<>]/.test(s) ? s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;") : s;
}

function escapeAttr(s: string): string {
  return /["&<>\n]/.test(s) ? s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "&#xA;") : s;
}

/** Excel's `_xHHHH_` escapes for control characters (and for literal `_xHHHH_` text). */
function escapeControlCharacters(s: string): string {
  return s
    .replace(/(_x[0-9a-fA-F]{4}_)/g, "_x005F$1")
    .replace(/[\x00-\x08\x0b-\x1f]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`)
    .replace(/￾/g, "_xFFFE_")
    .replace(/￿/g, "_xFFFF_");
}
