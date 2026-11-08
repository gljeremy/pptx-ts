// Custom element classes for table-related XML elements (`pptx.oxml.table`)
import { ValueError } from "../errors.js";
import { defineChildren } from "./xmlchemy.js";
import { lazy, pyInt, pyPercent } from "../util.js";
import { MSO_VERTICAL_ANCHOR } from "../enum/text.js";
import { parseXml } from "./parser.js";
import { CT_GradientFillProperties } from "./dml/fill.js";
import { nsdecls } from "./ns.js";
import { ST_Coordinate, ST_Coordinate32, XsdBoolean, XsdInt } from "./simpletypes.js";
import { CT_TextBody } from "./text.js";
import { BaseOxmlElement, Choice, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice } from "./xmlchemy.js";
import { Emu } from "../util.js";
import type { Length } from "../util.js";

/** `a:tbl` custom element class */
export class CT_Table extends BaseOxmlElement {
  declare getOrAddTblPr: (...args: any[]) => CT_TableProperties;
  declare trLst: CT_TableRow[];
  declare _addTr: (...args: any[]) => CT_TableRow;
  declare tblPr: CT_TableProperties | null;
  declare _addTblPr: (attrs?: Record<string, any>) => CT_TableProperties;
  declare _insertTblPr: (child: CT_TableProperties) => CT_TableProperties;
  declare _newTblPr: () => CT_TableProperties;
  declare _removeTblPr: () => void;
  declare tblGrid: CT_TableGrid;
  declare _insertTr: (child: any) => any;
  declare _newTr: () => any;
  static {
    const _tagSeq = ["a:tblPr", "a:tblGrid", "a:tr"] as const;
    defineChildren(this, {
      tblPr: new ZeroOrOne("a:tblPr", _tagSeq.slice(1)),
      tblGrid: new OneAndOnlyOne("a:tblGrid"),
      tr: new ZeroOrMore("a:tr", _tagSeq.slice(3)),
    });
  }

  /** Return a newly created `a:tr` child element having its `h` attribute set to `height`. */
  addTr(height: number): CT_TableRow {
    return this._addTr({ h: height });
  }

  get bandCol(): boolean {
    return this._getBooleanProperty("bandCol");
  }

  set bandCol(value: boolean) {
    this._setBooleanProperty("bandCol", value);
  }

  get bandRow(): boolean {
    return this._getBooleanProperty("bandRow");
  }

  set bandRow(value: boolean) {
    this._setBooleanProperty("bandRow", value);
  }

  get firstCol(): boolean {
    return this._getBooleanProperty("firstCol");
  }

  set firstCol(value: boolean) {
    this._setBooleanProperty("firstCol", value);
  }

  get firstRow(): boolean {
    return this._getBooleanProperty("firstRow");
  }

  set firstRow(value: boolean) {
    this._setBooleanProperty("firstRow", value);
  }

  /**
   * Generate each `a:tc` element in this tbl.
   *
   * `a:tc` elements are generated left-to-right, top-to-bottom.
   */
  *iterTcs(): Generator<CT_TableCell> {
    for (const tr of this.trLst) yield* tr.tcLst;
  }

  get lastCol(): boolean {
    return this._getBooleanProperty("lastCol");
  }

  set lastCol(value: boolean) {
    this._setBooleanProperty("lastCol", value);
  }

  get lastRow(): boolean {
    return this._getBooleanProperty("lastRow");
  }

  set lastRow(value: boolean) {
    this._setBooleanProperty("lastRow", value);
  }

  /** Return a new `p:tbl` element tree. */
  static newTbl(rows: number, cols: number, width: number, height: number, tableStyleId: string | null = null): CT_Table {
    if (tableStyleId === null) {
      tableStyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
    }
    const xml = pyPercent(this._tblTmpl(), [tableStyleId]);
    const tbl = parseXml(xml);
    let rowheight = Math.floor(height / rows);
    let colwidth = Math.floor(width / cols);
    for (let col = 0; col < cols; col++) {
      if (col === (cols - 1)) {
        colwidth = (width - ((cols - 1) * colwidth));
      }
      tbl.tblGrid.addGridCol(Emu(colwidth));
    }
    for (let row = 0; row < rows; row++) {
      if (row === (rows - 1)) {
        rowheight = (height - ((rows - 1) * rowheight));
      }
      const tr = tbl.addTr(Emu(rowheight));
      for (let col = 0; col < cols; col++) {
        tr.addTc();
      }
    }
    return tbl;
  }

  /** Return `a:tc` element at `row_idx`, `col_idx`. */
  tc(rowIdx: number, colIdx: number): CT_TableCell {
    return this.trLst[rowIdx].tcLst[colIdx];
  }

  /**
   * Generalized getter for the boolean properties on the `a:tblPr` child element.
   *
   * Defaults to False if `propname` attribute is missing or `a:tblPr` element itself is not
   * present.
   */
  _getBooleanProperty(propname: string): boolean {
    const tblPr = this.tblPr;
    if (tblPr === null) {
      return false;
    }
    const propval = (tblPr as any)[propname];
    return propval === true;
  }

  /**
   * Generalized setter for boolean properties on the `a:tblPr` child element.
   *
   * Sets `propname` attribute appropriately based on `value`. If `value` is True, the
   * attribute is set to "1"; a tblPr child element is added if necessary. If `value` is False,
   * the `propname` attribute is removed if present, allowing its default value of False to be
   * its effective value.
   */
  _setBooleanProperty(propname: string, value: boolean): void {
    if (![true, false].includes(value)) {
      throw new ValueError(`assigned value must be either True or False, got ${value}`);
    }
    const tblPr = this.getOrAddTblPr();
    ((tblPr) as any)[propname] = value;
  }

  static _tblTmpl() {
    return `<a:tbl ${nsdecls("a")}>
  <a:tblPr firstRow="1" bandRow="1">
    <a:tableStyleId>${"%s"}</a:tableStyleId>
  </a:tblPr>
  <a:tblGrid/>
</a:tbl>`;
  }
}

/** `a:tc` custom element class */
export class CT_TableCell extends BaseOxmlElement {
  declare getOrAddTcPr: (...args: any[]) => CT_TableCellProperties;
  declare getOrAddTxBody: (...args: any[]) => CT_TextBody;
  declare txBody: CT_TextBody | null;
  declare _addTxBody: (attrs?: Record<string, any>) => CT_TextBody;
  declare _insertTxBody: (child: CT_TextBody) => CT_TextBody;
  declare _removeTxBody: () => void;
  declare tcPr: CT_TableCellProperties | null;
  declare _addTcPr: (attrs?: Record<string, any>) => CT_TableCellProperties;
  declare _insertTcPr: (child: CT_TableCellProperties) => CT_TableCellProperties;
  declare _newTcPr: () => CT_TableCellProperties;
  declare _removeTcPr: () => void;
  declare gridSpan: number;
  declare rowSpan: number;
  declare hMerge: boolean;
  declare vMerge: boolean;
  static {
    const _tagSeq = ["a:txBody", "a:tcPr", "a:extLst"] as const;
    defineChildren(this, {
      txBody: new ZeroOrOne("a:txBody", _tagSeq.slice(1)),
      tcPr: new ZeroOrOne("a:tcPr", _tagSeq.slice(2)),
      gridSpan: new OptionalAttribute("gridSpan", XsdInt, 1),
      rowSpan: new OptionalAttribute("rowSpan", XsdInt, 1),
      hMerge: new OptionalAttribute("hMerge", XsdBoolean, false),
      vMerge: new OptionalAttribute("vMerge", XsdBoolean, false),
    });
  }

  /** String held in `anchor` attribute of `a:tcPr` child element of this `a:tc` element. */
  get anchor(): MSO_VERTICAL_ANCHOR | null {
    if (this.tcPr === null) {
      return null;
    }
    return this.tcPr.anchor;
  }

  set anchor(anchorEnumIdx: MSO_VERTICAL_ANCHOR | null) {
    if ((anchorEnumIdx === null && this.tcPr === null)) {
      return;
    }
    const tcPr = this.getOrAddTcPr();
    tcPr.anchor = anchorEnumIdx;
  }

  /**
   * Append `a:p` elements taken from `spanned_tc`.
   *
   * Any non-empty paragraph elements in `spanned_tc` are removed and appended to the
   * text-frame of this cell. If `spanned_tc` is left with no content after this process, a
   * single empty `a:p` element is added to ensure the cell is compliant with the spec.
   */
  appendPsFrom(spannedTc: CT_TableCell) {
    const sourceTxBody = spannedTc.getOrAddTxBody();
    const targetTxBody = this.getOrAddTxBody();
    if (sourceTxBody.isEmpty) {
      return;
    }
    if (targetTxBody.isEmpty) {
      targetTxBody.clearContent();
    }
    for (const p of sourceTxBody.pLst) {
      targetTxBody.append(p);
    }
    sourceTxBody.unclearContent();
    targetTxBody.unclearContent();
  }

  /** Offset of this cell's column in its table. */
  get colIdx(): number {
    return this.getparent()!.index(this);
  }

  /** True if cell is top-left in merged cell range. */
  get isMergeOrigin(): boolean {
    if ((this.gridSpan > 1 && !this.vMerge)) {
      return true;
    }
    return (this.rowSpan > 1 && !this.hMerge);
  }

  /** True if cell is in merged cell range but not merge origin cell. */
  get isSpanned(): boolean {
    return (this.hMerge || this.vMerge);
  }

  /**
   * Top margin for this cell.
   *
   * This value is stored in the `marT` attribute of the `a:tcPr` child element of this `a:tc`.
   *
   * Read/write. If the attribute is not present, the default value `45720` (0.05 inches) is
   * returned for top and bottom; `91440` (0.10 inches) is the default for left and right.
   * Assigning |None| to any `marX` property clears that attribute from the element,
   * effectively setting it to the default value.
   */
  get marT(): number {
    return this._getMarX("marT", Emu(45720));
  }

  set marT(value: number | null) {
    this._setMarX("marT", value);
  }

  /** Right margin value represented in `marR` attribute. */
  get marR(): number {
    return this._getMarX("marR", Emu(91440));
  }

  set marR(value: number | null) {
    this._setMarX("marR", value);
  }

  /** Bottom margin value represented in `marB` attribute. */
  get marB(): number {
    return this._getMarX("marB", Emu(45720));
  }

  set marB(value: number | null) {
    this._setMarX("marB", value);
  }

  /** Left margin value represented in `marL` attribute. */
  get marL(): number {
    return this._getMarX("marL", Emu(91440));
  }

  set marL(value: number | null) {
    this._setMarX("marL", value);
  }

  /** Return a new `a:tc` element subtree. */
  static new(): CT_TableCell {
    return parseXml(`<a:tc ${nsdecls("a")}>
  <a:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p/>
  </a:txBody>
  <a:tcPr/>
</a:tc>`);
  }

  /** Offset of this cell's row in its table. */
  get rowIdx(): number {
    return (this.getparent() as CT_TableRow).rowIdx;
  }

  /** Table element this cell belongs to. */
  get tbl(): CT_Table {
    return this.xpath("ancestor::a:tbl")[0];
  }

  /** str text contained in cell */
  get text(): string {
    const txBody = this.txBody;
    if (txBody === null) {
      return "";
    }
    return [...txBody.pLst].map((p) => p.text).join("\n");
  }

  /** Generalized method to get margin values. */
  _getMarX(attrName: string, default_: number): number {
    if (this.tcPr === null) {
      return Emu(default_);
    }
    return Emu(pyInt(this.tcPr.get(attrName) ?? default_));
  }

  _newTxBody(): CT_TextBody {
    return CT_TextBody.newATxBody();
  }

  /**
   * Set value of marX attribute on `a:tcPr` child element.
   *
   * If `marX` is |None|, the marX attribute is removed. `marX` is a string, one of `('marL',
   * 'marR', 'marT', 'marB')`.
   */
  _setMarX(marX: string, value: number | null): void {
    if ((value === null && this.tcPr === null)) {
      return;
    }
    const tcPr = this.getOrAddTcPr();
    ((tcPr) as any)[marX] = value;
  }
}

/** `a:tcPr` custom element class */
export class CT_TableCellProperties extends BaseOxmlElement {
  declare egFillProperties: any;
  declare _removeEgFillProperties: () => void;
  declare noFill: any;
  declare _newNoFill: () => any;
  declare _insertNoFill: (child: any) => any;
  declare _addNoFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToNoFill: () => any;
  declare solidFill: any;
  declare _newSolidFill: () => any;
  declare _insertSolidFill: (child: any) => any;
  declare _addSolidFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSolidFill: () => any;
  declare gradFill: any;
  declare _insertGradFill: (child: any) => any;
  declare _addGradFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToGradFill: () => any;
  declare blipFill: any;
  declare _newBlipFill: () => any;
  declare _insertBlipFill: (child: any) => any;
  declare _addBlipFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToBlipFill: () => any;
  declare pattFill: any;
  declare _newPattFill: () => any;
  declare _insertPattFill: (child: any) => any;
  declare _addPattFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToPattFill: () => any;
  declare grpFill: any;
  declare _newGrpFill: () => any;
  declare _insertGrpFill: (child: any) => any;
  declare _addGrpFill: (attrs?: Record<string, any>) => any;
  declare getOrChangeToGrpFill: () => any;
  declare anchor: MSO_VERTICAL_ANCHOR | null;
  declare marL: number | null;
  declare marR: number | null;
  declare marT: number | null;
  declare marB: number | null;
  static {
    defineChildren(this, {
      egFillProperties: new ZeroOrOneChoice([new Choice("a:noFill"), new Choice("a:solidFill"), new Choice("a:gradFill"), new Choice("a:blipFill"), new Choice("a:pattFill"), new Choice("a:grpFill")], ["a:headers", "a:extLst"]),
      anchor: new OptionalAttribute("anchor", MSO_VERTICAL_ANCHOR),
      marL: new OptionalAttribute("marL", ST_Coordinate32),
      marR: new OptionalAttribute("marR", ST_Coordinate32),
      marT: new OptionalAttribute("marT", ST_Coordinate32),
      marB: new OptionalAttribute("marB", ST_Coordinate32),
    });
  }

  _newGradFill() {
    return CT_GradientFillProperties.newGradFill();
  }
}

/** `a:gridCol` custom element class. */
export class CT_TableCol extends BaseOxmlElement {
  declare w: number;
  static {
    defineChildren(this, {
      w: new RequiredAttribute("w", ST_Coordinate),
    });
  }
}

/** `a:tblGrid` custom element class. */
export class CT_TableGrid extends BaseOxmlElement {
  declare gridColLst: CT_TableCol[];
  declare _addGridCol: (...args: any[]) => CT_TableCol;
  declare _insertGridCol: (child: any) => any;
  declare _newGridCol: () => any;
  static {
    defineChildren(this, {
      gridCol: new ZeroOrMore("a:gridCol"),
    });
  }

  /** A newly appended `a:gridCol` child element having its `w` attribute set to `width`. */
  addGridCol(width: number): CT_TableCol {
    return this._addGridCol({ w: width });
  }
}

/** `a:tblPr` custom element class. */
export class CT_TableProperties extends BaseOxmlElement {
  declare bandRow: any;
  declare bandCol: any;
  declare firstRow: any;
  declare firstCol: any;
  declare lastRow: any;
  declare lastCol: any;
  static {
    defineChildren(this, {
      bandRow: new OptionalAttribute("bandRow", XsdBoolean, false),
      bandCol: new OptionalAttribute("bandCol", XsdBoolean, false),
      firstRow: new OptionalAttribute("firstRow", XsdBoolean, false),
      firstCol: new OptionalAttribute("firstCol", XsdBoolean, false),
      lastRow: new OptionalAttribute("lastRow", XsdBoolean, false),
      lastCol: new OptionalAttribute("lastCol", XsdBoolean, false),
    });
  }
}

/** `a:tr` custom element class. */
export class CT_TableRow extends BaseOxmlElement {
  declare tcLst: CT_TableCell[];
  declare _addTc: (...args: any[]) => CT_TableCell;
  declare _insertTc: (child: any) => any;
  declare h: number;
  static {
    defineChildren(this, {
      tc: new ZeroOrMore("a:tc", ["a:extLst"]),
      h: new RequiredAttribute("h", ST_Coordinate),
    });
  }

  /** A newly added minimal valid `a:tc` child element. */
  addTc(): CT_TableCell {
    return this._addTc();
  }

  /** Offset of this row in its table. */
  get rowIdx(): number {
    return (this.getparent() as CT_Table).trLst.indexOf(this);
  }

  _newTc() {
    return CT_TableCell.new();
  }
}

/**
 * A 2D block of `a:tc` cell elements in a table.
 *
 * This object assumes the structure of the underlying table does not change during its lifetime.
 * Structural changes in this context would be insertion or removal of rows or columns.
 *
 * The client is expected to create, use, and then abandon an instance in the context of a single
 * user operation that is known to have no structural side-effects of this type.
 */
export class TcRange {
  declare _tc: any;
  declare _otherTc: any;

  constructor(tc: CT_TableCell, otherTc: CT_TableCell) {
    this._tc = tc;
    this._otherTc = otherTc;
  }

  /** Return instance created from merge-origin tc element. */
  static fromMergeOrigin(tc: CT_TableCell) {
    const otherTc = tc.tbl.tc(((tc.rowIdx + tc.rowSpan) - 1), ((tc.colIdx + tc.gridSpan) - 1));
    return new this(tc, otherTc);
  }

  /** True if one or more cells in range are part of a merged cell. */
  get containsMergedCell(): boolean {
    return lazy(this, "containsMergedCell", () => {
      for (const tc of this.iterTcs()) {
        if (tc.gridSpan > 1) {
          return true;
        }
        if (tc.rowSpan > 1) {
          return true;
        }
        if (tc.hMerge) {
          return true;
        }
        if (tc.vMerge) {
          return true;
        }
      }
      return false;
    });
  }

  /** (row_count, col_count) pair describing size of range. */
  get dimensions(): [number, number] {
    return lazy(this, "dimensions", () => {
      const [, , width, height] = this._extents;
      return [height, width];
    });
  }

  /** True if both cells provided to constructor are in same table. */
  get inSameTable() {
    return lazy(this, "inSameTable", () => {
      if (this._tc.tbl === this._otherTc.tbl) {
        return true;
      }
      return false;
    });
  }

  /** Generate each `a:tc` element not in leftmost column of range. */
  *iterExceptLeftColTcs() {
    for (const tr of this._tbl.trLst.slice(this._top, this._bottom)) {
      for (const tc of tr.tcLst.slice((this._left + 1), this._right)) {
        yield tc;
      }
    }
  }

  /** Generate each `a:tc` element in non-first rows of range. */
  *iterExceptTopRowTcs() {
    for (const tr of this._tbl.trLst.slice((this._top + 1), this._bottom)) {
      for (const tc of tr.tcLst.slice(this._left, this._right)) {
        yield tc;
      }
    }
  }

  /** Generate each `a:tc` element in leftmost column of range. */
  *iterLeftColTcs() {
    const colIdx = this._left;
    for (const tr of this._tbl.trLst.slice(this._top, this._bottom)) {
      yield tr.tcLst[colIdx];
    }
  }

  /**
   * Generate each `a:tc` element in this range.
   *
   * Cell elements are generated left-to-right, top-to-bottom.
   */
  iterTcs() {
    return [...this._tbl.trLst.slice(this._top, this._bottom)].flatMap((tr) => [...tr.tcLst.slice(this._left, this._right)].map((tc) => tc));
  }

  /** Generate each `a:tc` element in topmost row of range. */
  *iterTopRowTcs() {
    const tr = this._tbl.trLst[this._top];
    for (const tc of tr.tcLst.slice(this._left, this._right)) {
      yield tc;
    }
  }

  /** Move all paragraphs in range to origin cell. */
  moveContentToOrigin() {
    const tcs = [...this.iterTcs()];
    const originTc = tcs[0];
    for (const spannedTc of tcs.slice(1)) {
      originTc.appendPsFrom(spannedTc);
    }
  }

  /** Index of row following last row of range */
  get _bottom() {
    return lazy(this, "_bottom", () => {
      const [, top, , height] = this._extents;
      return (top + height);
    });
  }

  /**
   * A (left, top, width, height) tuple describing range extents.
   *
   * Note this is normalized to accommodate the various orderings of the corner cells provided
   * on construction, which may be in any of four configurations such as (top-left,
   * bottom-right), (bottom-left, top-right), etc.
   */
  get _extents(): [number, number, number, number] {
    return lazy(this, "_extents", () => {
      function startAndSize(idx: number, otherIdx: number) {
        return [Math.min(idx, otherIdx), (Math.abs((idx - otherIdx)) + 1)];
      }
      const [tc, otherTc] = [this._tc, this._otherTc];
      const [left, width] = startAndSize(tc.colIdx, otherTc.colIdx);
      const [top, height] = startAndSize(tc.rowIdx, otherTc.rowIdx);
      return [left, top, width, height];
    });
  }

  /** Index of leftmost column in range. */
  get _left() {
    return lazy(this, "_left", () => {
      const [left, , , ] = this._extents;
      return left;
    });
  }

  /** Index of column following the last column in range. */
  get _right() {
    return lazy(this, "_right", () => {
      const [left, , width, ] = this._extents;
      return (left + width);
    });
  }

  /** `a:tbl` element containing this cell range. */
  get _tbl() {
    return lazy(this, "_tbl", () => {
      return this._tc.tbl;
    });
  }

  /** Index of topmost row in range. */
  get _top() {
    return lazy(this, "_top", () => {
      const [, top, , ] = this._extents;
      return top;
    });
  }
}
