// Table-related objects such as Table and Cell. (`pptx.table`)
import { PY_GETITEM } from "./oxml/xml.js";
import { ValueError, IndexError } from "./errors.js";
import { lazy, pyPercent } from "./util.js";
import { FillFormat } from "./dml/fill.js";
import { TcRange } from "./oxml/table.js";
import { Subshape } from "./shapes/index.js";
import { TextFrame } from "./text/text.js";
import { Emu } from "./util.js";
import type { MSO_VERTICAL_ANCHOR } from "./enum/text.js";
import type { CT_Table, CT_TableCell, CT_TableCol, CT_TableRow } from "./oxml/table.js";
import type { BaseSlidePart } from "./parts/slide.js";
import type { GraphicFrame } from "./shapes/graphfrm.js";
import type { ProvidesPart } from "./types.js";
import type { Length } from "./util.js";

/**
 * A DrawingML table object.
 *
 * Not intended to be constructed directly, use
 * :meth:`.Slide.shapes.add_table` to add a table to a slide.
 */
export class Table {
  declare _tbl: any;
  declare _graphicFrame: any;

  constructor(tbl: CT_Table, graphicFrame: GraphicFrame) {
    this._tbl = tbl;
    this._graphicFrame = graphicFrame;
  }

  /**
   * Return cell at `row_idx`, `col_idx`.
   *
   * Return value is an instance of |_Cell|. `row_idx` and `col_idx` are zero-based, e.g.
   * cell(0, 0) is the top, left cell in the table.
   */
  cell(rowIdx: number, colIdx: number): _Cell {
    return new _Cell(this._tbl.tc(rowIdx, colIdx), this);
  }

  /**
   * |_ColumnCollection| instance for this table.
   *
   * Provides access to |_Column| objects representing the table's columns. |_Column| objects
   * are accessed using list notation, e.g. `col = tbl.columns[0]`.
   */
  get columns(): _ColumnCollection {
    return lazy(this, "columns", () => {
      return new _ColumnCollection(this._tbl, this);
    });
  }

  /**
   * When `True`, indicates first column should have distinct formatting.
   *
   * Read/write. Distinct formatting is used, for example, when the first column contains row
   * headings (is a side-heading column).
   */
  get firstCol(): boolean {
    return this._tbl.firstCol;
  }

  set firstCol(value: boolean) {
    this._tbl.firstCol = value;
  }

  /**
   * When `True`, indicates first row should have distinct formatting.
   *
   * Read/write. Distinct formatting is used, for example, when the first row contains column
   * headings.
   */
  get firstRow(): boolean {
    return this._tbl.firstRow;
  }

  set firstRow(value: boolean) {
    this._tbl.firstRow = value;
  }

  /**
   * When `True`, indicates rows should have alternating shading.
   *
   * Read/write. Used to allow rows to be traversed more easily without losing track of which
   * row is being read.
   */
  get horzBanding(): boolean {
    return this._tbl.bandRow;
  }

  set horzBanding(value: boolean) {
    this._tbl.bandRow = value;
  }

  /**
   * Generate _Cell object for each cell in this table.
   *
   * Each grid cell is generated in left-to-right, top-to-bottom order.
   */
  *iterCells(): Generator<_Cell> {
    for (const tc of this._tbl.iterTcs()) yield new _Cell(tc, this);
  }

  /**
   * When `True`, indicates the rightmost column should have distinct formatting.
   *
   * Read/write. Used, for example, when a row totals column appears at the far right of the
   * table.
   */
  get lastCol(): boolean {
    return this._tbl.lastCol;
  }

  set lastCol(value: boolean) {
    this._tbl.lastCol = value;
  }

  /**
   * When `True`, indicates the bottom row should have distinct formatting.
   *
   * Read/write. Used, for example, when a totals row appears as the bottom row.
   */
  get lastRow(): boolean {
    return this._tbl.lastRow;
  }

  set lastRow(value: boolean) {
    this._tbl.lastRow = value;
  }

  /**
   * Called by a row when its height changes.
   *
   * Triggers the graphic frame to recalculate its total height (as the sum of the row
   * heights).
   */
  notifyHeightChanged(): void {
    const newTableHeight = Emu([...this.rows].reduce((sum, row) => sum + row.height, 0));
    this._graphicFrame.height = newTableHeight;
  }

  /**
   * Called by a column when its width changes.
   *
   * Triggers the graphic frame to recalculate its total width (as the sum of the column
   * widths).
   */
  notifyWidthChanged(): void {
    const newTableWidth = Emu([...this.columns].reduce((sum, col) => sum + col.width, 0));
    this._graphicFrame.width = newTableWidth;
  }

  /** The package part containing this table. */
  get part(): BaseSlidePart {
    return this._graphicFrame.part;
  }

  /**
   * |_RowCollection| instance for this table.
   *
   * Provides access to |_Row| objects representing the table's rows. |_Row| objects are
   * accessed using list notation, e.g. `col = tbl.rows[0]`.
   */
  get rows() {
    return lazy(this, "rows", () => {
      return new _RowCollection(this._tbl, this);
    });
  }

  /**
   * When `True`, indicates columns should have alternating shading.
   *
   * Read/write. Used to allow columns to be traversed more easily without losing track of
   * which column is being read.
   */
  get vertBanding(): boolean {
    return this._tbl.bandCol;
  }

  set vertBanding(value: boolean) {
    this._tbl.bandCol = value;
  }
}

/** Table cell */
export class _Cell extends Subshape {
  declare _tc: any;

  constructor(tc: CT_TableCell, parent: ProvidesPart) {
    super(parent);
    this._tc = tc;
  }

  /**
   * |True| if this object proxies the same element as `other`.
   *
   * Equality for proxy objects is defined as referring to the same XML element, whether or not
   * they are the same proxy object instance.
   */
  equals(other: unknown): boolean {
    if (!(other instanceof this.constructor)) {
      return false;
    }
    return this._tc === (other as _Cell)._tc;
  }


  /**
   * |FillFormat| instance for this cell.
   *
   * Provides access to fill properties such as foreground color.
   */
  get fill(): FillFormat {
    return lazy(this, "fill", () => {
      const tcPr = this._tc.getOrAddTcPr();
      return FillFormat.fromFillParent(tcPr);
    });
  }

  /** True if this cell is the top-left grid cell in a merged cell. */
  get isMergeOrigin(): boolean {
    return this._tc.isMergeOrigin;
  }

  /**
   * True if this cell is spanned by a merge-origin cell.
   *
   * A merge-origin cell "spans" the other grid cells in its merge range, consuming their area
   * and "shadowing" the spanned grid cells.
   *
   * Note this value is |False| for a merge-origin cell. A merge-origin cell spans other grid
   * cells, but is not itself a spanned cell.
   */
  get isSpanned(): boolean {
    return this._tc.isSpanned;
  }

  /**
   * Left margin of cells.
   *
   * Read/write. If assigned |None|, the default value is used, 0.1 inches for left and right
   * margins and 0.05 inches for top and bottom.
   */
  get marginLeft(): number {
    return this._tc.marL;
  }

  set marginLeft(marginLeft: number | null) {
    (this.constructor as typeof _Cell)._validateMarginValue(marginLeft);
    this._tc.marL = marginLeft;
  }

  /** Right margin of cell. */
  get marginRight(): number {
    return this._tc.marR;
  }

  set marginRight(marginRight: number | null) {
    (this.constructor as typeof _Cell)._validateMarginValue(marginRight);
    this._tc.marR = marginRight;
  }

  /** Top margin of cell. */
  get marginTop(): number {
    return this._tc.marT;
  }

  set marginTop(marginTop: number | null) {
    (this.constructor as typeof _Cell)._validateMarginValue(marginTop);
    this._tc.marT = marginTop;
  }

  /** Bottom margin of cell. */
  get marginBottom(): number {
    return this._tc.marB;
  }

  set marginBottom(marginBottom: number | null) {
    (this.constructor as typeof _Cell)._validateMarginValue(marginBottom);
    this._tc.marB = marginBottom;
  }

  /**
   * Create merged cell from this cell to `other_cell`.
   *
   * This cell and `other_cell` specify opposite corners of the merged cell range. Either
   * diagonal of the cell region may be specified in either order, e.g. self=bottom-right,
   * other_cell=top-left, etc.
   *
   * Raises |ValueError| if the specified range already contains merged cells anywhere within
   * its extents or if `other_cell` is not in the same table as `self`.
   */
  merge(otherCell: _Cell): void {
    let tc;
    const tcRange = new TcRange(this._tc, otherCell._tc);
    if (!tcRange.inSameTable) {
      throw new ValueError("other_cell from different table");
    }
    if (tcRange.containsMergedCell) {
      throw new ValueError("range contains one or more merged cells");
    }
    tcRange.moveContentToOrigin();
    const [rowCount, colCount] = tcRange.dimensions;
    for (const tc of tcRange.iterTopRowTcs()) {
      tc.rowSpan = rowCount;
    }
    for (const tc of tcRange.iterLeftColTcs()) {
      tc.gridSpan = colCount;
    }
    for (const tc of tcRange.iterExceptLeftColTcs()) {
      tc.hMerge = true;
    }
    for (const tc of tcRange.iterExceptTopRowTcs()) {
      tc.vMerge = true;
    }
  }

  /**
   * int count of rows spanned by this cell.
   *
   * The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
   * is not |True|, since only a merge-origin cell contains complete span information. This
   * property is only intended for use on cells known to be a merge origin by testing
   * `.is_merge_origin`.
   */
  get spanHeight(): number {
    return this._tc.rowSpan;
  }

  /**
   * int count of columns spanned by this cell.
   *
   * The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
   * is not |True|, since only a merge-origin cell contains complete span information. This
   * property is only intended for use on cells known to be a merge origin by testing
   * `.is_merge_origin`.
   */
  get spanWidth(): number {
    return this._tc.gridSpan;
  }

  /**
   * Remove merge from this (merge-origin) cell.
   *
   * The merged cell represented by this object will be "unmerged", yielding a separate
   * unmerged cell for each grid cell previously spanned by this merge.
   *
   * Raises |ValueError| when this cell is not a merge-origin cell. Test with
   * `.is_merge_origin` before calling.
   */
  split(): void {
    let tc;
    if (!this.isMergeOrigin) {
      throw new ValueError("not a merge-origin cell; only a merge-origin cell can be split");
    }
    const tcRange = TcRange.fromMergeOrigin(this._tc);
    for (const tc of tcRange.iterTcs()) {
      tc.rowSpan = 1;
      tc.gridSpan = 1;
      tc.hMerge = false;
      tc.vMerge = false;
    }
  }

  /**
   * Textual content of cell as a single string.
   *
   * The returned string will contain a newline character (`"\n"`) separating each paragraph
   * and a vertical-tab (`"\v"`) character for each line break (soft carriage return) in the
   * cell's text.
   *
   * Assignment to `text` replaces all text currently contained in the cell. A newline
   * character (`"\n"`) in the assigned text causes a new paragraph to be started. A
   * vertical-tab (`"\v"`) character in the assigned text causes a line-break (soft
   * carriage-return) to be inserted. (The vertical-tab character appears in clipboard text
   * copied from PowerPoint as its encoding of line-breaks.)
   */
  get text(): string {
    return this.textFrame.text;
  }

  set text(text: string) {
    this.textFrame.text = text;
  }

  /** |TextFrame| containing the text that appears in the cell. */
  get textFrame(): TextFrame {
    const txBody = this._tc.getOrAddTxBody();
    return new TextFrame(txBody, this);
  }

  /**
   * Vertical alignment of this cell.
   *
   * This value is a member of the :ref:`MsoVerticalAnchor` enumeration or |None|. A value of
   * |None| indicates the cell has no explicitly applied vertical anchor setting and its
   * effective value is inherited from its style-hierarchy ancestors.
   *
   * Assigning |None| to this property causes any explicitly applied vertical anchor setting to
   * be cleared and inheritance of its effective value to be restored.
   */
  get verticalAnchor(): MSO_VERTICAL_ANCHOR | null {
    return this._tc.anchor;
  }

  set verticalAnchor(msoAnchorIdx: MSO_VERTICAL_ANCHOR | null) {
    this._tc.anchor = msoAnchorIdx;
  }

  /** Raise ValueError if `margin_value` is not a positive integer value or |None|. */
  static _validateMarginValue(marginValue: number | null): void {
    if ((!(typeof marginValue === "number") && marginValue !== null)) {
      const tmpl = "margin value must be integer or None, got '%s'";
      throw new TypeError(pyPercent(tmpl, [marginValue]));
    }
  }
}

/** Table column */
export class _Column extends Subshape {
  declare _parent: any;
  declare _gridCol: any;

  constructor(gridCol: CT_TableCol, parent: _ColumnCollection) {
    super(parent);
    this._parent = parent;
    this._gridCol = gridCol;
  }

  /** Width of column in EMU. */
  get width(): number {
    return this._gridCol.w;
  }

  set width(width: number) {
    this._gridCol.w = width;
    this._parent.notifyWidthChanged();
  }
}

/** Table row */
export class _Row extends Subshape {
  declare _parent: any;
  declare _tr: any;

  constructor(tr: CT_TableRow, parent: _RowCollection) {
    super(parent);
    this._parent = parent;
    this._tr = tr;
  }

  /**
   * Read-only reference to collection of cells in row.
   *
   * An individual cell is referenced using list notation, e.g. `cell = row.cells[0]`.
   */
  get cells() {
    return new _CellCollection(this._tr, this);
  }

  /** Height of row in EMU. */
  get height(): number {
    return this._tr.h;
  }

  set height(height: number) {
    this._tr.h = height;
    this._parent.notifyHeightChanged();
  }
}

/** Horizontal sequence of row cells */
export class _CellCollection extends Subshape {
  declare _parent: any;
  declare _tr: any;

  constructor(tr: CT_TableRow, parent: _Row) {
    super(parent);
    this._parent = parent;
    this._tr = tr;
  }

  /** Provides indexed access, (e.g. 'cells[0]'). */
  at(idx: number): _Cell {
    if (idx < 0 || idx >= this._tr.tcLst.length) {
      const msg = `cell index [${idx}] out of range`;
      throw new IndexError(msg);
    }
    return new _Cell(this._tr.tcLst[idx], this);
  }

  [PY_GETITEM](idx: number): _Cell {
    return this.at(idx);
  }

  /** Provides iterability. */
  *[Symbol.iterator](): Generator<_Cell> {
    for (const tc of this._tr.tcLst) yield new _Cell(tc, this);
  }

  /** Supports len() function (e.g. 'len(cells) == 1'). */
  get length(): number {
    return this._tr.tcLst.length;
  }
}

/** Sequence of table columns. */
export class _ColumnCollection extends Subshape {
  declare _parent: any;
  declare _tbl: any;

  constructor(tbl: CT_Table, parent: Table) {
    super(parent);
    this._parent = parent;
    this._tbl = tbl;
  }

  /** Provides indexed access, (e.g. 'columns[0]'). */
  at(idx: number): _Column {
    if (idx < 0 || idx >= this._tbl.tblGrid.gridColLst.length) {
      const msg = `column index [${idx}] out of range`;
      throw new IndexError(msg);
    }
    return new _Column(this._tbl.tblGrid.gridColLst[idx], this);
  }

  [PY_GETITEM](idx: number): _Column {
    return this.at(idx);
  }

  /** Iteration over the columns (python iterates via `__getitem__`). */
  *[Symbol.iterator](): Generator<_Column> {
    for (let i = 0; i < this.length; i++) yield this.at(i);
  }

  /** Supports len() function (e.g. 'len(columns) == 1'). */
  get length(): number {
    return this._tbl.tblGrid.gridColLst.length;
  }

  /** Called by a column when its width changes. Pass along to parent. */
  notifyWidthChanged() {
    this._parent.notifyWidthChanged();
  }
}

/** Sequence of table rows */
export class _RowCollection extends Subshape {
  declare _parent: any;
  declare _tbl: any;

  constructor(tbl: CT_Table, parent: Table) {
    super(parent);
    this._parent = parent;
    this._tbl = tbl;
  }

  /** Provides indexed access, (e.g. 'rows[0]'). */
  at(idx: number): _Row {
    if (idx < 0 || idx >= this.length) {
      const msg = `row index [${idx}] out of range`;
      throw new IndexError(msg);
    }
    return new _Row(this._tbl.trLst[idx], this);
  }

  [PY_GETITEM](idx: number): _Row {
    return this.at(idx);
  }

  /** Iteration over the rows (python iterates via `__getitem__`). */
  *[Symbol.iterator](): Generator<_Row> {
    for (let i = 0; i < this.length; i++) yield this.at(i);
  }

  /** Supports len() function (e.g. 'len(rows) == 1'). */
  get length(): number {
    return this._tbl.trLst.length;
  }

  /** Called by a row when its height changes. Pass along to parent. */
  notifyHeightChanged() {
    this._parent.notifyHeightChanged();
  }
}
