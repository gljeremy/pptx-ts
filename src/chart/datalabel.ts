// Data label-related objects. (`pptx.chart.datalabel`)
import { lazy, pyBool } from "../util.js";
import { Font, TextFrame } from "../text/text.js";

/**
 * Provides access to properties of data labels for a plot or a series.
 *
 * This is not a collection and does not provide access to individual data
 * labels. Access to individual labels is via the |Point| object. The
 * properties this object provides control formatting of *all* the data
 * labels in its scope.
 */
export class DataLabels {
  declare _element: any;

  constructor(dLbls: any) {
    this._element = dLbls;
  }

  /**
   * The |Font| object that provides access to the text properties for
   * these data labels, such as bold, italic, etc.
   */
  get font() {
    return lazy(this, "font", () => {
      const defRPr = this._element.defRPr;
      const font = new Font(defRPr);
      return font;
    });
  }

  /**
   * Read/write string specifying the format for the numbers on this set
   * of data labels. Returns 'General' if no number format has been set.
   * Note that this format string has no effect on rendered data labels
   * when :meth:`number_format_is_linked` is |True|. Assigning a format
   * string to this property automatically sets
   * :meth:`number_format_is_linked` to |False|.
   */
  get numberFormat() {
    const numFmt = this._element.numFmt;
    if (numFmt === null) {
      return "General";
    }
    return numFmt.formatCode;
  }

  set numberFormat(value: any) {
    this._element.getOrAddNumFmt().formatCode = value;
    this.numberFormatIsLinked = false;
  }

  /**
   * Read/write boolean specifying whether number formatting should be
   * taken from the source spreadsheet rather than the value of
   * :meth:`number_format`.
   */
  get numberFormatIsLinked() {
    const numFmt = this._element.numFmt;
    if (numFmt === null) {
      return true;
    }
    const souceLinked = numFmt.sourceLinked;
    if (souceLinked === null) {
      return true;
    }
    return numFmt.sourceLinked;
  }

  set numberFormatIsLinked(value: any) {
    const numFmt = this._element.getOrAddNumFmt();
    numFmt.sourceLinked = value;
  }

  /**
   * Read/write :ref:`XlDataLabelPosition` enumeration value specifying
   * the position of the data labels with respect to their data point, or
   * |None| if no position is specified. Assigning |None| causes
   * PowerPoint to choose the default position, which varies by chart
   * type.
   */
  get position() {
    const dLblPos = this._element.dLblPos;
    if (dLblPos === null) {
      return null;
    }
    return dLblPos.val;
  }

  set position(value: any) {
    if (value === null) {
      this._element._removeDLblPos();
      return;
    }
    this._element.getOrAddDLblPos().val = value;
  }

  /** Read/write. True when name of category should appear in label. */
  get showCategoryName() {
    return this._element.getOrAddShowCatName().val;
  }

  set showCategoryName(value: any) {
    this._element.getOrAddShowCatName().val = pyBool(value);
  }

  /** Read/write. True when data label displays legend-color swatch. */
  get showLegendKey() {
    return this._element.getOrAddShowLegendKey().val;
  }

  set showLegendKey(value: any) {
    this._element.getOrAddShowLegendKey().val = pyBool(value);
  }

  /**
   * Read/write. True when data label displays percentage.
   *
   * This option is not operative on all chart types. Percentage appears
   * on polar charts such as pie and donut.
   */
  get showPercentage() {
    return this._element.getOrAddShowPercent().val;
  }

  set showPercentage(value: any) {
    this._element.getOrAddShowPercent().val = pyBool(value);
  }

  /** Read/write. True when data label displays series name. */
  get showSeriesName() {
    return this._element.getOrAddShowSerName().val;
  }

  set showSeriesName(value: any) {
    this._element.getOrAddShowSerName().val = pyBool(value);
  }

  /** Read/write. True when label displays numeric value of datapoint. */
  get showValue() {
    return this._element.getOrAddShowVal().val;
  }

  set showValue(value: any) {
    this._element.getOrAddShowVal().val = pyBool(value);
  }
}

/** The data label associated with an individual data point. */
export class DataLabel {
  declare _ser: any;
  declare _element: any;
  declare _idx: any;

  constructor(ser: any, idx: any) {
    this._ser = ser;
    this._element = ser;
    this._idx = idx;
  }

  /**
   * The |Font| object providing text formatting for this data label.
   *
   * This font object is used to customize the appearance of automatically
   * inserted text, such as the data point value. The font applies to the
   * entire data label. More granular control of the appearance of custom
   * data label text is controlled by a font object on runs in the text
   * frame.
   */
  get font() {
    return lazy(this, "font", () => {
      const txPr = this._getOrAddTxPr();
      const textFrame = new TextFrame(txPr, this as any);
      const paragraph = textFrame.paragraphs[0];
      return paragraph.font;
    });
  }

  /**
   * Return |True| if this data label has a text frame (implying it has
   * custom data label text), and |False| otherwise. Assigning |True|
   * causes a text frame to be added if not already present. Assigning
   * |False| causes any existing text frame to be removed along with any
   * text contained in the text frame.
   */
  get hasTextFrame() {
    const dLbl = this._dLbl;
    if (dLbl === null) {
      return false;
    }
    if (dLbl.xpath("c:tx/c:rich").length) {
      return true;
    }
    return false;
  }

  set hasTextFrame(value: any) {
    if (pyBool(value) === true) {
      this._getOrAddTxRich();
    } else {
      this._removeTxRich();
    }
  }

  /**
   * Read/write :ref:`XlDataLabelPosition` member specifying the position
   * of this data label with respect to its data point, or |None| if no
   * position is specified. Assigning |None| causes PowerPoint to choose
   * the default position, which varies by chart type.
   */
  get position() {
    const dLbl = this._dLbl;
    if (dLbl === null) {
      return null;
    }
    const dLblPos = dLbl.dLblPos;
    if (dLblPos === null) {
      return null;
    }
    return dLblPos.val;
  }

  set position(value: any) {
    let dLbl;
    if (value === null) {
      dLbl = this._dLbl;
      if (dLbl === null) {
        return;
      }
      dLbl._removeDLblPos();
      return;
    }
    dLbl = this._getOrAddDLbl();
    dLbl.getOrAddDLblPos().val = value;
  }

  /**
   * |TextFrame| instance for this data label, containing the text of the
   * data label and providing access to its text formatting properties.
   */
  get textFrame() {
    const rich = this._getOrAddRich();
    return new TextFrame(rich, this as any);
  }

  /**
   * Return the |CT_DLbl| instance referring specifically to this
   * individual data label (having the same index value), or |None| if not
   * present.
   */
  get _dLbl() {
    return this._ser.getDLbl(this._idx);
  }

  /**
   * The ``CT_DLbl`` instance referring specifically to this individual
   * data label, newly created if not yet present in the XML.
   */
  _getOrAddDLbl() {
    return this._ser.getOrAddDLbl(this._idx);
  }

  /**
   * Return the `c:rich` element representing the text frame for this data
   * label, newly created with its ancestors if not present.
   */
  _getOrAddRich() {
    const dLbl = this._getOrAddDLbl();
    dLbl._removeSpPr();
    dLbl._removeTxPr();
    return dLbl.getOrAddRich();
  }

  /**
   * Return the `c:tx` element for this data label, with its `c:rich`
   * child and descendants, newly created if not yet present.
   */
  _getOrAddTxRich() {
    const dLbl = this._getOrAddDLbl();
    dLbl._removeSpPr();
    dLbl._removeTxPr();
    return dLbl.getOrAddTxRich();
  }

  /**
   * Return the `c:txPr` element for this data label.
   *
   * The `c:txPr` element and its parent `c:dLbl` element are created if
   * not yet present.
   */
  _getOrAddTxPr() {
    const dLbl = this._getOrAddDLbl();
    return dLbl.getOrAddTxPr();
  }

  /**
   * Remove any `c:tx/c:rich` child of the `c:dLbl` element for this data
   * label. Do nothing if that element is not present.
   */
  _removeTxRich() {
    const dLbl = this._dLbl;
    if (dLbl === null) {
      return;
    }
    dLbl.removeTxRich();
  }
}
