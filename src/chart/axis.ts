// Axis-related chart objects. (`pptx.chart.axis`)
import { ValueError } from "../errors.js";
import { lazy, pyBool, pyItem } from "../util.js";
import { ChartFormat } from "../dml/chtfmt.js";
import { XL_AXIS_CROSSES, XL_CATEGORY_TYPE, XL_TICK_LABEL_POSITION, XL_TICK_MARK } from "../enum/chart.js";
import { qn } from "../oxml/ns.js";
import { ST_Orientation } from "../oxml/simpletypes.js";
import { ElementProxy } from "../shared.js";
import { Font, TextFrame } from "../text/text.js";

/** Base class for chart axis objects. All axis objects share these properties. */
export class _BaseAxis {
  declare _element: any;
  declare _xAx: any;

  constructor(xAx: any) {
    this._element = xAx;
    this._xAx = xAx;
  }

  /**
   * An |AxisTitle| object providing access to title properties.
   *
   * Calling this property is destructive in the sense that it adds an
   * axis title element (`c:title`) to the axis XML if one is not already
   * present. Use :attr:`has_title` to test for presence of axis title
   * non-destructively.
   */
  get axisTitle() {
    return new AxisTitle(this._element.getOrAddTitle());
  }

  /**
   * The |ChartFormat| object providing access to the shape formatting
   * properties of this axis, such as its line color and fill.
   */
  get format() {
    return lazy(this, "format", () => {
      return new ChartFormat(this._element);
    });
  }

  /**
   * Read/write boolean value specifying whether this axis has gridlines
   * at its major tick mark locations. Assigning |True| to this property
   * causes major gridlines to be displayed. Assigning |False| causes them
   * to be removed.
   */
  get hasMajorGridlines() {
    if (this._element.majorGridlines === null) {
      return false;
    }
    return true;
  }

  set hasMajorGridlines(value: any) {
    if (pyBool(value) === true) {
      this._element.getOrAddMajorGridlines();
    } else {
      this._element._removeMajorGridlines();
    }
  }

  /**
   * Read/write boolean value specifying whether this axis has gridlines
   * at its minor tick mark locations. Assigning |True| to this property
   * causes minor gridlines to be displayed. Assigning |False| causes them
   * to be removed.
   */
  get hasMinorGridlines() {
    if (this._element.minorGridlines === null) {
      return false;
    }
    return true;
  }

  set hasMinorGridlines(value: any) {
    if (pyBool(value) === true) {
      this._element.getOrAddMinorGridlines();
    } else {
      this._element._removeMinorGridlines();
    }
  }

  /**
   * Read/write boolean specifying whether this axis has a title.
   *
   * |True| if this axis has a title, |False| otherwise. Assigning |True|
   * causes an axis title to be added if not already present. Assigning
   * |False| causes any existing title to be deleted.
   */
  get hasTitle() {
    if (this._element.title === null) {
      return false;
    }
    return true;
  }

  set hasTitle(value: any) {
    if (pyBool(value) === true) {
      this._element.getOrAddTitle();
    } else {
      this._element._removeTitle();
    }
  }

  /**
   * The |MajorGridlines| object representing the major gridlines for
   * this axis.
   */
  get majorGridlines() {
    return lazy(this, "majorGridlines", () => {
      return new MajorGridlines(this._element);
    });
  }

  /**
   * Read/write :ref:`XlTickMark` value specifying the type of major tick
   * mark to display on this axis.
   */
  get majorTickMark() {
    const majorTickMark = this._element.majorTickMark;
    if (majorTickMark === null) {
      return XL_TICK_MARK.CROSS;
    }
    return majorTickMark.val;
  }

  set majorTickMark(value: any) {
    this._element._removeMajorTickMark();
    if (value === XL_TICK_MARK.CROSS) {
      return;
    }
    this._element._addMajorTickMark({ val: value });
  }

  /**
   * Read/write float value specifying the upper limit of the value range
   * for this axis, the number at the top or right of the vertical or
   * horizontal value scale, respectively. The value |None| indicates the
   * upper limit should be determined automatically based on the range of
   * data point values associated with the axis.
   */
  get maximumScale() {
    return this._element.scaling.maximum;
  }

  set maximumScale(value: any) {
    const scaling = this._element.scaling;
    scaling.maximum = value;
  }

  /**
   * Read/write float value specifying lower limit of value range, the
   * number at the bottom or left of the value scale. |None| if no minimum
   * scale has been set. The value |None| indicates the lower limit should
   * be determined automatically based on the range of data point values
   * associated with the axis.
   */
  get minimumScale() {
    return this._element.scaling.minimum;
  }

  set minimumScale(value: any) {
    const scaling = this._element.scaling;
    scaling.minimum = value;
  }

  /**
   * Read/write :ref:`XlTickMark` value specifying the type of minor tick
   * mark for this axis.
   */
  get minorTickMark() {
    const minorTickMark = this._element.minorTickMark;
    if (minorTickMark === null) {
      return XL_TICK_MARK.CROSS;
    }
    return minorTickMark.val;
  }

  set minorTickMark(value: any) {
    this._element._removeMinorTickMark();
    if (value === XL_TICK_MARK.CROSS) {
      return;
    }
    this._element._addMinorTickMark({ val: value });
  }

  /**
   * Read/write bool value specifying whether to reverse plotting order for axis.
   *
   * For a category axis, this reverses the order in which the categories are
   * displayed. This may be desired, for example, on a (horizontal) bar-chart where
   * by default the first category appears at the bottom. Since we read from
   * top-to-bottom, many viewers may find it most natural for the first category to
   * appear on top.
   *
   * For a value axis, it reverses the direction of increasing value from
   * bottom-to-top to top-to-bottom.
   */
  get reverseOrder() {
    return this._element.orientation === ST_Orientation.MAX_MIN;
  }

  set reverseOrder(value: any) {
    this._element.orientation = (pyBool(value) === true ? ST_Orientation.MAX_MIN : ST_Orientation.MIN_MAX);
  }

  /**
   * The |TickLabels| instance providing access to axis tick label
   * formatting properties. Tick labels are the numbers appearing on
   * a value axis or the category names appearing on a category axis.
   */
  get tickLabels() {
    return lazy(this, "tickLabels", () => {
      return new TickLabels(this._element);
    });
  }

  /**
   * Read/write :ref:`XlTickLabelPosition` value specifying where the tick
   * labels for this axis should appear.
   */
  get tickLabelPosition() {
    const tickLblPos = this._element.tickLblPos;
    if (tickLblPos === null) {
      return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS;
    }
    if (tickLblPos.val === null) {
      return XL_TICK_LABEL_POSITION.NEXT_TO_AXIS;
    }
    return tickLblPos.val;
  }

  set tickLabelPosition(value: any) {
    const tickLblPos = this._element.getOrAddTickLblPos();
    tickLblPos.val = value;
  }

  /** Read/write. |True| if axis is visible, |False| otherwise. */
  get visible() {
    const delete_ = this._element.delete_;
    if (delete_ === null) {
      return false;
    }
    return (delete_.val ? false : true);
  }

  set visible(value: any) {
    if (![true, false].includes(value)) {
      throw new ValueError(`assigned value must be True or False, got: ${value}`);
    }
    const delete_ = this._element.getOrAddDelete_();
    delete_.val = !value;
  }
}

/** Provides properties for manipulating axis title. */
export class AxisTitle extends ElementProxy {
  declare _title: any;

  constructor(title: any) {
    super(title);
    this._title = title;
  }

  /**
   * |ChartFormat| object providing access to shape formatting.
   *
   * Return the |ChartFormat| object providing shape formatting properties
   * for this axis title, such as its line color and fill.
   */
  get format() {
    return lazy(this, "format", () => {
      return new ChartFormat(this._element);
    });
  }

  /**
   * Read/write Boolean specifying presence of a text frame.
   *
   * Return |True| if this axis title has a text frame, and |False|
   * otherwise. Assigning |True| causes a text frame to be added if not
   * already present. Assigning |False| causes any existing text frame to
   * be removed along with any text contained in the text frame.
   */
  get hasTextFrame() {
    if (this._title.txRich === null) {
      return false;
    }
    return true;
  }

  set hasTextFrame(value: any) {
    if (pyBool(value) === true) {
      this._title.getOrAddTxRich();
    } else {
      this._title._removeTx();
    }
  }

  /**
   * |TextFrame| instance for this axis title.
   *
   * Return a |TextFrame| instance allowing read/write access to the text
   * of this axis title and its text formatting properties. Accessing this
   * property is destructive as it adds a new text frame if not already
   * present.
   */
  get textFrame() {
    const rich = this._title.getOrAddTxRich();
    return new TextFrame(rich, this as any);
  }
}

/** A category axis of a chart. */
export class CategoryAxis extends _BaseAxis {

  /**
   * A member of :ref:`XlCategoryType` specifying the scale type of this
   * axis. Unconditionally ``CATEGORY_SCALE`` for a |CategoryAxis| object.
   */
  get categoryType() {
    return XL_CATEGORY_TYPE.CATEGORY_SCALE;
  }
}

/**
 * A category axis with dates as its category labels.
 *
 * This axis-type has some special display behaviors such as making length of equal
 * periods equal and normalizing month start dates despite unequal month lengths.
 */
export class DateAxis extends _BaseAxis {

  /**
   * A member of :ref:`XlCategoryType` specifying the scale type of this
   * axis. Unconditionally ``TIME_SCALE`` for a |DateAxis| object.
   */
  get categoryType() {
    return XL_CATEGORY_TYPE.TIME_SCALE;
  }
}

/** Provides access to the properties of the major gridlines appearing on an axis. */
export class MajorGridlines extends ElementProxy {
  declare _xAx: any;

  constructor(xAx: any) {
    super(xAx);
    this._xAx = xAx;
  }

  /**
   * The |ChartFormat| object providing access to the shape formatting
   * properties of this data point, such as line and fill.
   */
  get format() {
    return lazy(this, "format", () => {
      const majorGridlines = this._xAx.getOrAddMajorGridlines();
      return new ChartFormat(majorGridlines);
    });
  }
}

/** A service class providing access to formatting of axis tick mark labels. */
export class TickLabels {
  declare _element: any;

  constructor(xAxElm: any) {
    this._element = xAxElm;
  }

  /**
   * The |Font| object that provides access to the text properties for
   * these tick labels, such as bold, italic, etc.
   */
  get font() {
    return lazy(this, "font", () => {
      const defRPr = this._element.defRPr;
      const font = new Font(defRPr);
      return font;
    });
  }

  /**
   * Read/write string (e.g. "$#,##0.00") specifying the format for the
   * numbers on this axis. The syntax for these strings is the same as it
   * appears in the PowerPoint or Excel UI. Returns 'General' if no number
   * format has been set. Note that this format string has no effect on
   * rendered tick labels when :meth:`number_format_is_linked` is |True|.
   * Assigning a format string to this property automatically sets
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
    const numFmt = this._element.getOrAddNumFmt();
    numFmt.formatCode = value;
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
      return false;
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
   * Read/write int value in range 0-1000 specifying the spacing between
   * the tick mark labels and the axis as a percentange of the default
   * value. 100 if no label offset setting is present.
   */
  get offset() {
    const lblOffset = this._element.lblOffset;
    if (lblOffset === null) {
      return 100;
    }
    return lblOffset.val;
  }

  set offset(value: any) {
    if (this._element.tag !== qn("c:catAx")) {
      throw new ValueError("only a category axis has an offset");
    }
    this._element._removeLblOffset();
    if (value === 100) {
      return;
    }
    const lblOffset = this._element._addLblOffset();
    lblOffset.val = value;
  }
}

/**
 * An axis having continuous (as opposed to discrete) values.
 *
 * The vertical axis is generally a value axis, however both axes of an XY-type chart
 * are value axes.
 */
export class ValueAxis extends _BaseAxis {

  /**
   * Member of :ref:`XlAxisCrosses` enumeration specifying the point on
   * this axis where the other axis crosses, such as auto/zero, minimum,
   * or maximum. Returns `XL_AXIS_CROSSES.CUSTOM` when a specific numeric
   * crossing point (e.g. 1.5) is defined.
   */
  get crosses() {
    const crosses = this._crossXAx.crosses;
    if (crosses === null) {
      return XL_AXIS_CROSSES.CUSTOM;
    }
    return crosses.val;
  }

  set crosses(value: any) {
    const crossXAx = this._crossXAx;
    if (value === XL_AXIS_CROSSES.CUSTOM) {
      if (crossXAx.crossesAt !== null) {
        return;
      }
    }
    crossXAx._removeCrosses();
    crossXAx._removeCrossesAt();
    if (value === XL_AXIS_CROSSES.CUSTOM) {
      crossXAx._addCrossesAt({ val: 0.0 });
    } else {
      crossXAx._addCrosses({ val: value });
    }
  }

  /**
   * Numeric value on this axis at which the perpendicular axis crosses.
   * Returns |None| if no crossing value is set.
   */
  get crossesAt() {
    const crossesAt = this._crossXAx.crossesAt;
    if (crossesAt === null) {
      return null;
    }
    return crossesAt.val;
  }

  set crossesAt(value: any) {
    const crossXAx = this._crossXAx;
    crossXAx._removeCrosses();
    crossXAx._removeCrossesAt();
    if (value === null) {
      return;
    }
    crossXAx._addCrossesAt({ val: value });
  }

  /**
   * The float number of units between major tick marks on this value
   * axis. |None| corresponds to the 'Auto' setting in the UI, and
   * specifies the value should be calculated by PowerPoint based on the
   * underlying chart data.
   */
  get majorUnit() {
    const majorUnit = this._element.majorUnit;
    if (majorUnit === null) {
      return null;
    }
    return majorUnit.val;
  }

  set majorUnit(value: any) {
    this._element._removeMajorUnit();
    if (value === null) {
      return;
    }
    this._element._addMajorUnit({ val: value });
  }

  /**
   * The float number of units between minor tick marks on this value
   * axis. |None| corresponds to the 'Auto' setting in the UI, and
   * specifies the value should be calculated by PowerPoint based on the
   * underlying chart data.
   */
  get minorUnit() {
    const minorUnit = this._element.minorUnit;
    if (minorUnit === null) {
      return null;
    }
    return minorUnit.val;
  }

  set minorUnit(value: any) {
    this._element._removeMinorUnit();
    if (value === null) {
      return;
    }
    this._element._addMinorUnit({ val: value });
  }

  /**
   * The axis element in the same group (primary/secondary) that crosses
   * this axis.
   */
  get _crossXAx() {
    const crossAxIdVal = this._element.crossAx.val;
    const expr = `(../c:catAx | ../c:valAx | ../c:dateAx)/c:axId[@val="${crossAxIdVal}"]`;
    const crossAxId: any = pyItem(this._element.xpath(expr), 0);
    return crossAxId.getparent();
  }
}
