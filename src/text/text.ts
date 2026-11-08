// Text-related objects such as TextFrame and Paragraph. (`pptx.text.text`)
import { ValueError } from "../errors.js";
import { lazy } from "../util.js";
import { FillFormat } from "../dml/fill.js";
import { MSO_FILL } from "../enum/dml.js";
import { MSO_LANGUAGE_ID } from "../enum/lang.js";
import { MSO_AUTO_SIZE, MSO_UNDERLINE, MSO_VERTICAL_ANCHOR } from "../enum/text.js";
import { RELATIONSHIP_TYPE as RT } from "../opc/constants.js";
import { ST_TextWrappingType } from "../oxml/simpletypes.js";
import { Subshape } from "../shapes/index.js";
import { FontFiles } from "./fonts.js";
import { TextFitter } from "./layout.js";
import { Centipoints, Emu, Length, Pt } from "../util.js";
import type { ColorFormat } from "../dml/color.js";
import type { MSO_TEXT_UNDERLINE_TYPE, PP_PARAGRAPH_ALIGNMENT } from "../enum/text.js";
import type { CT_Hyperlink } from "../oxml/action.js";
import type { CT_RegularTextRun, CT_TextBody, CT_TextCharacterProperties, CT_TextParagraph, CT_TextParagraphProperties } from "../oxml/text.js";
import type { ProvidesExtents, ProvidesPart } from "../types.js";

/**
 * The part of a shape that contains its text.
 *
 * Not all shapes have a text frame. Corresponds to the `p:txBody` element that can
 * appear as a child element of `p:sp`. Not intended to be constructed directly.
 */
export class TextFrame extends Subshape {
  declare _element: any;
  declare _txBody: any;
  declare _parent: any;

  constructor(txBody: CT_TextBody, parent: ProvidesPart) {
    super(parent);
    this._element = txBody;
    this._txBody = txBody;
    this._parent = parent;
  }

  /**
   * Return new |_Paragraph| instance appended to the sequence of
   * paragraphs contained in this text frame.
   */
  addParagraph() {
    const p = this._txBody.addP();
    return new _Paragraph(p, this);
  }

  /**
   * Resizing strategy used to fit text within this shape.
   *
   * Determins the type of automatic resizing used to fit the text of this shape within its
   * bounding box when the text would otherwise extend beyond the shape boundaries. May be
   * |None|, `MSO_AUTO_SIZE.NONE`, `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`, or
   * `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`.
   */
  get autoSize(): MSO_AUTO_SIZE | null {
    return this._bodyPr.autofit;
  }

  set autoSize(value: MSO_AUTO_SIZE | null) {
    this._bodyPr.autofit = value;
  }

  /** Remove all paragraphs except one empty one. */
  clear() {
    let p;
    for (const p of this._txBody.pLst.slice(1)) {
      this._txBody.remove(p);
    }
    p = this.paragraphs[0];
    p.clear();
  }

  /**
   * Fit text-frame text entirely within bounds of its shape.
   *
   * Make the text in this text frame fit entirely within the bounds of its shape by setting
   * word wrap on and applying the "best-fit" font size to all the text it contains.
   *
   * :attr:`TextFrame.auto_size` is set to :attr:`MSO_AUTO_SIZE.NONE`. The font size will not
   * be set larger than `max_size` points. If the path to a matching TrueType font is provided
   * as `font_file`, that font file will be used for the font metrics. If `font_file` is |None|,
   * best efforts are made to locate a font file with matchhing `font_family`, `bold`, and
   * `italic` installed on the current system (usually succeeds if the font is installed).
   */
  fitText(fontFamily: string = "Calibri", maxSize: number = 18, bold: boolean = false, italic: boolean = false, fontFile: string | null = null) {
    if (this.text === "") {
      return;
    }
    const fontSize = this._bestFitFontSize(fontFamily, maxSize, bold, italic, fontFile);
    this._applyFit(fontFamily, fontSize, bold, italic);
  }

  /**
   * |Length| value representing the inset of text from the bottom text frame border.
   *
   * :meth:`pptx.util.Inches` provides a convenient way of setting the value, e.g.
   * `text_frame.margin_bottom = Inches(0.05)`.
   */
  get marginBottom(): number {
    return this._bodyPr.bIns;
  }

  set marginBottom(emu: number) {
    this._bodyPr.bIns = emu;
  }

  /** Inset of text from left text frame border as |Length| value. */
  get marginLeft(): number {
    return this._bodyPr.lIns;
  }

  set marginLeft(emu: number) {
    this._bodyPr.lIns = emu;
  }

  /** Inset of text from right text frame border as |Length| value. */
  get marginRight(): number {
    return this._bodyPr.rIns;
  }

  set marginRight(emu: number) {
    this._bodyPr.rIns = emu;
  }

  /** Inset of text from top text frame border as |Length| value. */
  get marginTop(): number {
    return this._bodyPr.tIns;
  }

  set marginTop(emu: number) {
    this._bodyPr.tIns = emu;
  }

  /**
   * Sequence of paragraphs in this text frame.
   *
   * A text frame always contains at least one paragraph.
   */
  get paragraphs(): _Paragraph[] {
    return [...[...this._txBody.pLst].map((p) => new _Paragraph(p, this))];
  }

  /**
   * All text in this text-frame as a single string.
   *
   * Read/write. The return value contains all text in this text-frame. A line-feed character
   * (`"\n"`) separates the text for each paragraph. A vertical-tab character (`"\v"`) appears
   * for each line break (aka. soft carriage-return) encountered.
   *
   * The vertical-tab character is how PowerPoint represents a soft carriage return in clipboard
   * text, which is why that encoding was chosen.
   *
   * Assignment replaces all text in the text frame. A new paragraph is added for each line-feed
   * character (`"\n"`) encountered. A line-break (soft carriage-return) is inserted for each
   * vertical-tab character (`"\v"`) encountered.
   *
   * Any control character other than newline, tab, or vertical-tab are escaped as plain-text
   * like "_x001B_" (for ESC (ASCII 32) in this example).
   */
  get text(): string {
    return [...this.paragraphs].map((paragraph) => paragraph.text).join("\n");
  }

  set text(text: string) {
    let pText;
    const txBody = this._txBody;
    txBody.clearContent();
    for (const pText of text.split("\n")) {
      const p = txBody.addP();
      p.appendText(pText);
    }
  }

  /**
   * Represents the vertical alignment of text in this text frame.
   *
   * |None| indicates the effective value should be inherited from this object's style hierarchy.
   */
  get verticalAnchor(): MSO_VERTICAL_ANCHOR | null {
    return this._txBody.bodyPr.anchor;
  }

  set verticalAnchor(value: MSO_VERTICAL_ANCHOR | null) {
    const bodyPr = this._txBody.bodyPr;
    bodyPr.anchor = value;
  }

  /**
   * `True` when lines of text in this shape are wrapped to fit within the shape's width.
   *
   * Read-write. Valid values are True, False, or None. True and False turn word wrap on and
   * off, respectively. Assigning None to word wrap causes any word wrap setting to be removed
   * from the text frame, causing it to inherit this setting from its style hierarchy.
   */
  get wordWrap(): boolean | null {
    const wrap = this._txBody.bodyPr.wrap;
    return wrap === ST_TextWrappingType.SQUARE ? true : wrap === ST_TextWrappingType.NONE ? false : null;
  }

  set wordWrap(value: boolean | null) {
    if (value !== true && value !== false && value !== null) {
      throw new ValueError(`assigned value must be True, False, or None, got ${value}`);
    }
    this._txBody.bodyPr.wrap = value === true ? ST_TextWrappingType.SQUARE : value === false ? ST_TextWrappingType.NONE : null;
  }

  /**
   * Arrange text in this text frame to fit inside its extents.
   *
   * This is accomplished by setting auto size off, wrap on, and setting the font of
   * all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
   */
  _applyFit(fontFamily: string, fontSize: number, isBold: boolean, isItalic: boolean) {
    this.autoSize = MSO_AUTO_SIZE.NONE;
    this.wordWrap = true;
    this._setFont(fontFamily, fontSize, isBold, isItalic);
  }

  /**
   * Return font-size in points that best fits text in this text-frame.
   *
   * The best-fit font size is the largest integer point size not greater than `max_size` that
   * allows all the text in this text frame to fit inside its extents when rendered using the
   * font described by `family`, `bold`, and `italic`. If `font_file` is specified, it is used
   * to calculate the fit, whether or not it matches `family`, `bold`, and `italic`.
   */
  _bestFitFontSize(family: string, maxSize: number, bold: boolean, italic: boolean, fontFile: string | null): number {
    if (fontFile === null) {
      fontFile = FontFiles.find(family, bold, italic);
    }
    return TextFitter.bestFitFontSize(this.text, this._extents, maxSize, fontFile);
  }

  get _bodyPr() {
    return this._txBody.bodyPr;
  }

  /**
   * (cx, cy) 2-tuple representing the effective rendering area of this text-frame.
   *
   * Margins are taken into account.
   */
  get _extents(): [number, number] {
    const parent = this._parent;
    return [Emu(parent.width - this.marginLeft - this.marginRight), Emu(parent.height - this.marginTop - this.marginBottom)];
  }

  /** Set the font properties of all the text in this text frame. */
  _setFont(family: string, size: number, bold: boolean, italic: boolean) {
    let rPr, p, elm;
    const iterRPrs = function* (this: any, txBody: CT_TextBody) {
      for (const p of txBody.pLst) {
        for (const elm of p.contentChildren) {
          yield elm.getOrAddRPr();
        }
        yield p.getOrAddEndParaRPr();
      }
    }.bind(this);
    const setRPrFont = (rPr: CT_TextCharacterProperties, name: string, size: number, bold: boolean, italic: boolean) => {
      const f = new Font(rPr);
      [f.name, f.size, f.bold, f.italic] = [family, Pt(size), bold, italic];
    };
    const txBody = this._element;
    for (const rPr of iterRPrs(txBody)) {
      setRPrFont(rPr, family, size, bold, italic);
    }
  }
}

/**
 * Character properties object, providing font size, font name, bold, italic, etc.
 *
 * Corresponds to `a:rPr` child element of a run. Also appears as `a:defRPr` and
 * `a:endParaRPr` in paragraph and `a:defRPr` in list style elements.
 */
export class Font {
  declare _element: any;
  declare _rPr: any;

  constructor(rPr: CT_TextCharacterProperties) {
    this._element = rPr;
    this._rPr = rPr;
  }

  /**
   * Get or set boolean bold value of |Font|, e.g. `paragraph.font.bold = True`.
   *
   * If set to |None|, the bold setting is cleared and is inherited from an enclosing shape's
   * setting, or a setting in a style or master. Returns None if no bold attribute is present,
   * meaning the effective bold value is inherited from a master or the theme.
   */
  get bold(): boolean | null {
    return this._rPr.b;
  }

  set bold(value: boolean | null) {
    this._rPr.b = value;
  }

  /** The |ColorFormat| instance that provides access to the color settings for this font. */
  get color(): ColorFormat {
    return lazy(this, "color", () => {
      if (this.fill.type !== MSO_FILL.SOLID) {
        this.fill.solid();
      }
      return this.fill.foreColor;
    });
  }

  /**
   * |FillFormat| instance for this font.
   *
   * Provides access to fill properties such as fill color.
   */
  get fill(): FillFormat {
    return lazy(this, "fill", () => {
      return FillFormat.fromFillParent(this._rPr);
    });
  }

  /**
   * Get or set boolean italic value of |Font| instance.
   *
   * Has the same behaviors as bold with respect to None values.
   */
  get italic(): boolean | null {
    return this._rPr.i;
  }

  set italic(value: boolean | null) {
    this._rPr.i = value;
  }

  /**
   * Get or set the language id of this |Font| instance.
   *
   * The language id is a member of the :ref:`MsoLanguageId` enumeration. Assigning |None|
   * removes any language setting, the same behavior as assigning `MSO_LANGUAGE_ID.NONE`.
   */
  get languageId(): MSO_LANGUAGE_ID | null {
    const lang = this._rPr.lang;
    if (lang === null) {
      return MSO_LANGUAGE_ID.NONE;
    }
    return this._rPr.lang;
  }

  set languageId(value: MSO_LANGUAGE_ID | null) {
    if (value === MSO_LANGUAGE_ID.NONE) {
      value = null;
    }
    this._rPr.lang = value;
  }

  /**
   * Get or set the typeface name for this |Font| instance.
   *
   * Causes the text it controls to appear in the named font, if a matching font is found.
   * Returns |None| if the typeface is currently inherited from the theme. Setting it to |None|
   * removes any override of the theme typeface.
   */
  get name(): string | null {
    const latin = this._rPr.latin;
    if (latin === null) {
      return null;
    }
    return latin.typeface;
  }

  set name(value: string | null) {
    if (value === null) {
      this._rPr._removeLatin();
    } else {
      const latin = this._rPr.getOrAddLatin();
      latin.typeface = value;
    }
  }

  /**
   * Indicates the font height in English Metric Units (EMU).
   *
   * Read/write. |None| indicates the font size should be inherited from its style hierarchy,
   * such as a placeholder or document defaults (usually 18pt). |Length| is a subclass of |int|
   * having properties for convenient conversion into points or other length units. Likewise,
   * the :class:`pptx.util.Pt` class allows convenient specification of point values::
   *
   *     >>> font.size = Pt(24)
   *     >>> font.size
   *     304800
   *     >>> font.size.pt
   *     24.0
   */
  get size(): number | null {
    const sz = this._rPr.sz;
    if (sz === null) {
      return null;
    }
    return Centipoints(sz);
  }

  set size(emu: number | null) {
    if (emu === null) {
      this._rPr.sz = null;
    } else {
      const sz = new Length(emu).centipoints;
      this._rPr.sz = sz;
    }
  }

  /**
   * Indicaties the underline setting for this font.
   *
   * Value is |True|, |False|, |None|, or a member of the :ref:`MsoTextUnderlineType`
   * enumeration. |None| is the default and indicates the underline setting should be inherited
   * from the style hierarchy, such as from a placeholder. |True| indicates single underline.
   * |False| indicates no underline. Other settings such as double and wavy underlining are
   * indicated with members of the :ref:`MsoTextUnderlineType` enumeration.
   */
  get underline(): boolean | MSO_TEXT_UNDERLINE_TYPE | null {
    const u = this._rPr.u;
    if (u === MSO_UNDERLINE.NONE) {
      return false;
    }
    if (u === MSO_UNDERLINE.SINGLE_LINE) {
      return true;
    }
    return u;
  }

  set underline(value: boolean | MSO_TEXT_UNDERLINE_TYPE | null) {
    if (value === true) {
      value = MSO_UNDERLINE.SINGLE_LINE;
    } else if (value === false) {
      value = MSO_UNDERLINE.NONE;
    }
    this._element.u = value;
  }
}

/**
 * Text run hyperlink object.
 *
 * Corresponds to `a:hlinkClick` child element of the run's properties element (`a:rPr`).
 */
export class _Hyperlink extends Subshape {
  declare _rPr: any;

  constructor(rPr: CT_TextCharacterProperties, parent: ProvidesPart) {
    super(parent);
    this._rPr = rPr;
  }

  /**
   * The URL of the hyperlink.
   *
   * Read/write. URL can be on http, https, mailto, or file scheme; others may work.
   */
  get address(): string | null {
    if (this._hlinkClick === null) {
      return null;
    }
    return this.part.targetRef(this._hlinkClick.rId);
  }

  set address(url: string | null) {
    if (this._hlinkClick !== null) {
      this._removeHlinkClick();
    }
    if (url) {
      this._addHlinkClick(url);
    }
  }

  _addHlinkClick(url: string) {
    const rId = this.part.relateTo(url, RT.HYPERLINK, true);
    this._rPr.addHlinkClick(rId);
  }

  get _hlinkClick(): CT_Hyperlink | null {
    return this._rPr.hlinkClick;
  }

  _removeHlinkClick() {
    if (!(this._hlinkClick !== null)) throw new Error("AssertionError");
    this.part.dropRel(this._hlinkClick.rId);
    this._rPr._removeHlinkClick();
  }
}

/** Paragraph object. Not intended to be constructed directly. */
export class _Paragraph extends Subshape {
  declare _element: any;
  declare _p: any;

  constructor(p: CT_TextParagraph, parent: ProvidesPart) {
    super(parent);
    this._element = p;
    this._p = p;
  }

  /** Add line break at end of this paragraph. */
  addLineBreak() {
    this._p.addBr();
  }

  /** Return a new run appended to the runs in this paragraph. */
  addRun(): _Run {
    const r = this._p.addR();
    return new _Run(r, this);
  }

  /**
   * Horizontal alignment of this paragraph.
   *
   * The value |None| indicates the paragraph should 'inherit' its effective value from its
   * style hierarchy. Assigning |None| removes any explicit setting, causing its inherited
   * value to be used.
   */
  get alignment(): PP_PARAGRAPH_ALIGNMENT | null {
    return this._pPr.algn;
  }

  set alignment(value: PP_PARAGRAPH_ALIGNMENT | null) {
    this._pPr.algn = value;
  }

  /**
   * Remove all content from this paragraph.
   *
   * Paragraph properties are preserved. Content includes runs, line breaks, and fields.
   */
  clear() {
    let elm;
    for (const elm of this._element.contentChildren) {
      this._element.remove(elm);
    }
    return this;
  }

  /**
   * |Font| object containing default character properties for the runs in this paragraph.
   *
   * These character properties override default properties inherited from parent objects such
   * as the text frame the paragraph is contained in and they may be overridden by character
   * properties set at the run level.
   */
  get font(): Font {
    return new Font(this._defRPr);
  }

  /**
   * Indentation level of this paragraph.
   *
   * Read-write. Integer in range 0..8 inclusive. 0 represents a top-level paragraph and is the
   * default value. Indentation level is most commonly encountered in a bulleted list, as is
   * found on a word bullet slide.
   */
  get level(): number {
    return this._pPr.lvl;
  }

  set level(level: number) {
    this._pPr.lvl = level;
  }

  /**
   * The space between baselines in successive lines of this paragraph.
   *
   * A value of |None| indicates no explicit value is assigned and its effective value is
   * inherited from the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
   * indicates spacing is applied in multiples of line heights. A |Length| value such as
   * `Pt(12)` indicates spacing is a fixed height. The |Pt| value class is a convenient way to
   * apply line spacing in units of points.
   */
  get lineSpacing(): number | number | number | null {
    const pPr = this._p.pPr;
    if (pPr === null) {
      return null;
    }
    return pPr.lineSpacing;
  }

  set lineSpacing(value: number | number | number | null) {
    const pPr = this._p.getOrAddPPr();
    pPr.lineSpacing = value;
  }

  /** Sequence of runs in this paragraph. */
  get runs(): _Run[] {
    return [...[...this._element.rLst].map((r) => new _Run(r, this))];
  }

  /**
   * The spacing to appear between this paragraph and the subsequent paragraph.
   *
   * A value of |None| indicates no explicit value is assigned and its effective value is
   * inherited from the paragraph's style hierarchy. |Length| objects provide convenience
   * properties, such as `.pt` and `.inches`, that allow easy conversion to various length
   * units.
   */
  get spaceAfter(): number | null {
    const pPr = this._p.pPr;
    if (pPr === null) {
      return null;
    }
    return pPr.spaceAfter;
  }

  set spaceAfter(value: number | null) {
    const pPr = this._p.getOrAddPPr();
    pPr.spaceAfter = value;
  }

  /**
   * The spacing to appear between this paragraph and the prior paragraph.
   *
   * A value of |None| indicates no explicit value is assigned and its effective value is
   * inherited from the paragraph's style hierarchy. |Length| objects provide convenience
   * properties, such as `.pt` and `.cm`, that allow easy conversion to various length units.
   */
  get spaceBefore(): number | null {
    const pPr = this._p.pPr;
    if (pPr === null) {
      return null;
    }
    return pPr.spaceBefore;
  }

  set spaceBefore(value: number | null) {
    const pPr = this._p.getOrAddPPr();
    pPr.spaceBefore = value;
  }

  /**
   * Text of paragraph as a single string.
   *
   * Read/write. This value is formed by concatenating the text in each run and field making up
   * the paragraph, adding a vertical-tab character (`"\v"`) for each line-break element
   * (`<a:br>`, soft carriage-return) encountered.
   *
   * While the encoding of line-breaks as a vertical tab might be surprising at first, doing so
   * is consistent with PowerPoint's clipboard copy behavior and allows a line-break to be
   * distinguished from a paragraph boundary within the str return value.
   *
   * Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
   * (`"\v"`) in the assigned str is translated to a line-break, as is each line-feed
   * character (`"\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
   * If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
   * instead. Any other control characters in the assigned string are escaped as a hex
   * representation like "_x001B_" (for ESC (ASCII 27) in this example).
   */
  get text(): string {
    return [...this._element.contentChildren].map((elm) => elm.text).join("");
  }

  set text(text: string) {
    this.clear();
    this._element.appendText(text);
  }

  /**
   * The element that defines the default run properties for runs in this paragraph.
   *
   * Causes the element to be added if not present.
   */
  get _defRPr(): CT_TextCharacterProperties {
    return this._pPr.getOrAddDefRPr();
  }

  /**
   * Contains the properties for this paragraph.
   *
   * Causes the element to be added if not present.
   */
  get _pPr(): CT_TextParagraphProperties {
    return this._p.getOrAddPPr();
  }
}

/** Text run object. Corresponds to `a:r` child element in a paragraph. */
export class _Run extends Subshape {
  declare _r: any;

  constructor(r: CT_RegularTextRun, parent: ProvidesPart) {
    super(parent);
    this._r = r;
  }

  /**
   * |Font| instance containing run-level character properties for the text in this run.
   *
   * Character properties can be and perhaps most often are inherited from parent objects such
   * as the paragraph and slide layout the run is contained in. Only those specifically
   * overridden at the run level are contained in the font object.
   */
  get font() {
    const rPr = this._r.getOrAddRPr();
    return new Font(rPr);
  }

  /**
   * Proxy for any `a:hlinkClick` element under the run properties element.
   *
   * Created on demand, the hyperlink object is available whether an `a:hlinkClick` element is
   * present or not, and creates or deletes that element as appropriate in response to actions
   * on its methods and attributes.
   */
  get hyperlink(): _Hyperlink {
    return lazy(this, "hyperlink", () => {
      const rPr = this._r.getOrAddRPr();
      return new _Hyperlink(rPr, this);
    });
  }

  /**
   * Read/write. A unicode string containing the text in this run.
   *
   * Assignment replaces all text in the run. The assigned value can be a 7-bit ASCII
   * string, a UTF-8 encoded 8-bit string, or unicode. String values are converted to
   * unicode assuming UTF-8 encoding.
   *
   * Any other control characters in the assigned string other than tab or newline
   * are escaped as a hex representation. For example, ESC (ASCII 27) is escaped as
   * "_x001B_". Contrast the behavior of `TextFrame.text` and `_Paragraph.text` with
   * respect to line-feed and vertical-tab characters.
   */
  get text() {
    return this._r.text;
  }

  set text(text: string) {
    this._r.text = text;
  }
}
