// Custom element classes for text-related XML elements (`pptx.oxml.text`)
import { ValueError } from "../errors.js";
import { defineChildren } from "./xmlchemy.js";
import { MSO_LANGUAGE_ID } from "../enum/lang.js";
import { MSO_AUTO_SIZE, MSO_TEXT_UNDERLINE_TYPE, MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT } from "../enum/text.js";
import { InvalidXmlError } from "../exc.js";
import { parseXml } from "./parser.js";
import { CT_GradientFillProperties } from "./dml/fill.js";
import { nsdecls } from "./ns.js";
import { ST_Coordinate32, ST_TextFontScalePercentOrPercentString, ST_TextFontSize, ST_TextIndentLevelType, ST_TextSpacingPercentOrPercentString, ST_TextSpacingPoint, ST_TextTypeface, ST_TextWrappingType, XsdBoolean } from "./simpletypes.js";
import { BaseOxmlElement, Choice, OneAndOnlyOne, OneOrMore, OptionalAttribute, RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice } from "./xmlchemy.js";
import { Emu, isLengthValue } from "../util.js";
import type { CT_Hyperlink } from "./action.js";

/** `a:r` custom element class */
export class CT_RegularTextRun extends BaseOxmlElement {
  declare getOrAddRPr: (...args: any[]) => CT_TextCharacterProperties;
  declare rPr: CT_TextCharacterProperties | null;
  declare _addRPr: (attrs?: Record<string, any>) => CT_TextCharacterProperties;
  declare _insertRPr: (child: CT_TextCharacterProperties) => CT_TextCharacterProperties;
  declare _newRPr: () => CT_TextCharacterProperties;
  declare _removeRPr: () => void;
  declare t: BaseOxmlElement;
  static {
    defineChildren(this, {
      rPr: new ZeroOrOne("a:rPr", ["a:t"]),
      t: new OneAndOnlyOne("a:t"),
    });
  }

  /** All text of (required) `a:t` child. */
  get text(): string {
    const text = this.t.text;
    return (text || "");
  }

  set text(value: string) {
    this.t.text = CT_RegularTextRun._escapeCtrlChars(value);
  }

  /**
   * Return str after replacing each control character with a plain-text escape.
   *
   * For example, a BEL character (x07) would appear as "_x0007_". Horizontal-tab
   * (x09) and line-feed (x0A) are not escaped. All other characters in the range
   * x00-x1F are escaped.
   */
  static _escapeCtrlChars(s: string): string {
    return s.replace(/[\x00-\x08\x0B-\x1F]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
  }
}

/**
 * `p:txBody` custom element class.
 *
 * Also used for `c:txPr` in charts and perhaps other elements.
 */
export class CT_TextBody extends BaseOxmlElement {
  declare addP: (...args: any[]) => CT_TextParagraph;
  declare pLst: CT_TextParagraph[];
  declare bodyPr: CT_TextBodyProperties;
  declare _addP: (attrs?: Record<string, any>) => CT_TextParagraph;
  declare _insertP: (child: CT_TextParagraph) => CT_TextParagraph;
  declare _newP: () => CT_TextParagraph;
  static {
    defineChildren(this, {
      bodyPr: new OneAndOnlyOne("a:bodyPr"),
      p: new OneOrMore("a:p"),
    });
  }

  /**
   * Remove all `a:p` children, but leave any others.
   *
   * cf. lxml `_Element.clear()` method which removes all children.
   */
  clearContent() {
    for (const p of this.pLst) {
      this.remove(p);
    }
  }

  /**
   * `a:defRPr` element of required first `p` child, added with its ancestors if not present.
   *
   * Used when element is a ``c:txPr`` in a chart and the `p` element is used only to specify
   * formatting, not content.
   */
  get defRPr(): CT_TextCharacterProperties {
    const p = this.pLst[0];
    const pPr = p.getOrAddPPr();
    const defRPr = pPr.getOrAddDefRPr();
    return defRPr;
  }

  /** True if only a single empty `a:p` element is present. */
  get isEmpty(): boolean {
    const ps = this.pLst;
    if (ps.length > 1) {
      return false;
    }
    if (ps.length === 0) {
      throw new InvalidXmlError("p:txBody must have at least one a:p");
    }
    if (ps[0].text !== "") {
      return false;
    }
    return true;
  }

  /** Return a new `p:txBody` element tree. */
  static new() {
    const xml = this._txBodyTmpl();
    const txBody = parseXml(xml);
    return txBody;
  }

  /**
   * Return a new `a:txBody` element tree.
   *
   * Suitable for use in a table cell and possibly other situations.
   */
  static newATxBody(): CT_TextBody {
    const xml = this._aTxBodyTmpl();
    const txBody = parseXml(xml);
    return txBody;
  }

  /** Return a new `p:txBody` element tree, suitable for use in an `p:sp` element. */
  static newPTxBody() {
    const xml = this._pTxBodyTmpl();
    return parseXml(xml);
  }

  /**
   * Return a `c:txPr` element tree.
   *
   * Suitable for use in a chart object like data labels or tick labels.
   */
  static newTxPr() {
    const xml = `<c:txPr ${nsdecls("c", "a")}>
  <a:bodyPr/>
  <a:lstStyle/>
  <a:p>
    <a:pPr>
      <a:defRPr/>
    </a:pPr>
  </a:p>
</c:txPr>
`;
    const txPr = parseXml(xml);
    return txPr;
  }

  /**
   * Ensure p:txBody has at least one a:p child.
   *
   * Intuitively, reverse a ".clear_content()" operation to minimum conformance with spec
   * (single empty paragraph).
   */
  unclearContent() {
    if (this.pLst.length > 0) {
      return;
    }
    this.addP();
  }

  static _aTxBodyTmpl() {
    return `<a:txBody ${nsdecls("a")}>
  <a:bodyPr/>
  <a:p/>
</a:txBody>
`;
  }

  static _pTxBodyTmpl() {
    return `<p:txBody ${nsdecls("p", "a")}>
  <a:bodyPr/>
  <a:p/>
</p:txBody>
`;
  }

  static _txBodyTmpl() {
    return `<p:txBody ${nsdecls("a", "p")}>
  <a:bodyPr/>
  <a:lstStyle/>
  <a:p/>
</p:txBody>
`;
  }
}

/** `a:bodyPr` custom element class. */
export class CT_TextBodyProperties extends BaseOxmlElement {
  declare _addNoAutofit: (...args: any[]) => BaseOxmlElement;
  declare _addNormAutofit: (...args: any[]) => CT_TextNormalAutofit;
  declare _addSpAutoFit: (...args: any[]) => BaseOxmlElement;
  declare _removeEgTextAutoFit: (...args: any[]) => null;
  declare noAutofit: BaseOxmlElement | null;
  declare normAutofit: CT_TextNormalAutofit | null;
  declare spAutoFit: BaseOxmlElement | null;
  declare egTextAutoFit: any;
  declare _newNoAutofit: () => any;
  declare _insertNoAutofit: (child: any) => any;
  declare getOrChangeToNoAutofit: () => any;
  declare _newNormAutofit: () => any;
  declare _insertNormAutofit: (child: any) => any;
  declare getOrChangeToNormAutofit: () => any;
  declare _newSpAutoFit: () => any;
  declare _insertSpAutoFit: (child: any) => any;
  declare getOrChangeToSpAutoFit: () => any;
  declare lIns: number;
  declare tIns: number;
  declare rIns: number;
  declare bIns: number;
  declare anchor: MSO_VERTICAL_ANCHOR | null;
  declare wrap: string | null;
  static {
    defineChildren(this, {
      egTextAutoFit: new ZeroOrOneChoice([new Choice("a:noAutofit"), new Choice("a:normAutofit"), new Choice("a:spAutoFit")], ["a:scene3d", "a:sp3d", "a:flatTx", "a:extLst"]),
      lIns: new OptionalAttribute("lIns", ST_Coordinate32, Emu(91440)),
      tIns: new OptionalAttribute("tIns", ST_Coordinate32, Emu(45720)),
      rIns: new OptionalAttribute("rIns", ST_Coordinate32, Emu(91440)),
      bIns: new OptionalAttribute("bIns", ST_Coordinate32, Emu(45720)),
      anchor: new OptionalAttribute("anchor", MSO_VERTICAL_ANCHOR),
      wrap: new OptionalAttribute("wrap", ST_TextWrappingType),
    });
  }

  /** The autofit setting for the text frame, a member of the `MSO_AUTO_SIZE` enumeration. */
  get autofit() {
    if (this.noAutofit !== null) {
      return MSO_AUTO_SIZE.NONE;
    }
    if (this.normAutofit !== null) {
      return MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE;
    }
    if (this.spAutoFit !== null) {
      return MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT;
    }
    return null;
  }

  set autofit(value: MSO_AUTO_SIZE | null) {
    if ((value !== null && !MSO_AUTO_SIZE.has(value))) {
      throw new ValueError(`only None or a member of the MSO_AUTO_SIZE enumeration can be assigned to CT_TextBodyProperties.autofit, got ${value}`);
    }
    this._removeEgTextAutoFit();
    if (value === MSO_AUTO_SIZE.NONE) {
      this._addNoAutofit();
    } else if (value === MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE) {
      this._addNormAutofit();
    } else if (value === MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT) {
      this._addSpAutoFit();
    }
  }
}

/**
 * Custom element class for `a:rPr`, `a:defRPr`, and `a:endParaRPr`.
 *
 * 'rPr' is short for 'run properties', and it corresponds to the |Font| proxy class.
 */
export class CT_TextCharacterProperties extends BaseOxmlElement {
  declare getOrAddHlinkClick: (...args: any[]) => CT_Hyperlink;
  declare getOrAddLatin: (...args: any[]) => CT_TextFont;
  declare _removeLatin: (...args: any[]) => null;
  declare _removeHlinkClick: (...args: any[]) => null;
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
  declare latin: CT_TextFont | null;
  declare _addLatin: (attrs?: Record<string, any>) => CT_TextFont;
  declare _insertLatin: (child: CT_TextFont) => CT_TextFont;
  declare _newLatin: () => CT_TextFont;
  declare hlinkClick: CT_Hyperlink | null;
  declare _addHlinkClick: (attrs?: Record<string, any>) => CT_Hyperlink;
  declare _insertHlinkClick: (child: CT_Hyperlink) => CT_Hyperlink;
  declare _newHlinkClick: () => CT_Hyperlink;
  declare lang: MSO_LANGUAGE_ID | null;
  declare sz: number | null;
  declare b: boolean | null;
  declare i: boolean | null;
  declare u: MSO_TEXT_UNDERLINE_TYPE | null;
  static {
    defineChildren(this, {
      egFillProperties: new ZeroOrOneChoice([new Choice("a:noFill"), new Choice("a:solidFill"), new Choice("a:gradFill"), new Choice("a:blipFill"), new Choice("a:pattFill"), new Choice("a:grpFill")], ["a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"]),
      latin: new ZeroOrOne("a:latin", ["a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"]),
      hlinkClick: new ZeroOrOne("a:hlinkClick", ["a:hlinkMouseOver", "a:rtl", "a:extLst"]),
      lang: new OptionalAttribute("lang", MSO_LANGUAGE_ID),
      sz: new OptionalAttribute("sz", ST_TextFontSize),
      b: new OptionalAttribute("b", XsdBoolean),
      i: new OptionalAttribute("i", XsdBoolean),
      u: new OptionalAttribute("u", MSO_TEXT_UNDERLINE_TYPE),
    });
  }

  _newGradFill() {
    return CT_GradientFillProperties.newGradFill();
  }

  /** Add an `a:hlinkClick` child element with r:id attribute set to `rId`. */
  addHlinkClick(rId: string): CT_Hyperlink {
    const hlinkClick = this.getOrAddHlinkClick();
    hlinkClick.rId = rId;
    return hlinkClick;
  }
}

/** `a:fld` field element, for either a slide number or date field. */
export class CT_TextField extends BaseOxmlElement {
  declare getOrAddRPr: (...args: any[]) => CT_TextCharacterProperties;
  declare rPr: CT_TextCharacterProperties | null;
  declare _addRPr: (attrs?: Record<string, any>) => CT_TextCharacterProperties;
  declare _insertRPr: (child: CT_TextCharacterProperties) => CT_TextCharacterProperties;
  declare _newRPr: () => CT_TextCharacterProperties;
  declare _removeRPr: () => void;
  declare t: BaseOxmlElement | null;
  declare getOrAddT: () => BaseOxmlElement;
  declare _addT: (attrs?: Record<string, any>) => BaseOxmlElement;
  declare _insertT: (child: BaseOxmlElement) => BaseOxmlElement;
  declare _newT: () => BaseOxmlElement;
  declare _removeT: () => void;
  static {
    defineChildren(this, {
      rPr: new ZeroOrOne("a:rPr", ["a:pPr", "a:t"]),
      t: new ZeroOrOne("a:t", []),
    });
  }

  /** The text of the `a:t` child element. */
  get text(): string {
    const t = this.t;
    if (t === null) {
      return "";
    }
    return (t.text || "");
  }
}

/**
 * Custom element class for `a:latin`, `a:ea`, `a:cs`, and `a:sym`.
 *
 * These occur as child elements of CT_TextCharacterProperties, e.g. `a:rPr`.
 */
export class CT_TextFont extends BaseOxmlElement {
  declare typeface: string;
  static {
    defineChildren(this, {
      typeface: new RequiredAttribute("typeface", ST_TextTypeface),
    });
  }
}

/** `a:br` line break element */
export class CT_TextLineBreak extends BaseOxmlElement {
  declare getOrAddRPr: (...args: any[]) => CT_TextCharacterProperties;
  declare rPr: any;
  declare _addRPr: (attrs?: Record<string, any>) => any;
  declare _insertRPr: (child: any) => any;
  declare _newRPr: () => any;
  declare _removeRPr: () => void;
  static {
    defineChildren(this, {
      rPr: new ZeroOrOne("a:rPr", []),
    });
  }

  /**
   * Unconditionally a single vertical-tab character.
   *
   * A line break element can contain no text other than the implicit line feed it
   * represents.
   */
  get text() {
    return "\u000b";
  }
}

/** `a:normAutofit` element specifying fit text to shape font reduction, etc. */
export class CT_TextNormalAutofit extends BaseOxmlElement {
  declare fontScale: any;
  static {
    defineChildren(this, {
      fontScale: new OptionalAttribute("fontScale", ST_TextFontScalePercentOrPercentString, 100.0),
    });
  }
}

/** `a:p` custom element class */
export class CT_TextParagraph extends BaseOxmlElement {
  declare getOrAddEndParaRPr: (...args: any[]) => CT_TextCharacterProperties;
  declare getOrAddPPr: (...args: any[]) => CT_TextParagraphProperties;
  declare rLst: CT_RegularTextRun[];
  declare _addBr: (...args: any[]) => CT_TextLineBreak;
  declare _addR: (...args: any[]) => CT_RegularTextRun;
  declare pPr: CT_TextParagraphProperties | null;
  declare _addPPr: (attrs?: Record<string, any>) => CT_TextParagraphProperties;
  declare _insertPPr: (child: CT_TextParagraphProperties) => CT_TextParagraphProperties;
  declare _newPPr: () => CT_TextParagraphProperties;
  declare _removePPr: () => void;
  declare _insertR: (child: any) => any;
  declare brLst: any[];
  declare _insertBr: (child: any) => any;
  declare _newBr: () => any;
  declare endParaRPr: CT_TextCharacterProperties | null;
  declare _addEndParaRPr: (attrs?: Record<string, any>) => CT_TextCharacterProperties;
  declare _insertEndParaRPr: (child: CT_TextCharacterProperties) => CT_TextCharacterProperties;
  declare _newEndParaRPr: () => CT_TextCharacterProperties;
  declare _removeEndParaRPr: () => void;
  static {
    defineChildren(this, {
      pPr: new ZeroOrOne("a:pPr", ["a:r", "a:br", "a:fld", "a:endParaRPr"]),
      r: new ZeroOrMore("a:r", ["a:endParaRPr"]),
      br: new ZeroOrMore("a:br", ["a:endParaRPr"]),
      endParaRPr: new ZeroOrOne("a:endParaRPr", []),
    });
  }

  /** Return a newly appended `a:br` element. */
  addBr(): CT_TextLineBreak {
    return this._addBr();
  }

  /** Return a newly appended `a:r` element. */
  addR(text: string | null = null): CT_RegularTextRun {
    const r = this._addR();
    if (text) {
      r.text = text;
    }
    return r;
  }

  /**
   * Append `a:r` and `a:br` elements to `p` based on `text`.
   *
   * Any `\n` or `\v` (vertical-tab) characters in `text` delimit `a:r` (run) elements and
   * themselves are translated to `a:br` (line-break) elements. The vertical-tab character
   * appears in clipboard text from PowerPoint at "soft" line-breaks (new-line, but not new
   *         paragraph).
   */
  appendText(text: string) {
    for (const [idx, rStr] of text.split(/\n|\v/).entries()) {
      if (idx > 0) {
        this.addBr();
      }
      if (rStr) {
        this.addR(rStr);
      }
    }
  }

  /**
   * Sequence containing text-container child elements of this `a:p` element.
   *
   * These include `a:r`, `a:br`, and `a:fld`.
   */
  get contentChildren(): (CT_RegularTextRun | CT_TextLineBreak | CT_TextField)[] {
    return [...this].filter((e): e is CT_RegularTextRun | CT_TextLineBreak | CT_TextField => e instanceof CT_RegularTextRun || e instanceof CT_TextLineBreak || e instanceof CT_TextField);
  }

  /** str text contained in this paragraph. */
  get text(): string {
    return [...this.contentChildren].map((child) => child.text).join("");
  }

  _newR() {
    const rXml = `<a:r ${nsdecls("a")}><a:t/></a:r>`;
    return parseXml(rXml);
  }
}

/** `a:pPr` custom element class. */
export class CT_TextParagraphProperties extends BaseOxmlElement {
  declare getOrAddDefRPr: (...args: any[]) => CT_TextCharacterProperties;
  declare _addLnSpc: (...args: any[]) => CT_TextSpacing;
  declare _addSpcAft: (...args: any[]) => CT_TextSpacing;
  declare _addSpcBef: (...args: any[]) => CT_TextSpacing;
  declare _removeLnSpc: (...args: any[]) => null;
  declare _removeSpcAft: (...args: any[]) => null;
  declare _removeSpcBef: (...args: any[]) => null;
  declare lnSpc: CT_TextSpacing | null;
  declare getOrAddLnSpc: () => CT_TextSpacing;
  declare _insertLnSpc: (child: CT_TextSpacing) => CT_TextSpacing;
  declare _newLnSpc: () => CT_TextSpacing;
  declare spcBef: CT_TextSpacing | null;
  declare getOrAddSpcBef: () => CT_TextSpacing;
  declare _insertSpcBef: (child: CT_TextSpacing) => CT_TextSpacing;
  declare _newSpcBef: () => CT_TextSpacing;
  declare spcAft: CT_TextSpacing | null;
  declare getOrAddSpcAft: () => CT_TextSpacing;
  declare _insertSpcAft: (child: CT_TextSpacing) => CT_TextSpacing;
  declare _newSpcAft: () => CT_TextSpacing;
  declare defRPr: CT_TextCharacterProperties | null;
  declare _addDefRPr: (attrs?: Record<string, any>) => CT_TextCharacterProperties;
  declare _insertDefRPr: (child: CT_TextCharacterProperties) => CT_TextCharacterProperties;
  declare _newDefRPr: () => CT_TextCharacterProperties;
  declare _removeDefRPr: () => void;
  declare lvl: number;
  declare algn: PP_PARAGRAPH_ALIGNMENT | null;
  static {
    const _tagSeq = ["a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct", "a:buSzPts", "a:buFontTx", "a:buFont", "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst"] as const;
    defineChildren(this, {
      lnSpc: new ZeroOrOne("a:lnSpc", _tagSeq.slice(1)),
      spcBef: new ZeroOrOne("a:spcBef", _tagSeq.slice(2)),
      spcAft: new ZeroOrOne("a:spcAft", _tagSeq.slice(3)),
      defRPr: new ZeroOrOne("a:defRPr", _tagSeq.slice(16)),
      lvl: new OptionalAttribute("lvl", ST_TextIndentLevelType, 0),
      algn: new OptionalAttribute("algn", PP_PARAGRAPH_ALIGNMENT),
    });
  }

  /**
   * The spacing between baselines of successive lines in this paragraph.
   *
   * A float value indicates a number of lines. A |Length| value indicates a fixed spacing.
   * Value is contained in `./a:lnSpc/a:spcPts/@val` or `./a:lnSpc/a:spcPct/@val`. Value is
   * |None| if no element is present.
   */
  get lineSpacing(): number | number | null {
    const lnSpc = this.lnSpc;
    if (lnSpc === null) {
      return null;
    }
    if (lnSpc.spcPts !== null) {
      return lnSpc.spcPts.val;
    }
    return lnSpc.spcPct!.val;
  }

  set lineSpacing(value: number | number | null) {
    this._removeLnSpc();
    if (value === null) {
      return;
    }
    if (isLengthValue(value)) {
      this._addLnSpc().setSpcPts(value);
    } else {
      this._addLnSpc().setSpcPct(value);
    }
  }

  /** The EMU equivalent of the centipoints value in `./a:spcAft/a:spcPts/@val`. */
  get spaceAfter(): number | null {
    const spcAft = this.spcAft;
    if (spcAft === null) {
      return null;
    }
    const spcPts = spcAft.spcPts;
    if (spcPts === null) {
      return null;
    }
    return spcPts.val;
  }

  set spaceAfter(value: number | null) {
    this._removeSpcAft();
    if (value !== null) {
      this._addSpcAft().setSpcPts(value);
    }
  }

  /** The EMU equivalent of the centipoints value in `./a:spcBef/a:spcPts/@val`. */
  get spaceBefore() {
    const spcBef = this.spcBef;
    if (spcBef === null) {
      return null;
    }
    const spcPts = spcBef.spcPts;
    if (spcPts === null) {
      return null;
    }
    return spcPts.val;
  }

  set spaceBefore(value: number | null) {
    this._removeSpcBef();
    if (value !== null) {
      this._addSpcBef().setSpcPts(value);
    }
  }
}

/** Used for `a:lnSpc`, `a:spcBef`, and `a:spcAft` elements. */
export class CT_TextSpacing extends BaseOxmlElement {
  declare getOrAddSpcPct: (...args: any[]) => CT_TextSpacingPercent;
  declare getOrAddSpcPts: (...args: any[]) => CT_TextSpacingPoint;
  declare _removeSpcPct: (...args: any[]) => null;
  declare _removeSpcPts: (...args: any[]) => null;
  declare spcPct: CT_TextSpacingPercent | null;
  declare _addSpcPct: (attrs?: Record<string, any>) => CT_TextSpacingPercent;
  declare _insertSpcPct: (child: CT_TextSpacingPercent) => CT_TextSpacingPercent;
  declare _newSpcPct: () => CT_TextSpacingPercent;
  declare spcPts: CT_TextSpacingPoint | null;
  declare _addSpcPts: (attrs?: Record<string, any>) => CT_TextSpacingPoint;
  declare _insertSpcPts: (child: CT_TextSpacingPoint) => CT_TextSpacingPoint;
  declare _newSpcPts: () => CT_TextSpacingPoint;
  static {
    defineChildren(this, {
      spcPct: new ZeroOrOne("a:spcPct"),
      spcPts: new ZeroOrOne("a:spcPts"),
    });
  }

  /**
   * Set spacing to `value` lines, e.g. 1.75 lines.
   *
   * A ./a:spcPts child is removed if present.
   */
  setSpcPct(value: number) {
    this._removeSpcPts();
    const spcPct = this.getOrAddSpcPct();
    spcPct.val = value;
  }

  /** Set spacing to `value` points. A ./a:spcPct child is removed if present. */
  setSpcPts(value: number) {
    this._removeSpcPct();
    const spcPts = this.getOrAddSpcPts();
    spcPts.val = value;
  }
}

/** `a:spcPct` element, specifying spacing in thousandths of a percent in its `val` attribute. */
export class CT_TextSpacingPercent extends BaseOxmlElement {
  declare val: number;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_TextSpacingPercentOrPercentString),
    });
  }
}

/** `a:spcPts` element, specifying spacing in centipoints in its `val` attribute. */
export class CT_TextSpacingPoint extends BaseOxmlElement {
  declare val: number;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_TextSpacingPoint),
    });
  }
}
