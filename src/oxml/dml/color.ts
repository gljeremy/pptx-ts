// lxml custom element classes for DrawingML-related XML elements. (`pptx.oxml.dml.color`)
import { defineChildren } from "../xmlchemy.js";
import { MSO_THEME_COLOR } from "../../enum/dml.js";
import { ST_HexColorRGB, ST_Percentage } from "../simpletypes.js";
import { BaseOxmlElement, Choice, RequiredAttribute, ZeroOrOne, ZeroOrOneChoice } from "../xmlchemy.js";

/** Base class for <a:srgbClr> and <a:schemeClr> elements. */
export class _BaseColorElement extends BaseOxmlElement {
  declare lumMod: any;
  declare getOrAddLumMod: () => any;
  declare _addLumMod: (attrs?: Record<string, any>) => any;
  declare _insertLumMod: (child: any) => any;
  declare _newLumMod: () => any;
  declare _removeLumMod: () => void;
  declare lumOff: any;
  declare getOrAddLumOff: () => any;
  declare _addLumOff: (attrs?: Record<string, any>) => any;
  declare _insertLumOff: (child: any) => any;
  declare _newLumOff: () => any;
  declare _removeLumOff: () => void;
  static {
    defineChildren(this, {
      lumMod: new ZeroOrOne("a:lumMod"),
      lumOff: new ZeroOrOne("a:lumOff"),
    });
  }

  /** Return a newly added <a:lumMod> child element. */
  addLumMod(value: any) {
    const lumMod = this._addLumMod();
    lumMod.val = value;
    return lumMod;
  }

  /** Return a newly added <a:lumOff> child element. */
  addLumOff(value: any) {
    const lumOff = this._addLumOff();
    lumOff.val = value;
    return lumOff;
  }

  /**
   * Return self after removing any <a:lumMod> and <a:lumOff> child
   * elements.
   */
  clearLum() {
    this._removeLumMod();
    this._removeLumOff();
    return this;
  }
}

/** Custom element class for `a:fgClr`, `a:bgClr` and perhaps others. */
export class CT_Color extends BaseOxmlElement {
  declare egColorChoice: any;
  declare _removeEgColorChoice: () => void;
  declare scrgbClr: any;
  declare _newScrgbClr: () => any;
  declare _insertScrgbClr: (child: any) => any;
  declare _addScrgbClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToScrgbClr: () => any;
  declare srgbClr: any;
  declare _newSrgbClr: () => any;
  declare _insertSrgbClr: (child: any) => any;
  declare _addSrgbClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSrgbClr: () => any;
  declare hslClr: any;
  declare _newHslClr: () => any;
  declare _insertHslClr: (child: any) => any;
  declare _addHslClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToHslClr: () => any;
  declare sysClr: any;
  declare _newSysClr: () => any;
  declare _insertSysClr: (child: any) => any;
  declare _addSysClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSysClr: () => any;
  declare schemeClr: any;
  declare _newSchemeClr: () => any;
  declare _insertSchemeClr: (child: any) => any;
  declare _addSchemeClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToSchemeClr: () => any;
  declare prstClr: any;
  declare _newPrstClr: () => any;
  declare _insertPrstClr: (child: any) => any;
  declare _addPrstClr: (attrs?: Record<string, any>) => any;
  declare getOrChangeToPrstClr: () => any;
  static {
    defineChildren(this, {
      egColorChoice: new ZeroOrOneChoice([new Choice("a:scrgbClr"), new Choice("a:srgbClr"), new Choice("a:hslClr"), new Choice("a:sysClr"), new Choice("a:schemeClr"), new Choice("a:prstClr")], []),
    });
  }
}

/** Custom element class for <a:hslClr> element. */
export class CT_HslColor extends _BaseColorElement {
}

/** Custom element class for <a:lumMod> and <a:lumOff> elements. */
export class CT_Percentage extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_Percentage),
    });
  }
}

/** Custom element class for <a:prstClr> element. */
export class CT_PresetColor extends _BaseColorElement {
}

/** Custom element class for <a:schemeClr> element. */
export class CT_SchemeColor extends _BaseColorElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", MSO_THEME_COLOR),
    });
  }
}

/** Custom element class for <a:scrgbClr> element. */
export class CT_ScRgbColor extends _BaseColorElement {
}

/** Custom element class for <a:srgbClr> element. */
export class CT_SRgbColor extends _BaseColorElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new RequiredAttribute("val", ST_HexColorRGB),
    });
  }
}

/** Custom element class for <a:sysClr> element. */
export class CT_SystemColor extends _BaseColorElement {
}
