// lxml custom element classes for DrawingML-related XML elements. (`pptx.oxml.dml.fill`)
import { defineChildren } from "../xmlchemy.js";
import { MSO_PATTERN_TYPE } from "../../enum/dml.js";
import { parseXml } from "../parser.js";
import { nsdecls } from "../ns.js";
import { ST_Percentage, ST_PositiveFixedAngle, ST_PositiveFixedPercentage, ST_RelationshipId } from "../simpletypes.js";
import { BaseOxmlElement, Choice, OneOrMore, OptionalAttribute, RequiredAttribute, ZeroOrOne, ZeroOrOneChoice } from "../xmlchemy.js";

/** <a:blip> element */
export class CT_Blip extends BaseOxmlElement {
  declare rEmbed: any;
  static {
    defineChildren(this, {
      rEmbed: new OptionalAttribute("r:embed", ST_RelationshipId),
    });
  }
}

/** Custom element class for <a:blipFill> element. */
export class CT_BlipFillProperties extends BaseOxmlElement {
  declare blip: any;
  declare getOrAddBlip: () => any;
  declare _addBlip: (attrs?: Record<string, any>) => any;
  declare _insertBlip: (child: any) => any;
  declare _newBlip: () => any;
  declare _removeBlip: () => void;
  declare srcRect: any;
  declare getOrAddSrcRect: () => any;
  declare _addSrcRect: (attrs?: Record<string, any>) => any;
  declare _insertSrcRect: (child: any) => any;
  declare _newSrcRect: () => any;
  declare _removeSrcRect: () => void;
  static {
    const _tagSeq = ["a:blip", "a:srcRect", "a:tile", "a:stretch"] as const;
    defineChildren(this, {
      blip: new ZeroOrOne("a:blip", _tagSeq.slice(1)),
      srcRect: new ZeroOrOne("a:srcRect", _tagSeq.slice(2)),
    });
  }

  /** Set `a:srcRect` child to crop according to *cropping* values. */
  crop(cropping: any) {
    const srcRect = this._addSrcRect();
    [srcRect.l, srcRect.t, srcRect.r, srcRect.b] = cropping;
  }
}

/** `a:gradFill` custom element class. */
export class CT_GradientFillProperties extends BaseOxmlElement {
  declare gsLst: any;
  declare getOrAddGsLst: () => any;
  declare _addGsLst: (attrs?: Record<string, any>) => any;
  declare _insertGsLst: (child: any) => any;
  declare _removeGsLst: () => void;
  declare lin: any;
  declare getOrAddLin: () => any;
  declare _addLin: (attrs?: Record<string, any>) => any;
  declare _insertLin: (child: any) => any;
  declare _newLin: () => any;
  declare _removeLin: () => void;
  declare path: any;
  declare getOrAddPath: () => any;
  declare _addPath: (attrs?: Record<string, any>) => any;
  declare _insertPath: (child: any) => any;
  declare _newPath: () => any;
  declare _removePath: () => void;
  static {
    const _tagSeq = ["a:gsLst", "a:lin", "a:path", "a:tileRect"] as const;
    defineChildren(this, {
      gsLst: new ZeroOrOne("a:gsLst", _tagSeq.slice(1)),
      lin: new ZeroOrOne("a:lin", _tagSeq.slice(2)),
      path: new ZeroOrOne("a:path", _tagSeq.slice(3)),
    });
  }

  /** Return newly-created "loose" default gradient subtree. */
  static newGradFill() {
    return parseXml(`<a:gradFill ${nsdecls("a")} rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0">
      <a:schemeClr val="accent1">
        <a:tint val="100000"/>
        <a:shade val="100000"/>
        <a:satMod val="130000"/>
      </a:schemeClr>
    </a:gs>
    <a:gs pos="100000">
      <a:schemeClr val="accent1">
        <a:tint val="50000"/>
        <a:shade val="100000"/>
        <a:satMod val="350000"/>
      </a:schemeClr>
    </a:gs>
  </a:gsLst>
  <a:lin scaled="0"/>
</a:gradFill>
`);
  }

  /** Override default to add minimum subtree. */
  _newGsLst() {
    return CT_GradientStopList.newGsLst();
  }
}

/** `a:gs` custom element class. */
export class CT_GradientStop extends BaseOxmlElement {
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
  declare pos: any;
  static {
    defineChildren(this, {
      egColorChoice: new ZeroOrOneChoice([new Choice("a:scrgbClr"), new Choice("a:srgbClr"), new Choice("a:hslClr"), new Choice("a:sysClr"), new Choice("a:schemeClr"), new Choice("a:prstClr")], []),
      pos: new RequiredAttribute("pos", ST_PositiveFixedPercentage),
    });
  }
}

/** `a:gsLst` custom element class. */
export class CT_GradientStopList extends BaseOxmlElement {
  declare gsLst: any[];
  declare _addGs: (attrs?: Record<string, any>) => any;
  declare _insertGs: (child: any) => any;
  declare _newGs: () => any;
  declare addGs: () => any;
  static {
    defineChildren(this, {
      gs: new OneOrMore("a:gs"),
    });
  }

  /**
   * Return newly-created "loose" default stop-list subtree.
   *
   * An `a:gsLst` element must have at least two `a:gs` children. These
   * are the default from the PowerPoint built-in "White" template.
   */
  static newGsLst() {
    return parseXml(`<a:gsLst ${nsdecls("a")}>
  <a:gs pos="0">
    <a:schemeClr val="accent1">
      <a:tint val="100000"/>
      <a:shade val="100000"/>
      <a:satMod val="130000"/>
    </a:schemeClr>
  </a:gs>
  <a:gs pos="100000">
    <a:schemeClr val="accent1">
      <a:tint val="50000"/>
      <a:shade val="100000"/>
      <a:satMod val="350000"/>
    </a:schemeClr>
  </a:gs>
</a:gsLst>
`);
  }
}

/** `a:grpFill` custom element class */
export class CT_GroupFillProperties extends BaseOxmlElement {
}

/** `a:lin` custom element class */
export class CT_LinearShadeProperties extends BaseOxmlElement {
  declare ang: any;
  static {
    defineChildren(this, {
      ang: new OptionalAttribute("ang", ST_PositiveFixedAngle),
    });
  }
}

/** `a:noFill` custom element class */
export class CT_NoFillProperties extends BaseOxmlElement {
}

/** `a:pattFill` custom element class */
export class CT_PatternFillProperties extends BaseOxmlElement {
  declare fgClr: any;
  declare getOrAddFgClr: () => any;
  declare _addFgClr: (attrs?: Record<string, any>) => any;
  declare _insertFgClr: (child: any) => any;
  declare _removeFgClr: () => void;
  declare bgClr: any;
  declare getOrAddBgClr: () => any;
  declare _addBgClr: (attrs?: Record<string, any>) => any;
  declare _insertBgClr: (child: any) => any;
  declare _removeBgClr: () => void;
  declare prst: any;
  static {
    const _tagSeq = ["a:fgClr", "a:bgClr"] as const;
    defineChildren(this, {
      fgClr: new ZeroOrOne("a:fgClr", _tagSeq.slice(1)),
      bgClr: new ZeroOrOne("a:bgClr", _tagSeq.slice(2)),
      prst: new OptionalAttribute("prst", MSO_PATTERN_TYPE),
    });
  }

  /** Override default to add minimum subtree. */
  _newBgClr() {
    const xml = `<a:bgClr ${nsdecls("a")}>
 <a:srgbClr val="FFFFFF"/>
</a:bgClr>
`;
    const bgClr = parseXml(xml);
    return bgClr;
  }

  /** Override default to add minimum subtree. */
  _newFgClr() {
    const xml = `<a:fgClr ${nsdecls("a")}>
 <a:srgbClr val="000000"/>
</a:fgClr>
`;
    const fgClr = parseXml(xml);
    return fgClr;
  }
}

/** `a:srcRect` element and perhaps others. */
export class CT_RelativeRect extends BaseOxmlElement {
  declare l: any;
  declare t: any;
  declare r: any;
  declare b: any;
  static {
    defineChildren(this, {
      l: new OptionalAttribute("l", ST_Percentage, 0.0),
      t: new OptionalAttribute("t", ST_Percentage, 0.0),
      r: new OptionalAttribute("r", ST_Percentage, 0.0),
      b: new OptionalAttribute("b", ST_Percentage, 0.0),
    });
  }
}

/** `a:solidFill` custom element class. */
export class CT_SolidColorFillProperties extends BaseOxmlElement {
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
