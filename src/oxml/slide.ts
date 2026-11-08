// Slide-related custom element classes, including those for masters. (`pptx.oxml.slide`)
import { defineChildren } from "./xmlchemy.js";
import { pyInt } from "../util.js";
import { parseFromTemplate, parseXml } from "./parser.js";
import { CT_GradientFillProperties } from "./dml/fill.js";
import { nsdecls } from "./ns.js";
import { XsdString } from "./simpletypes.js";
import { BaseOxmlElement, Choice, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice } from "./xmlchemy.js";
import type { CT_GroupShape } from "./shapes/groupshape.js";

/** Base class for the six slide types, providing common methods. */
export class _BaseSlideElement extends BaseOxmlElement {
  declare cSld: CT_CommonSlideData;

  /** Return required `p:cSld/p:spTree` grandchild. */
  get spTree(): CT_GroupShape {
    return this.cSld.spTree;
  }
}

/** `p:bg` element. */
export class CT_Background extends BaseOxmlElement {
  declare _insertBgPr: (...args: any[]) => null;
  declare bgPr: CT_BackgroundProperties | null;
  declare getOrAddBgPr: () => CT_BackgroundProperties;
  declare _addBgPr: (attrs?: Record<string, any>) => CT_BackgroundProperties;
  declare _newBgPr: () => CT_BackgroundProperties;
  declare _removeBgPr: () => void;
  declare bgRef: any;
  declare getOrAddBgRef: () => any;
  declare _addBgRef: (attrs?: Record<string, any>) => any;
  declare _insertBgRef: (child: any) => any;
  declare _newBgRef: () => any;
  declare _removeBgRef: () => void;
  static {
    const _tagSeq = ["p:bgPr", "p:bgRef"] as const;
    defineChildren(this, {
      bgPr: new ZeroOrOne("p:bgPr", []),
      bgRef: new ZeroOrOne("p:bgRef", []),
    });
  }

  /** Return a new `p:bgPr` element with noFill properties. */
  addNoFillBgPr() {
    const xml = `<p:bgPr ${nsdecls("a", "p")}>
  <a:noFill/>
  <a:effectLst/>
</p:bgPr>`;
    const bgPr = parseXml(xml);
    this._insertBgPr(bgPr);
    return bgPr;
  }
}

/** `p:bgPr` element. */
export class CT_BackgroundProperties extends BaseOxmlElement {
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
  static {
    const _tagSeq = ["a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:effectLst", "a:effectDag", "a:extLst"] as const;
    defineChildren(this, {
      egFillProperties: new ZeroOrOneChoice([new Choice("a:noFill"), new Choice("a:solidFill"), new Choice("a:gradFill"), new Choice("a:blipFill"), new Choice("a:pattFill"), new Choice("a:grpFill")], _tagSeq.slice(6)),
    });
  }

  /** Override default to add default gradient subtree. */
  _newGradFill() {
    return CT_GradientFillProperties.newGradFill();
  }
}

/** `p:cSld` element. */
export class CT_CommonSlideData extends BaseOxmlElement {
  declare _removeBg: (...args: any[]) => null;
  declare getOrAddBg: (...args: any[]) => CT_Background;
  declare bg: CT_Background | null;
  declare _addBg: (attrs?: Record<string, any>) => CT_Background;
  declare _insertBg: (child: CT_Background) => CT_Background;
  declare _newBg: () => CT_Background;
  declare spTree: CT_GroupShape;
  declare name: string;
  static {
    const _tagSeq = ["p:bg", "p:spTree", "p:custDataLst", "p:controls", "p:extLst"] as const;
    defineChildren(this, {
      bg: new ZeroOrOne("p:bg", _tagSeq.slice(1)),
      spTree: new OneAndOnlyOne("p:spTree"),
      name: new OptionalAttribute("name", XsdString, ""),
    });
  }

  /**
   * Return `p:bg/p:bgPr` grandchild.
   *
   * If no such grandchild is present, any existing `p:bg` child is first removed and a new
   * default `p:bg` with noFill settings is added.
   */
  getOrAddBgPr(): CT_BackgroundProperties {
    let bg = this.bg;
    if ((bg === null || bg.bgPr === null)) {
      bg = this._changeToNoFillBg();
    }
    return bg.bgPr!;
  }

  /**
   * Establish a `p:bg` child with no-fill settings.
   *
   * Any existing `p:bg` child is first removed.
   */
  _changeToNoFillBg(): CT_Background {
    this._removeBg();
    const bg = this.getOrAddBg();
    bg.addNoFillBgPr();
    return bg;
  }
}

/** `p:notesMaster` element, root of a notes master part. */
export class CT_NotesMaster extends _BaseSlideElement {
  declare cSld: CT_CommonSlideData;
  static {
    const _tagSeq = ["p:cSld", "p:clrMap", "p:hf", "p:notesStyle", "p:extLst"] as const;
    defineChildren(this, {
      cSld: new OneAndOnlyOne("p:cSld"),
    });
  }

  /** Return a new `p:notesMaster` element based on the built-in default template. */
  static newDefault(): CT_NotesMaster {
    return parseFromTemplate("notesMaster");
  }
}

/** `p:notes` element, root of a notes slide part. */
export class CT_NotesSlide extends _BaseSlideElement {
  declare cSld: CT_CommonSlideData;
  static {
    const _tagSeq = ["p:cSld", "p:clrMapOvr", "p:extLst"] as const;
    defineChildren(this, {
      cSld: new OneAndOnlyOne("p:cSld"),
    });
  }

  /**
   * Return a new ``<p:notes>`` element based on the default template.
   *
   * Note that the template does not include placeholders, which must be subsequently cloned
   * from the notes master.
   */
  static new(): CT_NotesSlide {
    return parseFromTemplate("notes");
  }
}

/** `p:sld` element, root element of a slide part (XML document). */
export class CT_Slide extends _BaseSlideElement {
  declare cSld: CT_CommonSlideData;
  declare clrMapOvr: any;
  declare getOrAddClrMapOvr: () => any;
  declare _addClrMapOvr: (attrs?: Record<string, any>) => any;
  declare _insertClrMapOvr: (child: any) => any;
  declare _newClrMapOvr: () => any;
  declare _removeClrMapOvr: () => void;
  declare timing: any;
  declare getOrAddTiming: () => any;
  declare _addTiming: (attrs?: Record<string, any>) => any;
  declare _insertTiming: (child: any) => any;
  declare _newTiming: () => any;
  declare _removeTiming: () => void;
  static {
    const _tagSeq = ["p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:extLst"] as const;
    defineChildren(this, {
      cSld: new OneAndOnlyOne("p:cSld"),
      clrMapOvr: new ZeroOrOne("p:clrMapOvr", _tagSeq.slice(2)),
      timing: new ZeroOrOne("p:timing", _tagSeq.slice(4)),
    });
  }

  /** Return new `p:sld` element configured as base slide shape. */
  static new(): CT_Slide {
    return parseXml(this._sldXml());
  }

  /** Return `p:bg` grandchild or None if not present. */
  get bg() {
    return this.cSld.bg;
  }

  /**
   * Return parent element for a new `p:video` child element.
   *
   * The `p:video` element causes play controls to appear under a video
   * shape (pic shape containing video). There can be more than one video
   * shape on a slide, which causes the precondition to vary. It needs to
   * handle the case when there is no `p:sld/p:timing` element and when
   * that element already exists. If the case isn't simple, it just nukes
   * what's there and adds a fresh one. This could theoretically remove
   * desired existing timing information, but there isn't any evidence
   * available to me one way or the other, so I've taken the simple
   * approach.
   */
  getOrAddChildTnLst() {
    let childTnLst = this._childTnLst;
    if (childTnLst === null) {
      childTnLst = this._addChildTnLst();
    }
    return childTnLst;
  }

  /**
   * Add `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.
   *
   * Any existing `p:timing` child element is ruthlessly removed and
   * replaced.
   */
  _addChildTnLst() {
    this.remove(this.getOrAddTiming());
    const timing = parseXml(CT_Slide._childTnLstTimingXml());
    this._insertTiming(timing);
    return timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst")[0];
  }

  /**
   * Return `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.
   *
   * Return None if that element is not present.
   */
  get _childTnLst() {
    const childTnLsts = this.xpath("./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst");
    if (childTnLsts.length === 0) {
      return null;
    }
    return childTnLsts[0];
  }

  static _childTnLstTimingXml() {
    return `<p:timing ${nsdecls("p")}>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst/>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>`;
  }

  static _sldXml() {
    return `<p:sld ${nsdecls("a", "p", "r")}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
  }
}

/** `p:sldLayout` element, root of a slide layout part. */
export class CT_SlideLayout extends _BaseSlideElement {
  declare cSld: CT_CommonSlideData;
  static {
    const _tagSeq = ["p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst"] as const;
    defineChildren(this, {
      cSld: new OneAndOnlyOne("p:cSld"),
    });
  }
}

/**
 * `p:sldLayoutIdLst` element, child of `p:sldMaster`.
 *
 * Contains references to the slide layouts that inherit from the slide master.
 */
export class CT_SlideLayoutIdList extends BaseOxmlElement {
  declare sldLayoutIdLst: CT_SlideLayoutIdListEntry[];
  declare _addSldLayoutId: (attrs?: Record<string, any>) => any;
  declare _insertSldLayoutId: (child: any) => any;
  declare _newSldLayoutId: () => any;
  static {
    defineChildren(this, {
      sldLayoutId: new ZeroOrMore("p:sldLayoutId"),
    });
  }
}

/**
 * `p:sldLayoutId` element, child of `p:sldLayoutIdLst`.
 *
 * Contains a reference to a slide layout.
 */
export class CT_SlideLayoutIdListEntry extends BaseOxmlElement {
  declare rId: string;
  static {
    defineChildren(this, {
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }
}

/** `p:sldMaster` element, root of a slide master part. */
export class CT_SlideMaster extends _BaseSlideElement {
  declare getOrAddSldLayoutIdLst: (...args: any[]) => CT_SlideLayoutIdList;
  declare cSld: CT_CommonSlideData;
  declare sldLayoutIdLst: CT_SlideLayoutIdList;
  declare _addSldLayoutIdLst: (attrs?: Record<string, any>) => CT_SlideLayoutIdList;
  declare _insertSldLayoutIdLst: (child: CT_SlideLayoutIdList) => CT_SlideLayoutIdList;
  declare _newSldLayoutIdLst: () => CT_SlideLayoutIdList;
  declare _removeSldLayoutIdLst: () => void;
  static {
    const _tagSeq = ["p:cSld", "p:clrMap", "p:sldLayoutIdLst", "p:transition", "p:timing", "p:hf", "p:txStyles", "p:extLst"] as const;
    defineChildren(this, {
      cSld: new OneAndOnlyOne("p:cSld"),
      sldLayoutIdLst: new ZeroOrOne("p:sldLayoutIdLst", _tagSeq.slice(3)),
    });
  }
}

/** `p:timing` element, specifying animations and timed behaviors. */
export class CT_SlideTiming extends BaseOxmlElement {
  declare tnLst: any;
  declare getOrAddTnLst: () => any;
  declare _addTnLst: (attrs?: Record<string, any>) => any;
  declare _insertTnLst: (child: any) => any;
  declare _newTnLst: () => any;
  declare _removeTnLst: () => void;
  static {
    const _tagSeq = ["p:tnLst", "p:bldLst", "p:extLst"] as const;
    defineChildren(this, {
      tnLst: new ZeroOrOne("p:tnLst", _tagSeq.slice(1)),
    });
  }
}

/** `p:tnLst` or `p:childTnList` element. */
export class CT_TimeNodeList extends BaseOxmlElement {

  /** Add a new `p:video` child element for movie having *shape_id*. */
  addVideo(shapeId: any) {
    const videoXml = `<p:video ${nsdecls("p")}>
  <p:cMediaNode vol="80000">
    <p:cTn id="${this._nextCTnId}" fill="hold" display="0">
      <p:stCondLst>
        <p:cond delay="indefinite"/>
      </p:stCondLst>
    </p:cTn>
    <p:tgtEl>
      <p:spTgt spid="${shapeId}"/>
    </p:tgtEl>
  </p:cMediaNode>
</p:video>
`;
    const video = parseXml(videoXml);
    this.append(video);
  }

  /** Return the next available unique ID (int) for p:cTn element. */
  get _nextCTnId() {
    const cTnIdStrs = this.xpath("/p:sld/p:timing//p:cTn/@id");
    const ids = (cTnIdStrs as string[]).map((idStr) => pyInt(idStr));
    return (Math.max(...ids) + 1);
  }
}

/** `p:video` element, specifying video media details. */
export class CT_TLMediaNodeVideo extends BaseOxmlElement {
  declare cMediaNode: any;
  static {
    const _tagSeq = ["p:cMediaNode"] as const;
    defineChildren(this, {
      cMediaNode: new OneAndOnlyOne("p:cMediaNode"),
    });
  }
}
