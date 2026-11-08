// Custom element classes for presentation-related XML elements. (`pptx.oxml.presentation`)
import { defineChildren } from "./xmlchemy.js";
import { pyInt } from "../util.js";
import { ST_SlideId, ST_SlideSizeCoordinate, XsdString } from "./simpletypes.js";
import { BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne } from "./xmlchemy.js";
import type { Length } from "../util.js";

/** `p:presentation` element, root of the Presentation part stored as `/ppt/presentation.xml`. */
export class CT_Presentation extends BaseOxmlElement {
  declare getOrAddSldSz: (...args: any[]) => CT_SlideSize;
  declare getOrAddSldIdLst: (...args: any[]) => CT_SlideIdList;
  declare getOrAddSldMasterIdLst: (...args: any[]) => CT_SlideMasterIdList;
  declare sldMasterIdLst: CT_SlideMasterIdList | null;
  declare _addSldMasterIdLst: (attrs?: Record<string, any>) => CT_SlideMasterIdList;
  declare _insertSldMasterIdLst: (child: CT_SlideMasterIdList) => CT_SlideMasterIdList;
  declare _newSldMasterIdLst: () => CT_SlideMasterIdList;
  declare _removeSldMasterIdLst: () => void;
  declare sldIdLst: CT_SlideIdList | null;
  declare _addSldIdLst: (attrs?: Record<string, any>) => CT_SlideIdList;
  declare _insertSldIdLst: (child: CT_SlideIdList) => CT_SlideIdList;
  declare _newSldIdLst: () => CT_SlideIdList;
  declare _removeSldIdLst: () => void;
  declare sldSz: CT_SlideSize | null;
  declare _addSldSz: (attrs?: Record<string, any>) => CT_SlideSize;
  declare _insertSldSz: (child: CT_SlideSize) => CT_SlideSize;
  declare _newSldSz: () => CT_SlideSize;
  declare _removeSldSz: () => void;
  static {
    defineChildren(this, {
      sldMasterIdLst: new ZeroOrOne("p:sldMasterIdLst", ["p:notesMasterIdLst", "p:handoutMasterIdLst", "p:sldIdLst", "p:sldSz", "p:notesSz"]),
      sldIdLst: new ZeroOrOne("p:sldIdLst", ["p:sldSz", "p:notesSz"]),
      sldSz: new ZeroOrOne("p:sldSz", ["p:notesSz"]),
    });
  }
}

/**
 * `p:sldId` element.
 *
 * Direct child of `p:sldIdLst` that contains an `rId` reference to a slide in the presentation.
 */
export class CT_SlideId extends BaseOxmlElement {
  declare id: number;
  declare rId: string;
  static {
    defineChildren(this, {
      id: new RequiredAttribute("id", ST_SlideId),
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }
}

/**
 * `p:sldIdLst` element.
 *
 * Direct child of <p:presentation> that contains a list of the slide parts in the presentation.
 */
export class CT_SlideIdList extends BaseOxmlElement {
  declare sldIdLst: CT_SlideId[];
  declare _addSldId: (...args: any[]) => CT_SlideId;
  declare _insertSldId: (child: any) => any;
  declare _newSldId: () => any;
  static {
    defineChildren(this, {
      sldId: new ZeroOrMore("p:sldId"),
    });
  }

  /**
   * Create and return a reference to a new `p:sldId` child element.
   *
   * The new `p:sldId` element has its r:id attribute set to `rId`.
   */
  addSldId(rId: string): CT_SlideId {
    return this._addSldId({ id: this._nextId, rId: rId });
  }

  /**
   * The next available slide ID as an `int`.
   *
   * Valid slide IDs start at 256. The next integer value greater than the max value in use is
   * chosen, which minimizes that chance of reusing the id of a deleted slide.
   */
  get _nextId(): number {
    const MIN_SLIDE_ID = 256;
    const MAX_SLIDE_ID = 2147483647;
    const usedIds = (this.xpath("./p:sldId/@id") as string[]).map((s) => pyInt(s));
    const simpleNext = Math.max(MIN_SLIDE_ID - 1, ...usedIds) + 1;
    if (simpleNext <= MAX_SLIDE_ID) {
      return simpleNext;
    }
    // fall back to search for next unused from bottom
    const validUsedIds = usedIds.filter((id) => MIN_SLIDE_ID <= id && id <= MAX_SLIDE_ID).sort((a, b) => a - b);
    if (validUsedIds.length === 0) return 256;
    const i = validUsedIds.findIndex((usedId, idx) => MIN_SLIDE_ID + idx !== usedId);
    if (i === -1) throw new Error("StopIteration");
    return MIN_SLIDE_ID + i;
  }
}

/**
 * `p:sldMasterIdLst` element.
 *
 * Child of `p:presentation` containing references to the slide masters that belong to the
 * presentation.
 */
export class CT_SlideMasterIdList extends BaseOxmlElement {
  declare sldMasterIdLst: CT_SlideMasterIdListEntry[];
  declare _addSldMasterId: (attrs?: Record<string, any>) => any;
  declare _insertSldMasterId: (child: any) => any;
  declare _newSldMasterId: () => any;
  static {
    defineChildren(this, {
      sldMasterId: new ZeroOrMore("p:sldMasterId"),
    });
  }
}

/**
 * ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
 * a reference to a slide master.
 */
export class CT_SlideMasterIdListEntry extends BaseOxmlElement {
  declare rId: string;
  static {
    defineChildren(this, {
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }
}

/**
 * `p:sldSz` element.
 *
 * Direct child of <p:presentation> that contains the width and height of slides in the
 * presentation.
 */
export class CT_SlideSize extends BaseOxmlElement {
  declare cx: number;
  declare cy: number;
  static {
    defineChildren(this, {
      cx: new RequiredAttribute("cx", ST_SlideSizeCoordinate),
      cy: new RequiredAttribute("cy", ST_SlideSizeCoordinate),
    });
  }
}
