// Presentation part, the main part in a .pptx package. (`pptx.parts.presentation`)
import { ValueError, KeyError } from "../errors.js";
import { lazy } from "../util.js";
import { RELATIONSHIP_TYPE as RT } from "../opc/constants.js";
import { XmlPart } from "../opc/package.js";
import { PackURI } from "../opc/packuri.js";
import { NotesMasterPart, SlidePart } from "./slide.js";
import { Presentation } from "../presentation.js";
import type { FileTarget } from "../fileio.js";
import type { CT_Presentation } from "../oxml/presentation.js";
import type { Package } from "../package.js";
import type { CorePropertiesPart } from "./coreprops.js";
import type { NotesMaster, Slide, SlideLayout, SlideMaster } from "../slide.js";

/**
 * Top level class in object model.
 *
 * Represents the contents of the /ppt directory of a .pptx file.
 */
export class PresentationPart extends XmlPart {
  declare _element: CT_Presentation;
  declare _package: Package;

  get package(): Package {
    return this._package;
  }

  /**
   * Return (rId, slide) pair of a newly created blank slide.
   *
   * New slide inherits appearance from `slide_layout`.
   */
  addSlide(slideLayout: SlideLayout): [string, Slide] {
    const partname = this._nextSlidePartname;
    const slideLayoutPart = slideLayout.part;
    const slidePart = SlidePart.new(partname, this.package, slideLayoutPart);
    const rId = this.relateTo(slidePart, RT.SLIDE);
    return [rId, slidePart.slide];
  }

  /**
   * A |CoreProperties| object for the presentation.
   *
   * Provides read/write access to the Dublin Core properties of this presentation.
   */
  get coreProperties(): CorePropertiesPart {
    return this.package.coreProperties;
  }

  /**
   * Return optional related |Slide| object identified by `slide_id`.
   *
   * Returns |None| if no slide with `slide_id` is related to this presentation.
   */
  getSlide(slideId: number): Slide | null {
    for (const sldId of this._element.sldIdLst!.sldIdLst) {
      if (sldId.id === slideId) {
        return this.relatedPart(sldId.rId).slide;
      }
    }
    return null;
  }

  /**
   * Return the |NotesMaster| object for this presentation. If the
   * presentation does not have a notes master, one is created from
   * a default template. The same single instance is returned on each
   * call.
   */
  get notesMaster(): NotesMaster {
    return lazy(this, "notesMaster", () => {
      return this.notesMasterPart.notesMaster;
    });
  }

  /**
   * Return the |NotesMasterPart| object for this presentation.
   *
   * If the presentation does not have a notes master, one is created from a default template.
   * The same single instance is returned on each call.
   */
  get notesMasterPart(): NotesMasterPart {
    return lazy(this, "notesMasterPart", () => {
      try {
        return this.partRelatedBy(RT.NOTES_MASTER);
      } catch (e) {
        if (!(e instanceof KeyError)) throw e;
        const notesMasterPart = NotesMasterPart.createDefault(this.package);
        this.relateTo(notesMasterPart, RT.NOTES_MASTER);
        return notesMasterPart;
      }
    });
  }

  /**
   * A |Presentation| object providing access to the content of this
   * presentation.
   */
  get presentation(): Presentation {
    return lazy(this, "presentation", () => {
      return new Presentation(this._element, this);
    });
  }

  /** Return |Slide| object for related |SlidePart| related by `rId`. */
  relatedSlide(rId: string): Slide {
    return this.relatedPart(rId).slide;
  }

  /** Return |SlideMaster| object for |SlideMasterPart| related by `rId`. */
  relatedSlideMaster(rId: string): SlideMaster {
    return this.relatedPart(rId).slideMaster;
  }

  /**
   * Assign incrementing partnames to the slide parts identified by `rIds`.
   *
   * Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
   * in the `rIds` sequence. The name portion is always `slide`. The number part forms a
   * continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
   * `.xml`.
   */
  renameSlideParts(rIds: string[]): void {
    for (const [idx, rId] of rIds.entries()) {
      const slidePart = this.relatedPart(rId);
      slidePart.partname = new PackURI(`/ppt/slides/slide${(idx + 1)}.xml`);
    }
  }

  /**
   * Save this presentation package to `path_or_stream`.
   *
   * `path_or_stream` can be either a path to a filesystem location (a string) or a
   * file-like object.
   */
  save(pathOrStream: FileTarget): void {
    this.package.save(pathOrStream);
  }

  /** Return the slide-id associated with `slide_part`. */
  slideId(slidePart: SlidePart): number {
    for (const sldId of this._element.sldIdLst!.sldIdLst) {
      if (this.relatedPart(sldId.rId) === slidePart) {
        return sldId.id;
      }
    }
    throw new ValueError("matching slide_part not found");
  }

  /** Return |PackURI| instance containing next available slide partname. */
  get _nextSlidePartname(): PackURI {
    const sldIdLst = this._element.getOrAddSldIdLst();
    const partnameStr = `/ppt/slides/slide${sldIdLst.length + 1}.xml`;
    return new PackURI(partnameStr);
  }
}
