// Main presentation object (`pptx.presentation`).
import type { FileTarget } from "./fileio.js";
import type { CT_Presentation } from "./oxml/presentation.js";
import type { CorePropertiesPart } from "./parts/coreprops.js";
import type { PresentationPart } from "./parts/presentation.js";
import { PartElementProxy } from "./shared.js";
import { type NotesMaster, type SlideLayouts, type SlideMaster, SlideMasters, Slides } from "./slide.js";
import { lazy } from "./util.js";

/**
 * PresentationML (PML) presentation.
 *
 * Not intended to be constructed directly. Use `Presentation()` (python `pptx.Presentation`)
 * to open or create a presentation.
 */
export class Presentation extends PartElementProxy {
  declare _element: CT_Presentation;

  get part(): PresentationPart {
    return this._part;
  }

  /**
   * `CoreProperties` instance containing this presentation's core properties.
   *
   * Read/write access to the Dublin Core document properties for the presentation.
   */
  get coreProperties(): CorePropertiesPart {
    return this.part.coreProperties;
  }

  /**
   * Instance of `NotesMaster` for this presentation.
   *
   * If the presentation does not have a notes master, one is created from a default template
   * and returned. The same single instance is returned on each call.
   */
  get notesMaster(): NotesMaster {
    return this.part.notesMaster;
  }

  /**
   * Writes this presentation to `file`.
   *
   * `file` can be either a file-path or a file-like object open for writing bytes.
   */
  save(file: FileTarget): void {
    this.part.save(file);
  }

  /**
   * Height of slides in this presentation, in English Metric Units (EMU).
   *
   * Returns null if no slide height is defined. Read/write.
   */
  get slideHeight(): number | null {
    const sldSz = this._element.sldSz;
    if (sldSz === null) return null;
    return sldSz.cy;
  }
  set slideHeight(height: number) {
    const sldSz = this._element.getOrAddSldSz();
    sldSz.cy = height;
  }

  /** `SlideLayouts` object containing the slide layouts belonging to the first slide master. */
  get slideLayouts(): SlideLayouts {
    return this.slideMasters.at(0).slideLayouts;
  }

  /**
   * First `SlideMaster` object belonging to this presentation.
   *
   * Typically, presentations have only a single slide master. This property provides simpler
   * access in that common case.
   */
  get slideMaster(): SlideMaster {
    return this.slideMasters.at(0);
  }

  /** `SlideMasters` collection of slide-masters belonging to this presentation. */
  get slideMasters(): SlideMasters {
    return lazy(this, "slideMasters", () => new SlideMasters(this._element.getOrAddSldMasterIdLst(), this));
  }

  /**
   * Width of slides in this presentation, in English Metric Units (EMU).
   *
   * Returns null if no slide width is defined. Read/write.
   */
  get slideWidth(): number | null {
    const sldSz = this._element.sldSz;
    if (sldSz === null) return null;
    return sldSz.cx;
  }
  set slideWidth(width: number) {
    const sldSz = this._element.getOrAddSldSz();
    sldSz.cx = width;
  }

  /** `Slides` object containing the slides in this presentation. */
  get slides(): Slides {
    return lazy(this, "slides", () => {
      const sldIdLst = this._element.getOrAddSldIdLst();
      this.part.renameSlideParts(sldIdLst.sldIdLst.map((sldId) => sldId.rId));
      return new Slides(sldIdLst, this);
    });
  }
}
