// Slide-related objects, including masters, layouts, and notes (`pptx.slide`).
//
// Collections support `for...of`, `.length` and indexed access with `.at(idx)` (python
// `slides[idx]`; negative indexes count from the end as in python).
import { FillFormat } from "./dml/fill.js";
import { IndexError, ValueError } from "./errors.js";
import { PP_PLACEHOLDER } from "./enum/shapes.js";
import type { CT_SlideIdList, CT_SlideMasterIdList } from "./oxml/presentation.js";
import type { CT_CommonSlideData, CT_NotesMaster, CT_NotesSlide, CT_Slide, CT_SlideLayout, CT_SlideLayoutIdList, CT_SlideMaster } from "./oxml/slide.js";
import { PY_GETITEM } from "./oxml/xml.js";
import type { PresentationPart } from "./parts/presentation.js";
import type { NotesMasterPart, NotesSlidePart, SlideLayoutPart, SlideMasterPart, SlidePart } from "./parts/slide.js";
import type { Presentation } from "./presentation.js";
import type { LayoutPlaceholder, MasterPlaceholder, NotesSlidePlaceholder } from "./shapes/placeholder.js";
import { LayoutPlaceholders, LayoutShapes, MasterPlaceholders, MasterShapes, NotesSlidePlaceholders, NotesSlideShapes, SlidePlaceholders, SlideShapes } from "./shapes/shapetree.js";
import { ElementProxy, ParentedElementProxy, PartElementProxy, proxyEquals } from "./shared.js";
import type { TextFrame } from "./text/text.js";
import { lazy } from "./util.js";

type SlideElement = CT_Slide | CT_SlideLayout | CT_SlideMaster | CT_NotesSlide | CT_NotesMaster;

/** Base class for slide objects, including masters, layouts and notes. */
export class _BaseSlide extends PartElementProxy {
  declare _element: SlideElement;

  /**
   * `_Background` object providing slide background properties.
   *
   * This property returns a `_Background` object whether or not the slide, master, or layout
   * has an explicitly defined background. The same `_Background` object is returned on every
   * call for the same slide object.
   */
  get background(): _Background {
    return lazy(this, "background", () => new _Background(this._element.cSld));
  }

  /**
   * String representing the internal name of this slide.
   *
   * Returns an empty string (`''`) if no name is assigned. Assigning an empty string or null
   * to this property causes any name to be removed.
   */
  get name(): string {
    return this._element.cSld.name;
  }
  set name(value: string | null) {
    this._element.cSld.name = value === null ? "" : value;
  }
}

/** Base object for master objects such as slide master and notes master. */
export class _BaseMaster extends _BaseSlide {
  declare _element: CT_SlideMaster | CT_NotesMaster;

  /**
   * Instance of `MasterPlaceholders` containing sequence of placeholder shapes in this master,
   * sorted in `idx` order.
   */
  get placeholders(): MasterPlaceholders {
    return lazy(this, "placeholders", () => new MasterPlaceholders(this._element.spTree, this));
  }

  /** Instance of `MasterShapes` containing sequence of shape objects appearing on this slide. */
  get shapes(): MasterShapes {
    return lazy(this, "shapes", () => new MasterShapes(this._element.spTree, this));
  }
}

/**
 * Proxy for the notes master XML document.
 *
 * Provides access to shapes, the most commonly used of which are placeholders.
 */
export class NotesMaster extends _BaseMaster {
  declare _element: CT_NotesMaster;

  get part(): NotesMasterPart {
    return this._part;
  }
}

/**
 * Notes slide object.
 *
 * Provides access to slide notes placeholder and other shapes on the notes handout page.
 */
export class NotesSlide extends _BaseSlide {
  declare _element: CT_NotesSlide;

  get part(): NotesSlidePart {
    return this._part;
  }

  /**
   * Selectively add placeholder shape elements from `notesMaster`.
   *
   * Only placeholders of the slide-image, body and slide-number types are cloned; other
   * placeholders in `notesMaster` are ignored.
   */
  cloneMasterPlaceholders(notesMaster: NotesMaster): void {
    const cloneable = [PP_PLACEHOLDER.SLIDE_IMAGE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SLIDE_NUMBER];
    const shapes = this.shapes;
    for (const placeholder of notesMaster.placeholders) {
      if (cloneable.includes(placeholder.element.phType)) shapes.clonePlaceholder(placeholder as MasterPlaceholder);
    }
  }

  /**
   * The notes placeholder on this notes slide, the shape that contains the actual notes text.
   *
   * null if no notes placeholder is present; while this is probably uncommon, it can happen if
   * the notes master does not have a body placeholder, or if the notes placeholder has been
   * deleted from the notes slide.
   */
  get notesPlaceholder(): NotesSlidePlaceholder | null {
    for (const placeholder of this.placeholders) {
      if (placeholder.placeholderFormat.type === PP_PLACEHOLDER.BODY) return placeholder as NotesSlidePlaceholder;
    }
    return null;
  }

  /**
   * The text frame of the notes placeholder on this notes slide, or null if there is no notes
   * placeholder. This is a shortcut to accommodate the common case of simply adding "notes"
   * text to the notes "page".
   */
  get notesTextFrame(): TextFrame | null {
    const notesPlaceholder = this.notesPlaceholder;
    if (notesPlaceholder === null) return null;
    return notesPlaceholder.textFrame;
  }

  /**
   * Instance of `NotesSlidePlaceholders` for this notes-slide.
   *
   * Contains the sequence of placeholder shapes in this notes slide.
   */
  get placeholders(): NotesSlidePlaceholders {
    return lazy(this, "placeholders", () => new NotesSlidePlaceholders(this.element.spTree, this));
  }

  /** Sequence of shape objects appearing on this notes slide. */
  get shapes(): NotesSlideShapes {
    return lazy(this, "shapes", () => new NotesSlideShapes(this._element.spTree, this));
  }
}

/** Slide object. Provides access to shapes and slide-level properties. */
export class Slide extends _BaseSlide {
  declare _element: CT_Slide;

  get part(): SlidePart {
    return this._part;
  }

  /**
   * True if this slide inherits the slide master background.
   *
   * Assigning false causes background inheritance from the master to be interrupted; if there
   * is no custom background for this slide, a default background is added. If a custom
   * background already exists for this slide, assigning false has no effect.
   *
   * Assigning true causes any custom background for this slide to be deleted and inheritance
   * from the master restored.
   */
  get followMasterBackground(): boolean {
    return this._element.bg === null;
  }

  /**
   * True if this slide already has a notes slide, false otherwise. When no notes slide is
   * present, reading `Slide.notesSlide` creates one.
   */
  get hasNotesSlide(): boolean {
    return this.part.hasNotesSlide;
  }

  /**
   * The `NotesSlide` instance for this slide.
   *
   * If the slide does not have a notes slide, one is created. The same single instance is
   * returned on each call.
   */
  get notesSlide(): NotesSlide {
    return this.part.notesSlide;
  }

  /** Sequence of placeholder shapes in this slide. */
  get placeholders(): SlidePlaceholders {
    return lazy(this, "placeholders", () => new SlidePlaceholders(this._element.spTree, this));
  }

  /** Sequence of shape objects appearing on this slide. */
  get shapes(): SlideShapes {
    return lazy(this, "shapes", () => new SlideShapes(this._element.spTree, this));
  }

  /**
   * Integer value that uniquely identifies this slide within this presentation.
   *
   * The slide id does not change if the position of this slide in the slide sequence is
   * changed by adding, rearranging, or deleting slides.
   */
  get slideId(): number {
    return this.part.slideId;
  }

  /** `SlideLayout` object this slide inherits appearance from. */
  get slideLayout(): SlideLayout {
    return this.part.slideLayout;
  }
}

/**
 * Sequence of slides belonging to an instance of `Presentation`.
 *
 * Has list semantics for access to individual slides. Supports indexed access, `.length`,
 * and iteration.
 */
export class Slides extends ParentedElementProxy implements Iterable<Slide> {
  declare _sldIdLst: CT_SlideIdList;

  constructor(sldIdLst: CT_SlideIdList, prs: Presentation) {
    super(sldIdLst, prs);
    this._sldIdLst = sldIdLst;
  }

  get part(): PresentationPart {
    return this._parent.part;
  }

  /** Indexed access, e.g. `slides.at(0)` (python `slides[0]`). */
  at(idx: number): Slide {
    const sldId = this._sldIdLst.sldIdLst.at(idx);
    if (sldId === undefined) throw new IndexError("slide index out of range");
    return this.part.relatedSlide(sldId.rId);
  }

  [PY_GETITEM](idx: number): Slide {
    return this.at(idx);
  }

  /** Support iteration, e.g. `for (const slide of slides)`. */
  *[Symbol.iterator](): Generator<Slide> {
    for (const sldId of this._sldIdLst.sldIdLst) yield this.part.relatedSlide(sldId.rId);
  }

  /** Number of slides, python `len(slides)`. */
  get length(): number {
    return this._sldIdLst.sldIdLst.length;
  }

  /** A newly added slide that inherits layout from `slideLayout`. */
  addSlide(slideLayout: SlideLayout): Slide {
    const [rId, slide] = this.part.addSlide(slideLayout);
    slide.shapes.cloneLayoutPlaceholders(slideLayout);
    this._sldIdLst.addSldId(rId);
    return slide;
  }

  /** The slide identified by integer `slideId` in this presentation, or `dflt` if not found. */
  get(slideId: number, dflt: Slide | null = null): Slide | null {
    const slide = this.part.getSlide(slideId);
    return slide === null ? dflt : slide;
  }

  /** Map `slide` to an integer representing its zero-based position in this slide collection. */
  index(slide: Slide): number {
    for (const [idx, thisSlide] of [...this].entries()) {
      if (proxyEquals(thisSlide, slide)) return idx;
    }
    throw new ValueError(`${slide} is not in slide collection`);
  }
}

/**
 * Slide layout object.
 *
 * Provides access to placeholders, regular shapes, and slide layout-level properties.
 */
export class SlideLayout extends _BaseSlide {
  declare _element: CT_SlideLayout;

  get part(): SlideLayoutPart {
    return this._part;
  }

  /**
   * Generate layout-placeholders on this slide-layout that should be cloned to a new slide.
   *
   * Used when creating a new slide from this slide-layout.
   */
  *iterCloneablePlaceholders(): Generator<LayoutPlaceholder> {
    const latentPhTypes = [PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER];
    for (const ph of this.placeholders) {
      if (!latentPhTypes.includes(ph.element.phType)) yield ph as LayoutPlaceholder;
    }
  }

  /** Sequence of placeholder shapes in this slide layout, sorted in `idx` order. */
  get placeholders(): LayoutPlaceholders {
    return lazy(this, "placeholders", () => new LayoutPlaceholders(this._element.spTree, this));
  }

  /** Sequence of shapes appearing on this slide layout. */
  get shapes(): LayoutShapes {
    return lazy(this, "shapes", () => new LayoutShapes(this._element.spTree, this));
  }

  /** Slide master from which this slide-layout inherits properties. */
  get slideMaster(): SlideMaster {
    return this.part.slideMaster;
  }

  /** Slides based on this slide layout. */
  get usedBySlides(): Slide[] {
    // getting Slides collection requires going around the horn a bit
    const slides = this.part.package.presentationPart.presentation.slides;
    return [...slides].filter((s) => proxyEquals(s.slideLayout, this));
  }
}

/**
 * Sequence of slide layouts belonging to a slide-master.
 *
 * Supports indexed access, `.length`, iteration, `index()` and `remove()`.
 */
export class SlideLayouts extends ParentedElementProxy implements Iterable<SlideLayout> {
  declare _sldLayoutIdLst: CT_SlideLayoutIdList;

  constructor(sldLayoutIdLst: CT_SlideLayoutIdList, parent: SlideMaster) {
    super(sldLayoutIdLst, parent);
    this._sldLayoutIdLst = sldLayoutIdLst;
  }

  get part(): SlideMasterPart {
    return this._parent.part;
  }

  /** Indexed access, e.g. `slideLayouts.at(2)` (python `slide_layouts[2]`). */
  at(idx: number): SlideLayout {
    const sldLayoutId = this._sldLayoutIdLst.sldLayoutIdLst.at(idx);
    if (sldLayoutId === undefined) throw new IndexError("slide layout index out of range");
    return this.part.relatedSlideLayout(sldLayoutId.rId);
  }

  [PY_GETITEM](idx: number): SlideLayout {
    return this.at(idx);
  }

  /** Generate each `SlideLayout` in the collection, in sequence. */
  *[Symbol.iterator](): Generator<SlideLayout> {
    for (const sldLayoutId of this._sldLayoutIdLst.sldLayoutIdLst) yield this.part.relatedSlideLayout(sldLayoutId.rId);
  }

  /** Number of slide layouts, python `len(slide_layouts)`. */
  get length(): number {
    return this._sldLayoutIdLst.sldLayoutIdLst.length;
  }

  /** SlideLayout object having `name`, or `dflt` if not found. */
  getByName(name: string, dflt: SlideLayout | null = null): SlideLayout | null {
    for (const slideLayout of this) {
      if (slideLayout.name === name) return slideLayout;
    }
    return dflt;
  }

  /** Zero-based index of `slideLayout` in this collection; ValueError if not present. */
  index(slideLayout: SlideLayout): number {
    for (const [idx, thisLayout] of [...this].entries()) {
      if (proxyEquals(slideLayout, thisLayout)) return idx;
    }
    throw new ValueError("layout not in this SlideLayouts collection");
  }

  /**
   * Remove `slideLayout` from the collection.
   *
   * Raises ValueError when `slideLayout` is in use; a slide layout which is the basis for one
   * or more slides cannot be removed.
   */
  remove(slideLayout: SlideLayout): void {
    // a layout that is in use cannot be deleted
    if (slideLayout.usedBySlides.length > 0) throw new ValueError("cannot remove slide-layout in use by one or more slides");
    // target layout is identified by its index in this collection
    const targetIdx = this.index(slideLayout);
    // remove layout from p:sldLayoutIds of its master; this stops reference from
    // `p:sldMaster` to its layout part
    const targetSldLayoutId = this._sldLayoutIdLst.sldLayoutIdLst[targetIdx];
    this._sldLayoutIdLst.remove(targetSldLayoutId);
    // drop relationship from master to layout; this removes the layout part from the package
    slideLayout.slideMaster.part.dropRel(targetSldLayoutId.rId);
  }
}

/**
 * Slide master object.
 *
 * Provides access to slide layouts. Access to placeholders, regular shapes, and slide
 * master-level properties is inherited from `_BaseMaster`.
 */
export class SlideMaster extends _BaseMaster {
  declare _element: CT_SlideMaster;

  get part(): SlideMasterPart {
    return this._part;
  }

  /** `SlideLayouts` object providing access to this slide-master's layouts. */
  get slideLayouts(): SlideLayouts {
    return lazy(this, "slideLayouts", () => new SlideLayouts(this._element.getOrAddSldLayoutIdLst(), this));
  }
}

/**
 * Sequence of `SlideMaster` objects belonging to a presentation.
 *
 * Has list access semantics, supporting indexed access, `.length`, and iteration.
 */
export class SlideMasters extends ParentedElementProxy implements Iterable<SlideMaster> {
  declare _sldMasterIdLst: CT_SlideMasterIdList;

  constructor(sldMasterIdLst: CT_SlideMasterIdList, parent: Presentation) {
    super(sldMasterIdLst, parent);
    this._sldMasterIdLst = sldMasterIdLst;
  }

  get part(): PresentationPart {
    return this._parent.part;
  }

  /** Indexed access, e.g. `slideMasters.at(2)` (python `slide_masters[2]`). */
  at(idx: number): SlideMaster {
    const sldMasterId = this._sldMasterIdLst.sldMasterIdLst.at(idx);
    if (sldMasterId === undefined) throw new IndexError("slide master index out of range");
    return this.part.relatedSlideMaster(sldMasterId.rId);
  }

  [PY_GETITEM](idx: number): SlideMaster {
    return this.at(idx);
  }

  /** Generate each `SlideMaster` instance in the collection, in sequence. */
  *[Symbol.iterator](): Generator<SlideMaster> {
    for (const smi of this._sldMasterIdLst.sldMasterIdLst) yield this.part.relatedSlideMaster(smi.rId);
  }

  /** Number of slide masters, python `len(slide_masters)`. */
  get length(): number {
    return this._sldMasterIdLst.sldMasterIdLst.length;
  }
}

/**
 * Provides access to slide background properties.
 *
 * Note that the presence of this object does not by itself imply an explicitly-defined
 * background; a slide with an inherited background still has a `_Background` object.
 */
export class _Background extends ElementProxy {
  declare _cSld: CT_CommonSlideData;

  constructor(cSld: CT_CommonSlideData) {
    super(cSld);
    this._cSld = cSld;
  }

  /**
   * `FillFormat` instance for this background.
   *
   * This `FillFormat` object is used to interrogate or specify the fill of the slide
   * background.
   *
   * Note that accessing this property is potentially destructive. A slide background can also
   * be specified by a background style reference and accessing this property will remove that
   * style reference, if present, and replace it with NoFill. This is frequently the case for a
   * slide master background.
   *
   * This is because the `FillFormat` object requires a `p:bgPr` element and only one of a
   * `p:bgPr` or a `p:bgRef` element can be present in a `p:bg` element.
   */
  get fill(): FillFormat {
    return lazy(this, "fill", () => {
      const bgPr = this._cSld.getOrAddBgPr();
      return FillFormat.fromFillParent(bgPr);
    });
  }
}
