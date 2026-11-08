// Slide and related objects (`pptx.parts.slide`).
import type { XL_CHART_TYPE } from "../enum/chart.js";
import { PROG_ID } from "../enum/shapes.js";
import { KeyError } from "../errors.js";
import type { FileSource } from "../fileio.js";
import type { Video } from "../media.js";
import { CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT } from "../opc/constants.js";
import { XmlPart } from "../opc/package.js";
import { PackURI } from "../opc/packuri.js";
import { CT_NotesMaster, CT_NotesSlide, CT_Slide, type CT_SlideLayout, type CT_SlideMaster } from "../oxml/slide.js";
import { CT_OfficeStyleSheet } from "../oxml/theme.js";
import type { Package } from "../package.js";
import { NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster } from "../slide.js";
import { lazy } from "../util.js";
import { ChartPart, type ChartDataSource } from "./chart.js";
import { EmbeddedPackagePart } from "./embeddedpackage.js";
import type { Image, ImagePart } from "./image.js";
import type { MediaPart } from "./media.js";

/** Base class for slide parts, e.g. slide, slide-layout, slide-master, notes-slide, etc. */
export class BaseSlidePart extends XmlPart {
  declare _element: CT_Slide | CT_SlideLayout | CT_SlideMaster | CT_NotesSlide | CT_NotesMaster;
  declare _package: Package;

  get package(): Package {
    return this._package;
  }

  /**
   * An `Image` object containing the image related to this slide by key `rId`.
   *
   * Raises KeyError if no image is related by that id, which would generally indicate a
   * corrupted .pptx file.
   */
  getImage(rId: string): Image {
    return (this.relatedPart(rId) as ImagePart).image;
  }

  /**
   * [imagePart, rId] for `imageFile`.
   *
   * The returned `ImagePart` contains the image in `imageFile` and is related to this slide
   * with the key `rId`. If either the image part or relationship already exists, they are
   * reused, otherwise they are newly created.
   */
  getOrAddImagePart(imageFile: FileSource): [ImagePart, string] {
    const imagePart = this._package.getOrAddImagePart(imageFile);
    const rId = this.relateTo(imagePart, RT.IMAGE);
    return [imagePart, rId];
  }

  /** Internal name of this slide. */
  get name(): string {
    return this._element.cSld.name;
  }
}

/**
 * Notes master part.
 *
 * Corresponds to package file `ppt/notesMasters/notesMaster1.xml`.
 */
export class NotesMasterPart extends BaseSlidePart {
  declare _element: CT_NotesMaster;

  /**
   * A new notes master part with default content.
   *
   * Its theme part is also newly created and related; both are added to `package_`.
   */
  static createDefault(package_: Package): NotesMasterPart {
    const notesMasterPart = this._new(package_);
    const themePart = this._newThemePart(package_);
    notesMasterPart.relateTo(themePart, RT.THEME);
    return notesMasterPart;
  }

  /** The `NotesMaster` object that proxies this notes master part. */
  get notesMaster(): NotesMaster {
    return lazy(this, "notesMaster", () => new NotesMaster(this._element, this));
  }

  /** A new `NotesMasterPart` with default content, to be added to `package_`. */
  static _new(package_: Package): NotesMasterPart {
    return new NotesMasterPart(new PackURI("/ppt/notesMasters/notesMaster1.xml"), CT.PML_NOTES_MASTER, package_, CT_NotesMaster.newDefault());
  }

  /** A new default theme-part suitable for use with a notes master. */
  static _newThemePart(package_: Package): XmlPart {
    return new XmlPart(package_.nextPartname("/ppt/theme/theme%d.xml"), CT.OFC_THEME, package_, CT_OfficeStyleSheet.newDefault());
  }
}

/**
 * Notes slide part.
 *
 * Contains the slide notes content and the layout for the slide handout page. Corresponds to
 * package file `ppt/notesSlides/notesSlide[1-9][0-9]*.xml`.
 */
export class NotesSlidePart extends BaseSlidePart {
  declare _element: CT_NotesSlide;

  /**
   * A new `NotesSlidePart` for the slide in `slidePart`.
   *
   * The new notes-slide part is based on the (singleton) notes master and related to both
   * the notes-master part and `slidePart`. If no notes-master is present, one is created
   * based on the default template.
   */
  static new(package_: Package, slidePart: SlidePart): NotesSlidePart {
    const notesMasterPart = package_.presentationPart.notesMasterPart;
    const notesSlidePart = this._addNotesSlidePart(package_, slidePart, notesMasterPart);
    const notesSlide = notesSlidePart.notesSlide;
    notesSlide.cloneMasterPlaceholders(notesMasterPart.notesMaster);
    return notesSlidePart;
  }

  /** The `NotesMaster` object this notes slide inherits from. */
  get notesMaster(): NotesMaster {
    return lazy(this, "notesMaster", () => (this.partRelatedBy(RT.NOTES_MASTER) as NotesMasterPart).notesMaster);
  }

  /** The `NotesSlide` object that proxies this notes slide part. */
  get notesSlide(): NotesSlide {
    return lazy(this, "notesSlide", () => new NotesSlide(this._element, this));
  }

  /** A newly-created notes slide part, related to `notesMasterPart` and `slidePart`. */
  static _addNotesSlidePart(package_: Package, slidePart: SlidePart, notesMasterPart: NotesMasterPart): NotesSlidePart {
    const notesSlidePart = new NotesSlidePart(package_.nextPartname("/ppt/notesSlides/notesSlide%d.xml"), CT.PML_NOTES_SLIDE, package_, CT_NotesSlide.new());
    notesSlidePart.relateTo(notesMasterPart, RT.NOTES_MASTER);
    notesSlidePart.relateTo(slidePart, RT.SLIDE);
    return notesSlidePart;
  }
}

/** Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml. */
export class SlidePart extends BaseSlidePart {
  declare _element: CT_Slide;

  /**
   * A newly-created blank slide part.
   *
   * The new slide-part has `partname` and a relationship to `slideLayoutPart`.
   */
  static new(partname: PackURI, package_: Package, slideLayoutPart: SlideLayoutPart): SlidePart {
    const slidePart = new this(partname, CT.PML_SLIDE, package_, CT_Slide.new());
    slidePart.relateTo(slideLayoutPart, RT.SLIDE_LAYOUT);
    return slidePart;
  }

  /**
   * The rId of a new `ChartPart` object containing a chart of `chartType`.
   *
   * The chart depicts `chartData` and is related to the slide contained in this part by `rId`.
   */
  addChartPart(chartType: XL_CHART_TYPE, chartData: ChartDataSource): string {
    return this.relateTo(ChartPart.new(chartType, chartData, this._package), RT.CHART);
  }

  /**
   * The rId of a newly-added OLE-object part formed from `oleObjectFile`.
   *
   * A member of `PROG_ID` relates the part as an embedded package; any other progId string
   * relates it as an OLE object.
   */
  addEmbeddedOleObjectPart(progId: PROG_ID | string, oleObjectFile: FileSource): string {
    const relationshipType = progId instanceof PROG_ID ? RT.PACKAGE : RT.OLE_OBJECT;
    return this.relateTo(EmbeddedPackagePart.factory(progId, this._blobFromFile(oleObjectFile), this._package), relationshipType);
  }

  /**
   * [mediaRId, videoRId] pair for a `MediaPart` containing `video`.
   *
   * The video media part is reused if it already exists in the package; the two relationships
   * from this slide to it are reused when present.
   */
  getOrAddVideoMediaPart(video: Video): [string, string] {
    const mediaPart: MediaPart = this._package.getOrAddMediaPart(video);
    const mediaRId = this.relateTo(mediaPart, RT.MEDIA);
    const videoRId = this.relateTo(mediaPart, RT.VIDEO);
    return [mediaRId, videoRId];
  }

  /**
   * True if this slide has a notes slide, false otherwise. Reading `notesSlide` creates one
   * when none is present.
   */
  get hasNotesSlide(): boolean {
    try {
      this.partRelatedBy(RT.NOTES_SLIDE);
    } catch (e) {
      if (!(e instanceof KeyError)) throw e;
      return false;
    }
    return true;
  }

  /**
   * The `NotesSlide` instance associated with this slide.
   *
   * If the slide does not have a notes slide, a new one is created. The same single instance
   * is returned on each call.
   */
  get notesSlide(): NotesSlide {
    return lazy(this, "notesSlide", () => {
      let notesSlidePart: NotesSlidePart;
      try {
        notesSlidePart = this.partRelatedBy(RT.NOTES_SLIDE);
      } catch (e) {
        if (!(e instanceof KeyError)) throw e;
        notesSlidePart = this._addNotesSlidePart();
      }
      return notesSlidePart.notesSlide;
    });
  }

  /** The `Slide` object representing this slide part. */
  get slide(): Slide {
    return lazy(this, "slide", () => new Slide(this._element, this));
  }

  /**
   * The slide identifier stored in the presentation part for this slide part, unique within
   * the presentation and unchanged when the slide is moved.
   */
  get slideId(): number {
    return this.package.presentationPart.slideId(this);
  }

  /** `SlideLayout` object the slide in this part inherits appearance from. */
  get slideLayout(): SlideLayout {
    return (this.partRelatedBy(RT.SLIDE_LAYOUT) as SlideLayoutPart).slideLayout;
  }

  /** A newly created `NotesSlidePart` object related to this slide part. */
  _addNotesSlidePart(): NotesSlidePart {
    const notesSlidePart = NotesSlidePart.new(this.package, this);
    this.relateTo(notesSlidePart, RT.NOTES_SLIDE);
    return notesSlidePart;
  }
}

/** Slide layout part. Corresponds to package files ppt/slideLayouts/slideLayout[1-9][0-9]*.xml. */
export class SlideLayoutPart extends BaseSlidePart {
  declare _element: CT_SlideLayout;

  /** The `SlideLayout` object representing this part. */
  get slideLayout(): SlideLayout {
    return lazy(this, "slideLayout", () => new SlideLayout(this._element, this));
  }

  /** Slide master from which this slide layout inherits properties. */
  get slideMaster(): SlideMaster {
    return (this.partRelatedBy(RT.SLIDE_MASTER) as SlideMasterPart).slideMaster;
  }
}

/** Slide master part. Corresponds to package files ppt/slideMasters/slideMaster[1-9][0-9]*.xml. */
export class SlideMasterPart extends BaseSlidePart {
  declare _element: CT_SlideMaster;

  /** `SlideLayout` related to this slide-master by key `rId`. */
  relatedSlideLayout(rId: string): SlideLayout {
    return (this.relatedPart(rId) as SlideLayoutPart).slideLayout;
  }

  /** The `SlideMaster` object representing this part. */
  get slideMaster(): SlideMaster {
    return lazy(this, "slideMaster", () => new SlideMaster(this._element, this));
  }
}
