// Objects related to mouse click and hover actions on a shape or text. (`pptx.action`)
import { KeyError, ValueError } from "./errors.js";
import { lazy } from "./util.js";
import { PP_ACTION } from "./enum/action.js";
import { RELATIONSHIP_TYPE as RT } from "./opc/constants.js";
import { Subshape } from "./shapes/index.js";
import type { CT_Hyperlink } from "./oxml/action.js";
import { CT_NonVisualDrawingProps } from "./oxml/shapes/shared.js";
import type { CT_TextCharacterProperties } from "./oxml/text.js";
import type { SlidePart } from "./parts/slide.js";
import type { BaseShape } from "./shapes/base.js";
import type { Slide, Slides } from "./slide.js";

/** Properties specifying how a shape or run reacts to mouse actions. */
export class ActionSetting extends Subshape {
  declare _element: any;
  declare _hover: any;

  constructor(xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties, parent: BaseShape, hover: boolean = false) {
    super(parent);
    this._element = xPr;
    this._hover = hover;
  }

  /**
   * Member of :ref:`PpActionType` enumeration, such as `PP_ACTION.HYPERLINK`.
   *
   * The returned member indicates the type of action that will result when the
   * specified shape or text is clicked or the mouse pointer is positioned over the
   * shape during a slide show.
   *
   * If there is no click-action or the click-action value is not recognized (is not
   * one of the official `MsoPpAction` values) then `PP_ACTION.NONE` is returned.
   */
  get action() {
    const hlink = this._hlink;
    if (hlink === null) {
      return PP_ACTION.NONE;
    }
    const actionVerb = hlink.actionVerb;
    if (actionVerb === "hlinkshowjump") {
      const relativeTarget = hlink.actionFields["jump"];
      const byTarget: Record<string, PP_ACTION> = {
        firstslide: PP_ACTION.FIRST_SLIDE,
        lastslide: PP_ACTION.LAST_SLIDE,
        lastslideviewed: PP_ACTION.LAST_SLIDE_VIEWED,
        nextslide: PP_ACTION.NEXT_SLIDE,
        previousslide: PP_ACTION.PREVIOUS_SLIDE,
        endshow: PP_ACTION.END_SHOW,
      };
      if (!(relativeTarget in byTarget)) throw new KeyError(relativeTarget);
      return byTarget[relativeTarget];
    }
    if (actionVerb === null) return PP_ACTION.HYPERLINK;
    const byVerb: Record<string, PP_ACTION> = {
      hlinksldjump: PP_ACTION.NAMED_SLIDE,
      hlinkpres: PP_ACTION.PLAY,
      hlinkfile: PP_ACTION.OPEN_FILE,
      customshow: PP_ACTION.NAMED_SLIDE_SHOW,
      ole: PP_ACTION.OLE_VERB,
      macro: PP_ACTION.RUN_MACRO,
      program: PP_ACTION.RUN_PROGRAM,
    };
    return Object.prototype.hasOwnProperty.call(byVerb, actionVerb) ? byVerb[actionVerb] : PP_ACTION.NONE;
  }

  /**
   * A |Hyperlink| object representing the hyperlink action defined on
   * this click or hover mouse event. A |Hyperlink| object is always
   * returned, even if no hyperlink or other click action is defined.
   */
  get hyperlink(): Hyperlink {
    return lazy(this, "hyperlink", () => {
      return new Hyperlink(this._element, this._parent, this._hover);
    });
  }

  /**
   * A reference to the slide in this presentation that is the target of
   * the slide jump action in this shape. Slide jump actions include
   * `PP_ACTION.FIRST_SLIDE`, `LAST_SLIDE`, `NEXT_SLIDE`,
   * `PREVIOUS_SLIDE`, and `NAMED_SLIDE`. Returns |None| for all other
   * actions. In particular, the `LAST_SLIDE_VIEWED` action and the `PLAY`
   * (start other presentation) actions are not supported.
   *
   * A slide object may be assigned to this property, which makes the
   * shape an "internal hyperlink" to the assigened slide::
   *
   *     slide, target_slide = prs.slides[0], prs.slides[1]
   *     shape = slide.shapes[0]
   *     shape.target_slide = target_slide
   *
   * Assigning |None| removes any slide jump action. Note that this is
   * accomplished by removing any action present (such as a hyperlink),
   * without first checking that it is a slide jump action.
   */
  get targetSlide(): Slide | null {
    const slideJumpActions = [PP_ACTION.FIRST_SLIDE, PP_ACTION.LAST_SLIDE, PP_ACTION.NEXT_SLIDE, PP_ACTION.PREVIOUS_SLIDE, PP_ACTION.NAMED_SLIDE];
    if (!slideJumpActions.includes(this.action)) {
      return null;
    }
    if (this.action === PP_ACTION.FIRST_SLIDE) {
      return this._slides.at(0);
    } else if (this.action === PP_ACTION.LAST_SLIDE) {
      return this._slides.at(-1);
    } else if (this.action === PP_ACTION.NEXT_SLIDE) {
      const nextSlideIdx = (this._slideIndex + 1);
      if (nextSlideIdx >= this._slides.length) {
        throw new ValueError("no next slide");
      }
      return this._slides.at(nextSlideIdx);
    } else if (this.action === PP_ACTION.PREVIOUS_SLIDE) {
      const prevSlideIdx = (this._slideIndex - 1);
      if (prevSlideIdx < 0) {
        throw new ValueError("no previous slide");
      }
      return this._slides.at(prevSlideIdx);
    } else if (this.action === PP_ACTION.NAMED_SLIDE) {
      if (!(this._hlink !== null)) throw new Error("AssertionError");
      const rId = this._hlink.rId;
      const slidePart = this.part.relatedPart(rId) as SlidePart;
      return slidePart.slide;
    }
    return null;
  }

  set targetSlide(slide: Slide | null) {
    this._clearClickAction();
    if (slide === null) {
      return;
    }
    const hlink = this._element.getOrAddHlinkClick();
    hlink.action = "ppaction://hlinksldjump";
    hlink.rId = this.part.relateTo(slide.part, RT.SLIDE);
  }

  /** Remove any existing click action. */
  _clearClickAction() {
    const hlink = this._hlink;
    if (hlink === null) {
      return;
    }
    const rId = hlink.rId;
    if (rId) {
      this.part.dropRel(rId);
    }
    this._element.remove(hlink);
  }

  /**
   * Reference to the `a:hlinkClick` or `a:hlinkHover` element for this
   * click action. Returns |None| if the element is not present.
   */
  get _hlink(): CT_Hyperlink | null {
    if (this._hover) {
      if (!(this._element instanceof CT_NonVisualDrawingProps)) throw new Error("AssertionError");
      return this._element.hlinkHover;
    }
    return this._element.hlinkClick;
  }

  /** Reference to the slide containing the shape having this click action. */
  get _slide(): Slide {
    return lazy(this, "_slide", () => {
      return (this.part as SlidePart).slide;
    });
  }

  /**
   * Position in the slide collection of the slide containing the shape
   * having this click action.
   */
  get _slideIndex() {
    return lazy(this, "_slideIndex", () => {
      return this._slides.index(this._slide);
    });
  }

  /** Reference to the slide collection for this presentation. */
  get _slides(): Slides {
    return lazy(this, "_slides", () => {
      return this.part.package.presentationPart.presentation.slides;
    });
  }
}

/** Represents a hyperlink action on a shape or text run. */
export class Hyperlink extends Subshape {
  declare _element: any;
  declare _hover: any;

  constructor(xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties, parent: BaseShape, hover: boolean = false) {
    super(parent);
    this._element = xPr;
    this._hover = hover;
  }

  /**
   * Read/write. The URL of the hyperlink.
   *
   * URL can be on http, https, mailto, or file scheme; others may work. Returns |None| if no
   * hyperlink is defined, including when another action such as `RUN_MACRO` is defined on the
   * object. Assigning |None| removes any action defined on the object, whether it is a hyperlink
   * action or not.
   */
  get address(): string | null {
    const hlink = this._hlink;
    if (hlink === null) {
      return null;
    }
    const rId = hlink.rId;
    if (!rId) {
      return null;
    }
    return this.part.targetRef(rId);
  }

  set address(url: string | null) {
    this._removeHlink();
    if (url) {
      const rId = this.part.relateTo(url, RT.HYPERLINK, true);
      const hlink = this._getOrAddHlink();
      hlink.rId = rId;
    }
  }

  /**
   * Get the `a:hlinkClick` or `a:hlinkHover` element for the Hyperlink object.
   *
   * The actual element depends on the value of `self._hover`. Create the element if not present.
   */
  _getOrAddHlink(): CT_Hyperlink {
    if (this._hover) {
      return this._element.getOrAddHlinkHover();
    }
    return this._element.getOrAddHlinkClick();
  }

  /**
   * Reference to the `a:hlinkClick` or `h:hlinkHover` element for this click action.
   *
   * Returns |None| if the element is not present.
   */
  get _hlink(): CT_Hyperlink | null {
    if (this._hover) {
      return this._element.hlinkHover;
    }
    return this._element.hlinkClick;
  }

  /**
   * Remove the a:hlinkClick or a:hlinkHover element.
   *
   * Also drops any relationship it might have.
   */
  _removeHlink() {
    const hlink = this._hlink;
    if (hlink === null) {
      return;
    }
    const rId = hlink.rId;
    if (rId) {
      this.part.dropRel(rId);
    }
    this._element.remove(hlink);
  }
}
