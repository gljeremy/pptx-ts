// GroupShape and related objects. (`pptx.shapes.group`)
import { lazy } from "../util.js";
import { ShadowFormat } from "../dml/effect.js";
import { MSO_SHAPE_TYPE } from "../enum/shapes.js";
import { BaseShape } from "./base.js";
import type { ActionSetting } from "../action.js";
import type { CT_GroupShape } from "../oxml/shapes/groupshape.js";
import { GroupShapes } from "./shapetree.js";
import type { ProvidesPart } from "../types.js";

/** A shape that acts as a container for other shapes. */
export class GroupShape extends BaseShape {
  declare _grpSp: any;

  constructor(grpSp: CT_GroupShape, parent: ProvidesPart) {
    super(grpSp, parent);
    this._grpSp = grpSp;
  }

  /**
   * Unconditionally raises `TypeError`.
   *
   * A group shape cannot have a click action or hover action.
   */
  get clickAction(): ActionSetting {
    return lazy(this, "clickAction", () => {
      throw new TypeError("a group shape cannot have a click action");
    });
  }

  /**
   * Unconditionally |False|.
   *
   * A group shape does not have a textframe and cannot itself contain text. This does not
   * impact the ability of shapes contained by the group to each have their own text.
   */
  get hasTextFrame(): boolean {
    return false;
  }

  /**
   * |ShadowFormat| object representing shadow effect for this group.
   *
   * A |ShadowFormat| object is always returned, even when no shadow is explicitly defined on
   * this group shape (i.e. when the group inherits its shadow behavior).
   */
  get shadow(): ShadowFormat {
    return lazy(this, "shadow", () => {
      return new ShadowFormat(this._grpSp.grpSpPr);
    });
  }

  /**
   * Member of :ref:`MsoShapeType` identifying the type of this shape.
   *
   * Unconditionally `MSO_SHAPE_TYPE.GROUP` in this case
   */
  get shapeType(): MSO_SHAPE_TYPE {
    return MSO_SHAPE_TYPE.GROUP;
  }

  /**
   * |GroupShapes| object for this group.
   *
   * The |GroupShapes| object provides access to the group's member shapes and provides methods
   * for adding new ones.
   */
  get shapes(): GroupShapes {
    return lazy(this, "shapes", () => {
      return new GroupShapes(this._element, this);
    });
  }
}
