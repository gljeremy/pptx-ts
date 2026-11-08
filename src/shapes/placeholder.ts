// Placeholder-related objects (`pptx.shapes.placeholder`).
//
// Specific to shapes having a `p:ph` element. A placeholder has distinct behaviors depending on
// whether it appears on a slide, layout, or master. Hence there is a non-trivial class
// inheritance structure. python's `_InheritsDimensions` mixin is a class-factory mixin here,
// applied to `Shape` or `Picture`.
import type { XL_CHART_TYPE } from "../enum/chart.js";
import { MSO_SHAPE_TYPE, PP_PLACEHOLDER } from "../enum/shapes.js";
import { KeyError, NotImplementedError } from "../errors.js";
import type { FileSource } from "../fileio.js";
import type { CT_Shape } from "../oxml/shapes/autoshape.js";
import { CT_GraphicalObjectFrame } from "../oxml/shapes/graphfrm.js";
import { CT_Picture } from "../oxml/shapes/picture.js";
import type { ChartDataSource } from "../parts/chart.js";
import type { NotesSlidePart, SlideLayoutPart, SlidePart } from "../parts/slide.js";
import { Emu } from "../util.js";
import { Shape } from "./autoshape.js";
import { GraphicFrame } from "./graphfrm.js";
import { Picture } from "./picture.js";

type Dimension = "left" | "top" | "width" | "height";
type Ctor<T = object> = abstract new (...args: any[]) => T;

/**
 * Mixin providing inherited dimension behavior. Specifically, left, top, width, and height
 * report the value from the layout placeholder where they would have otherwise reported null.
 * This behavior is distinctive to placeholders. `_basePlaceholder` must be overridden by all
 * subclasses to provide lookup of the appropriate base placeholder to inherit from.
 */
export function _InheritsDimensions<TBase extends Ctor<{ _element: any }>>(Base: TBase) {
  abstract class InheritsDimensions extends Base {
    /**
     * The effective height of this placeholder shape; its directly-applied height if it has
     * one, otherwise the height of its parent layout placeholder.
     */
    get height(): any {
      return this._effectiveValue("height");
    }
    set height(value: number) {
      this._element.cy = value;
    }

    /**
     * The effective left of this placeholder shape; its directly-applied left if it has one,
     * otherwise the left of its parent layout placeholder.
     */
    get left(): any {
      return this._effectiveValue("left");
    }
    set left(value: number) {
      this._element.x = value;
    }

    /** Member of `MSO_SHAPE_TYPE` specifying the type of this shape; always PLACEHOLDER here. */
    get shapeType(): MSO_SHAPE_TYPE {
      return MSO_SHAPE_TYPE.PLACEHOLDER;
    }

    /**
     * The effective top of this placeholder shape; its directly-applied top if it has one,
     * otherwise the top of its parent layout placeholder.
     */
    get top(): any {
      return this._effectiveValue("top");
    }
    set top(value: number) {
      this._element.y = value;
    }

    /**
     * The effective width of this placeholder shape; its directly-applied width if it has one,
     * otherwise the width of its parent layout placeholder.
     */
    get width(): any {
      return this._effectiveValue("width");
    }
    set width(value: number) {
      this._element.cx = value;
    }

    /**
     * The layout or master placeholder shape this placeholder inherits from. Not to be confused
     * with an instance of `BasePlaceholder` (necessarily).
     */
    get _basePlaceholder(): any {
      throw new NotImplementedError("Must be implemented by all subclasses.");
    }

    /**
     * The effective value of `attrName` on this placeholder shape; its directly-applied value if
     * it has one, otherwise the value on the layout placeholder it inherits from.
     */
    _effectiveValue(attrName: Dimension): number | null {
      // python `getattr(super(_InheritsDimensions, self), attr_name)`
      const directlyAppliedValue = Reflect.get(Base.prototype, attrName, this) as number | null | undefined;
      if (directlyAppliedValue !== null && directlyAppliedValue !== undefined) return directlyAppliedValue;
      return this._inheritedValue(attrName);
    }

    /** The attribute value, e.g. 'width', of the base placeholder this placeholder inherits from. */
    _inheritedValue(attrName: Dimension): number | null {
      const basePlaceholder = this._basePlaceholder;
      if (basePlaceholder === null || basePlaceholder === undefined) return null;
      return basePlaceholder[attrName];
    }
  }
  return InheritsDimensions;
}

/**
 * Base class for placeholders on slides.
 *
 * Provides common behaviors such as inherited dimensions.
 */
export class _BaseSlidePlaceholder extends _InheritsDimensions(Shape) {
  declare _element: CT_Shape;

  get part(): SlidePart {
    return this._parent.part;
  }

  /** Whether this shape is a placeholder; unconditionally true in this case. */
  get isPlaceholder(): boolean {
    return true;
  }

  /**
   * The layout placeholder this slide placeholder inherits from. Not to be confused with an
   * instance of `BasePlaceholder` (necessarily).
   */
  get _basePlaceholder(): LayoutPlaceholder | null {
    const [layout, idx] = [this.part.slideLayout, this._element.phIdx];
    return layout.placeholders.get(idx);
  }

  /**
   * Substitute `element` for this placeholder element in the shapetree.
   *
   * This placeholder's `_element` attribute is set to null and its original element is free
   * for garbage collection. Any attribute access (including a method call) on this placeholder
   * after this call fails.
   */
  _replacePlaceholderWith(element: any): void {
    element._nvXxPr.nvPr._insertPh(this._element.ph);
    this._element.addprevious(element);
    this._element.getparent()!.remove(this._element);
    (this as any)._element = null;
  }
}

/**
 * NOTE: This class is deprecated and will be removed from a future release along with the
 * properties `idx`, `orient`, `phType`, and `sz`. The `idx` property will be available via the
 * `.placeholderFormat` property. The others will be accessed directly from the oxml layer as
 * they are only used for internal purposes.
 *
 * Base class for placeholder subclasses that differentiate the varying behaviors of
 * placeholders on a master, layout, and slide.
 */
export class BasePlaceholder extends Shape {
  /** Integer placeholder 'idx' attribute, e.g. 0. */
  get idx(): number {
    return this._sp.phIdx;
  }

  /** Placeholder orientation, e.g. ST_Direction.HORZ. */
  get orient(): string {
    return this._sp.phOrient;
  }

  /** Placeholder type, e.g. PP_PLACEHOLDER.CENTER_TITLE. */
  get phType(): PP_PLACEHOLDER {
    return this._sp.phType;
  }

  /** Placeholder 'sz' attribute, e.g. ST_PlaceholderSize.FULL. */
  get sz(): string {
    return this._sp.phSz;
  }
}

const LAYOUT_BASE_PH_TYPE = new Map<PP_PLACEHOLDER, PP_PLACEHOLDER>([
  [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.CHART, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.BITMAP, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE],
  [PP_PLACEHOLDER.ORG_CHART, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.DATE],
  [PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.FOOTER],
  [PP_PLACEHOLDER.MEDIA_CLIP, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.SLIDE_NUMBER],
  [PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.TABLE, PP_PLACEHOLDER.BODY],
  [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.TITLE],
]);

/**
 * Placeholder shape on a slide layout.
 *
 * Provides differentiated behavior for slide layout placeholders, in particular, inheriting
 * shape properties from the master placeholder having the same type, when a matching one
 * exists.
 */
export class LayoutPlaceholder extends _InheritsDimensions(Shape) {
  declare _element: CT_Shape;

  get element(): CT_Shape {
    return this._element;
  }

  get part(): SlideLayoutPart {
    return this._parent.part;
  }

  /** The master placeholder this layout placeholder inherits from. */
  get _basePlaceholder(): MasterPlaceholder | null {
    const phType = this._element.phType;
    const basePhType = LAYOUT_BASE_PH_TYPE.get(phType);
    if (basePhType === undefined) throw new KeyError(String(phType));
    const slideMaster = this.part.slideMaster;
    return slideMaster.placeholders.get(basePhType, null);
  }
}

/** Placeholder shape on a slide master. */
export class MasterPlaceholder extends BasePlaceholder {
  declare _element: CT_Shape;

  get element(): CT_Shape {
    return this._element;
  }
}

/**
 * Placeholder shape on a notes slide. Inherits shape properties from the placeholder on the
 * notes master that has the same type (e.g. 'body').
 */
export class NotesSlidePlaceholder extends _InheritsDimensions(Shape) {
  declare _element: CT_Shape;

  get part(): NotesSlidePart {
    return this._parent.part;
  }

  /**
   * The notes master placeholder this notes slide placeholder inherits from, or null if no
   * placeholder of the matching type is present.
   */
  get _basePlaceholder(): MasterPlaceholder | null {
    const notesMaster = this.part.notesMaster;
    const phType = this.element.phType;
    return notesMaster.placeholders.get(phType);
  }
}

/** Placeholder shape on a slide. Inherits shape properties from its slide layout placeholder. */
export class SlidePlaceholder extends _BaseSlidePlaceholder {}

/** Placeholder shape that can only accept a chart. */
export class ChartPlaceholder extends _BaseSlidePlaceholder {
  /**
   * A `PlaceholderGraphicFrame` object containing a new chart of `chartType` depicting
   * `chartData` and having the same position and size as this placeholder. `chartType` is one
   * of the `XL_CHART_TYPE` enumeration values. `chartData` is a `ChartData` object populated
   * with the categories and series values for the chart.
   *
   * Note that the new `Chart` object is not returned directly. The chart object may be accessed
   * using the `chart` property of the returned `PlaceholderGraphicFrame` object.
   */
  insertChart(chartType: XL_CHART_TYPE, chartData: ChartDataSource): PlaceholderGraphicFrame {
    const rId = this.part.addChartPart(chartType, chartData);
    const graphicFrame = this._newChartGraphicFrame(rId, this.left, this.top, this.width, this.height);
    this._replacePlaceholderWith(graphicFrame);
    return new PlaceholderGraphicFrame(graphicFrame, this._parent);
  }

  /**
   * A newly created `p:graphicFrame` element having the specified position and size and
   * containing the chart identified by `rId`.
   */
  _newChartGraphicFrame(rId: string, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const [id_, name] = [this.shapeId, this.name];
    return CT_GraphicalObjectFrame.newChartGraphicFrame(id_, name, rId, x, y, cx, cy);
  }
}

/** Placeholder shape that can only accept a picture. */
export class PicturePlaceholder extends _BaseSlidePlaceholder {
  /**
   * A `PlaceholderPicture` object depicting the image in `imageFile`.
   *
   * `imageFile` may be either a path (string) or a file-like object. The image is cropped to
   * fill the entire space of the placeholder. A `PlaceholderPicture` object has all the
   * properties and methods of a `Picture` shape except that the value of its `shapeType`
   * property is `MSO_SHAPE_TYPE.PLACEHOLDER` instead of `MSO_SHAPE_TYPE.PICTURE`.
   */
  insertPicture(imageFile: FileSource): PlaceholderPicture {
    const pic = this._newPlaceholderPic(imageFile);
    this._replacePlaceholderWith(pic);
    return new PlaceholderPicture(pic, this._parent);
  }

  /**
   * A new `p:pic` element depicting the image in `imageFile`, suitable for use as a
   * placeholder. In particular this means not having an `a:xfrm` element, allowing its extents
   * to be inherited from its layout placeholder.
   */
  _newPlaceholderPic(imageFile: FileSource): CT_Picture {
    const [rId, desc, imageSize] = this._getOrAddImage(imageFile);
    const [shapeId, name] = [this.shapeId, this.name];
    const pic = CT_Picture.newPhPic(shapeId, name, desc, rId);
    pic.cropToFit(imageSize, [this.width, this.height]);
    return pic;
  }

  /**
   * [rId, description, imageSize] identifying the related image part containing `imageFile`
   * and describing the image.
   */
  _getOrAddImage(imageFile: FileSource): [string, string, [number, number]] {
    const [imagePart, rId] = this.part.getOrAddImagePart(imageFile);
    const [desc, imageSize] = [imagePart.desc, imagePart._pxSize];
    return [rId, desc, imageSize];
  }
}

/** Placeholder shape populated with a table, chart, or smart art. */
export class PlaceholderGraphicFrame extends GraphicFrame {
  /** Whether this shape is a placeholder; unconditionally true in this case. */
  get isPlaceholder(): boolean {
    return true;
  }
}

/** Placeholder shape populated with a picture. */
export class PlaceholderPicture extends _InheritsDimensions(Picture) {
  declare _element: CT_Picture;

  get part(): SlidePart {
    return this._parent.part;
  }

  /** The layout placeholder this picture placeholder inherits from. */
  get _basePlaceholder(): LayoutPlaceholder | null {
    const [layout, idx] = [this.part.slideLayout, this._element.phIdx];
    return layout.placeholders.get(idx);
  }
}

/** Placeholder shape that can only accept a table. */
export class TablePlaceholder extends _BaseSlidePlaceholder {
  /**
   * A `PlaceholderGraphicFrame` object containing a `rows` by `cols` table.
   *
   * The position and width of the table are those of the placeholder and its height is
   * proportional to the number of rows. Note that the return value is not the new table but
   * rather *contains* the new table. The table can be accessed using the `table` property of
   * the returned `PlaceholderGraphicFrame` object.
   */
  insertTable(rows: number, cols: number): PlaceholderGraphicFrame {
    const graphicFrame = this._newPlaceholderTable(rows, cols);
    this._replacePlaceholderWith(graphicFrame);
    return new PlaceholderGraphicFrame(graphicFrame, this._parent);
  }

  /**
   * A newly added `p:graphicFrame` element containing an empty table with `rows` rows and
   * `cols` columns, positioned at the location of this placeholder and having its same width.
   * The table's height is determined by the number of rows.
   */
  _newPlaceholderTable(rows: number, cols: number): CT_GraphicalObjectFrame {
    const [shapeId, name, height] = [this.shapeId, this.name, Emu(rows * 370840)];
    return CT_GraphicalObjectFrame.newTableGraphicFrame(shapeId, name, rows, cols, this.left, this.top, this.width, height);
  }
}

