// The shape tree, the structure that holds a slide's shapes. (`pptx.shapes.shapetree`)
import { IndexError, KeyError, ValueError } from "../errors.js";
import { BytesIO, type FileSource } from "../fileio.js";
import { PY_GETITEM } from "../oxml/xml.js";
import { templateBytes } from "../templates/index.js";
import { lazy, pyBool } from "../util.js";
import { PP_PLACEHOLDER, PROG_ID } from "../enum/shapes.js";
import { SPEAKER_IMAGE_BYTES, Video } from "../media.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { qn } from "../oxml/ns.js";
import { CT_Shape } from "../oxml/shapes/autoshape.js";
import { CT_GraphicalObjectFrame } from "../oxml/shapes/graphfrm.js";
import { CT_Picture } from "../oxml/shapes/picture.js";
import { ST_Direction } from "../oxml/simpletypes.js";
import { AutoShapeType, Shape } from "./autoshape.js";
import { BaseShape } from "./base.js";
import { Connector } from "./connector.js";
import { FreeformBuilder } from "./freeform.js";
import { GraphicFrame } from "./graphfrm.js";
import { GroupShape } from "./group.js";
import { Movie, Picture } from "./picture.js";
import { ChartPlaceholder, LayoutPlaceholder, MasterPlaceholder, NotesSlidePlaceholder, PicturePlaceholder, PlaceholderGraphicFrame, PlaceholderPicture, SlidePlaceholder, TablePlaceholder } from "./placeholder.js";
import { ParentedElementProxy } from "../shared.js";
import { Emu } from "../util.js";
import type { ChartDataSource } from "../parts/chart.js";
import type { XL_CHART_TYPE } from "../enum/chart.js";
import type { MSO_CONNECTOR_TYPE, MSO_SHAPE } from "../enum/shapes.js";
import type { ShapeElement } from "../oxml/shapes/index.js";
import type { CT_Connector } from "../oxml/shapes/connector.js";
import type { CT_GroupShape } from "../oxml/shapes/groupshape.js";
import type { ImagePart } from "../parts/image.js";
import type { SlidePart } from "../parts/slide.js";
import type { Slide, SlideLayout } from "../slide.js";
import type { CT_Slide } from "../oxml/slide.js";
import type { ProvidesPart } from "../types.js";

/**
 * Base class for a shape collection appearing in a slide-type object.
 *
 * Subclasses include Slide, SlideLayout, and SlideMaster. Provides common methods.
 */
export class _BaseShapes extends ParentedElementProxy {
  declare _spTree: CT_GroupShape;
  declare _cachedMaxShapeId: number | null;

  constructor(spTree: CT_GroupShape, parent: ProvidesPart) {
    super(spTree, parent);
    this._spTree = spTree;
    this._cachedMaxShapeId = null;
  }

  /** Return shape at `idx` in sequence, e.g. `shapes[2]`. */
  at(idx: number): BaseShape {
    const shapeElm = [...this._iterMemberElms()].at(idx);
    if (shapeElm === undefined) throw new IndexError("shape index out of range");
    return this._shapeFactory(shapeElm);
  }

  [PY_GETITEM](idx: number): BaseShape {
    return this.at(idx);
  }

  /** Generate a reference to each shape in the collection, in sequence. */
  *[Symbol.iterator](): Generator<BaseShape> {
    for (const shapeElm of this._iterMemberElms()) {
      yield this._shapeFactory(shapeElm);
    }
  }

  /**
   * Return count of shapes in this shape tree.
   *
   * A group shape contributes 1 to the total, without regard to the number of shapes contained
   * in the group.
   */
  get length(): number {
    const shapeElms = [...this._iterMemberElms()];
    return shapeElms.length;
  }

  /** Add a new placeholder shape based on `placeholder`. */
  clonePlaceholder(placeholder: LayoutPlaceholder | MasterPlaceholder): void {
    const sp = placeholder.element;
    const [phType, orient, sz, idx] = [sp.phType, sp.phOrient, sp.phSz, sp.phIdx];
    const id_ = this._nextShapeId;
    const name = this._nextPhName(phType, id_, orient);
    this._spTree.addPlaceholder(id_, name, phType, orient, sz, idx);
  }

  /**
   * Return the base name for a placeholder of `ph_type` in this shape collection.
   *
   * There is some variance between slide types, for example a notes slide uses a different
   * name for the body placeholder, so this method can be overriden by subclasses.
   */
  phBasename(phType: PP_PLACEHOLDER): string {
    return lookup(PH_BASENAMES, phType);
  }

  /**
   * True if "turbo-add" mode is enabled. Read/Write.
   *
   * EXPERIMENTAL: This feature can radically improve performance when adding large numbers
   * (hundreds of shapes) to a slide. It works by caching the last shape ID used and
   * incrementing that value to assign the next shape id. This avoids repeatedly searching all
   * shape ids in the slide each time a new ID is required.
   *
   * Performance is not noticeably improved for a slide with a relatively small number of
   * shapes, but because the search time rises with the square of the shape count, this option
   * can be useful for optimizing generation of a slide composed of many shapes.
   *
   * Shape-id collisions can occur (causing a repair error on load) if more than one |Slide|
   * object is used to interact with the same slide in the presentation. Note that the |Slides|
   * collection creates a new |Slide| object each time a slide is accessed (e.g. `slide =
   * prs.slides[0]`, so you must be careful to limit use to a single |Slide| object.
   */
  get turboAddEnabled(): boolean {
    return this._cachedMaxShapeId !== null;
  }

  set turboAddEnabled(value: boolean) {
    const enable = pyBool(value);
    this._cachedMaxShapeId = (enable ? this._spTree.maxShapeId : null);
  }

  /** Return true if `shape_elm` represents a member of this collection, False otherwise. */
  static _isMemberElm(_shapeElm: ShapeElement): boolean {
    return true;
  }

  /**
   * Generate each child of the `p:spTree` element that corresponds to a shape.
   *
   * Items appear in XML document order.
   */
  *_iterMemberElms(): Generator<ShapeElement> {
    for (const shapeElm of this._spTree.iterShapeElms()) {
      if ((this.constructor as typeof _BaseShapes)._isMemberElm(shapeElm)) {
        yield shapeElm;
      }
    }
  }

  /**
   * Next unique placeholder name for placeholder shape of type `ph_type`.
   *
   * Usually will be standard placeholder root name suffixed with id-1, e.g.
   * _next_ph_name(ST_PlaceholderType.TBL, 4, 'horz') ==> 'Table Placeholder 3'. The number is
   * incremented as necessary to make the name unique within the collection. If `orient` is
   * `'vert'`, the placeholder name is prefixed with `'Vertical '`.
   */
  _nextPhName(phType: PP_PLACEHOLDER, id: number, orient: string): string {
    let basename = this.phBasename(phType);
    if (orient === ST_Direction.VERT) {
      basename = `Vertical ${basename}`;
    }
    // increment numpart as necessary to make name unique
    let numpart = id - 1;
    const names: string[] = this._spTree.xpath("//p:cNvPr/@name");
    let name = `${basename} ${numpart}`;
    while (names.includes(name)) {
      numpart += 1;
      name = `${basename} ${numpart}`;
    }
    return name;
  }

  /**
   * Return a unique shape id suitable for use with a new shape.
   *
   * The returned id is 1 greater than the maximum shape id used so far. In practice, the
   * minimum id is 2 because the spTree element is always assigned id="1".
   */
  get _nextShapeId(): number {
    // presence of cached-max-shape-id indicates turbo mode is on
    if (this._cachedMaxShapeId !== null) {
      this._cachedMaxShapeId += 1;
      return this._cachedMaxShapeId;
    }
    return this._spTree.maxShapeId + 1;
  }

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return BaseShapeFactory(shapeElm, this);
  }
}

/** Base class for shape-trees that can add shapes. */
export class _BaseGroupShapes extends _BaseShapes {
  declare _element: CT_GroupShape;
  declare _grpSp: CT_GroupShape;

  get part(): SlidePart {
    return this._parent.part;
  }

  constructor(grpSp: CT_GroupShape, parent: ProvidesPart) {
    super(grpSp, parent);
    this._grpSp = grpSp;
  }

  /**
   * Add a new chart of `chart_type` to the slide.
   *
   * The chart is positioned at (`x`, `y`), has size (`cx`, `cy`), and depicts `chart_data`.
   * `chart_type` is one of the :ref:`XlChartType` enumeration values. `chart_data` is a
   * |ChartData| object populated with the categories and series values for the chart.
   *
   * Note that a |GraphicFrame| shape object is returned, not the |Chart| object contained in
   * that graphic frame shape. The chart object may be accessed using the :attr:`chart`
   * property of the returned |GraphicFrame| object.
   */
  addChart(chartType: XL_CHART_TYPE, x: number, y: number, cx: number, cy: number, chartData: ChartDataSource): GraphicFrame {
    const rId = this.part.addChartPart(chartType, chartData);
    const graphicFrame = this._addChartGraphicFrame(rId, x, y, cx, cy);
    this._recalculateExtents();
    return this._shapeFactory(graphicFrame) as GraphicFrame;
  }

  /**
   * Add a newly created connector shape to the end of this shape tree.
   *
   * `connector_type` is a member of the :ref:`MsoConnectorType` enumeration and the end-point
   * values are specified as EMU values. The returned connector is of type `connector_type` and
   * has begin and end points as specified.
   */
  addConnector(connectorType: MSO_CONNECTOR_TYPE, beginX: number, beginY: number, endX: number, endY: number): Connector {
    const cxnSp = this._addCxnSp(connectorType, beginX, beginY, endX, endY);
    this._recalculateExtents();
    return this._shapeFactory(cxnSp) as Connector;
  }

  /**
   * Return a |GroupShape| object newly appended to this shape tree.
   *
   * The group shape is empty and must be populated with shapes using methods on its shape
   * tree, available on its `.shapes` property. The position and extents of the group shape are
   * determined by the shapes it contains; its position and extents are recalculated each time
   * a shape is added to it.
   */
  addGroupShape(shapes: Iterable<BaseShape> = []): GroupShape {
    const shapeList = [...shapes];
    const grpSp = this._element.addGrpSp();
    for (const shape of shapeList) {
      grpSp.insertElementBefore(shape._element, "p:extLst");
    }
    if (shapeList.length > 0) {
      grpSp.recalculateExtents();
    }
    return this._shapeFactory(grpSp) as GroupShape;
  }

  /**
   * Return newly-created GraphicFrame shape embedding `object_file`.
   *
   * The returned graphic-frame shape contains `object_file` as an embedded OLE object. It is
   * displayed as an icon at `left`, `top` with size `width`, `height`. `width` and `height`
   * may be omitted when `prog_id` is a member of `PROG_ID`, in which case the default icon
   * size is used. This is advised for best appearance where applicable because it avoids an
   * icon with a "stretched" appearance.
   *
   * `object_file` may either be a str path to a file or file-like object (such as
   * `io.BytesIO`) containing the bytes of the object to be embedded (such as an Excel file).
   *
   * `prog_id` can be either a member of `pptx.enum.shapes.PROG_ID` or a str value like
   * `"Adobe.Exchange.7"` determined by inspecting the XML generated by PowerPoint for an
   * object of the desired type.
   *
   * `icon_file` may either be a str path to an image file or a file-like object containing the
   * image. The image provided will be displayed in lieu of the OLE object; double-clicking on
   * the image opens the object (subject to operating-system limitations). The image file can
   * be any supported image file. Those produced by PowerPoint itself are generally EMF and can
   * be harvested from a PPTX package that embeds such an object. PNG and JPG also work fine.
   *
   * `icon_width` and `icon_height` are `Length` values (e.g. Emu() or Inches()) that describe
   * the size of the icon image within the shape. These should be omitted unless a custom
   * `icon_file` is provided. The dimensions must be discovered by inspecting the XML.
   * Automatic resizing of the OLE-object shape can occur when the icon is double-clicked if
   * these values are not as set by PowerPoint. This behavior may only manifest in the Windows
   * version of PowerPoint.
   */
  addOleObject(
    objectFile: FileSource,
    progId: PROG_ID | string,
    left: number,
    top: number,
    width: number | null = null,
    height: number | null = null,
    iconFile: FileSource | null = null,
    iconWidth: number | null = null,
    iconHeight: number | null = null,
  ): GraphicFrame {
    const graphicFrame = _OleObjectElementCreator.graphicFrame(this, this._nextShapeId, objectFile, progId, left, top, width, height, iconFile, iconWidth, iconHeight);
    this._spTree.append(graphicFrame);
    this._recalculateExtents();
    return this._shapeFactory(graphicFrame) as GraphicFrame;
  }

  /**
   * Add picture shape displaying image in `image_file`.
   *
   * `image_file` can be either a path to a file (a string) or a file-like object. The picture
   * is positioned with its top-left corner at (`top`, `left`). If `width` and `height` are
   * both |None|, the native size of the image is used. If only one of `width` or `height` is
   * used, the unspecified dimension is calculated to preserve the aspect ratio of the image.
   * If both are specified, the picture is stretched to fit, without regard to its native
   * aspect ratio.
   */
  addPicture(imageFile: FileSource, left: number, top: number, width: number | null = null, height: number | null = null): Picture {
    const [imagePart, rId] = this.part.getOrAddImagePart(imageFile);
    const pic = this._addPicFromImagePart(imagePart, rId, left, top, width, height);
    this._recalculateExtents();
    return this._shapeFactory(pic) as Picture;
  }

  /**
   * Return new |Shape| object appended to this shape tree.
   *
   * `autoshape_type_id` is a member of :ref:`MsoAutoShapeType` e.g. `MSO_SHAPE.RECTANGLE`
   * specifying the type of shape to be added. The remaining arguments specify the new shape's
   * position and size.
   */
  addShape(autoshapeTypeId: MSO_SHAPE, left: number, top: number, width: number, height: number): Shape {
    const autoshapeType = new AutoShapeType(autoshapeTypeId);
    const sp = this._addSp(autoshapeType, left, top, width, height);
    this._recalculateExtents();
    return this._shapeFactory(sp) as Shape;
  }

  /**
   * Return newly added text box shape appended to this shape tree.
   *
   * The text box is of the specified size, located at the specified position on the slide.
   */
  addTextbox(left: number, top: number, width: number, height: number): Shape {
    const sp = this._addTextboxSp(left, top, width, height);
    this._recalculateExtents();
    return this._shapeFactory(sp) as Shape;
  }

  /**
   * Return |FreeformBuilder| object to specify a freeform shape.
   *
   * The optional `start_x` and `start_y` arguments specify the starting pen position in local
   * coordinates. They will be rounded to the nearest integer before use and each default to
   * zero.
   *
   * The optional `scale` argument specifies the size of local coordinates proportional to
   * slide coordinates (EMU). If the vertical scale is different than the horizontal scale
   * (local coordinate units are "rectangular"), a pair of numeric values can be provided as
   * the `scale` argument, e.g. `scale=(1.0, 2.0)`. In this case the first number is
   * interpreted as the horizontal (X) scale and the second as the vertical (Y) scale.
   *
   * A convenient method for calculating scale is to divide a |Length| object by an equivalent
   * count of local coordinate units, e.g. `scale = Inches(1)/1000` for 1000 local units per
   * inch.
   */
  buildFreeform(startX: number = 0, startY: number = 0, scale: [number, number] | number = 1.0): FreeformBuilder {
    const [xScale, yScale] = Array.isArray(scale) ? scale : [scale, scale];
    return FreeformBuilder.new(this, startX, startY, xScale, yScale);
  }

  /**
   * Return the index of `shape` in this sequence.
   *
   * Raises |ValueError| if `shape` is not in the collection.
   */
  index(shape: BaseShape): number {
    const i = [...this._element.iterShapeElms()].indexOf(shape.element);
    if (i === -1) throw new ValueError("shape not in collection");
    return i;
  }

  /**
   * Return new `p:graphicFrame` element appended to this shape tree.
   *
   * The `p:graphicFrame` element has the specified position and size and refers to the chart
   * part identified by `rId`.
   */
  _addChartGraphicFrame(rId: string, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const shapeId = this._nextShapeId;
    const name = `Chart ${(shapeId - 1)}`;
    const graphicFrame = CT_GraphicalObjectFrame.newChartGraphicFrame(shapeId, name, rId, x, y, cx, cy);
    this._spTree.append(graphicFrame);
    return graphicFrame;
  }

  /**
   * Return a newly-added `p:cxnSp` element as specified.
   *
   * The `p:cxnSp` element is for a connector of `connector_type` beginning at (`begin_x`,
   * `begin_y`) and extending to (`end_x`, `end_y`).
   */
  _addCxnSp(connectorType: MSO_CONNECTOR_TYPE, beginX: number, beginY: number, endX: number, endY: number): CT_Connector {
    const id_ = this._nextShapeId;
    const name = `Connector ${(id_ - 1)}`;
    const [flipH, flipV] = [beginX > endX, beginY > endY];
    const [x, y] = [Math.min(beginX, endX), Math.min(beginY, endY)];
    const [cx, cy] = [Math.abs((endX - beginX)), Math.abs((endY - beginY))];
    return this._element.addCxnSp(id_, name, connectorType, x, y, cx, cy, flipH, flipV);
  }

  /**
   * Return a newly appended `p:pic` element as specified.
   *
   * The `p:pic` element displays the image in `image_part` with size and position specified by
   * `x`, `y`, `cx`, and `cy`. The element is appended to the shape tree, causing it to be
   * displayed first in z-order on the slide.
   */
  _addPicFromImagePart(imagePart: ImagePart, rId: string, x: number, y: number, cx: number | null, cy: number | null): CT_Picture {
    const id_ = this._nextShapeId;
    const [scaledCx, scaledCy] = imagePart.scale(cx, cy);
    const name = `Picture ${(id_ - 1)}`;
    const desc = imagePart.desc;
    const pic = this._grpSp.addPic(id_, name, desc, rId, x, y, scaledCx, scaledCy);
    return pic;
  }

  /**
   * Return newly-added `p:sp` element as specified.
   *
   * `p:sp` element is of `autoshape_type` at position (`x`, `y`) and of size (`cx`, `cy`).
   */
  _addSp(autoshapeType: AutoShapeType, x: number, y: number, cx: number, cy: number): CT_Shape {
    const id_ = this._nextShapeId;
    const name = `${autoshapeType.basename} ${(id_ - 1)}`;
    const sp = this._grpSp.addAutoshape(id_, name, autoshapeType.prst, x, y, cx, cy);
    return sp;
  }

  /**
   * Return newly-appended textbox `p:sp` element.
   *
   * Element has position (`x`, `y`) and size (`cx`, `cy`).
   */
  _addTextboxSp(x: number, y: number, cx: number, cy: number): CT_Shape {
    const id_ = this._nextShapeId;
    const name = `TextBox ${(id_ - 1)}`;
    const sp = this._spTree.addTextbox(id_, name, x, y, cx, cy);
    return sp;
  }

  /**
   * Adjust position and size to incorporate all contained shapes.
   *
   * This would typically be called when a contained shape is added, removed, or its position
   * or size updated.
   */
  _recalculateExtents(): void {
    // default behavior is to do nothing, GroupShapes overrides to produce the distinctive
    // behavior of groups and subgroups
  }
}

/**
 * The sequence of child shapes belonging to a group shape.
 *
 * Note that this collection can itself contain a group shape, making this part of a recursive,
 * tree data structure (acyclic graph).
 */
export class GroupShapes extends _BaseGroupShapes {

  /**
   * Adjust position and size to incorporate all contained shapes.
   *
   * This would typically be called when a contained shape is added, removed, or its position
   * or size updated.
   */
  _recalculateExtents(): void {
    this._grpSp.recalculateExtents();
  }
}

/**
 * Sequence of shapes appearing on a slide.
 *
 * The first shape in the sequence is the backmost in z-order and the last shape is topmost.
 * Supports indexed access, len(), index(), and iteration.
 */
export class SlideShapes extends _BaseGroupShapes {
  get parent(): Slide {
    return this._parent;
  }

  /**
   * Return newly added movie shape displaying video in `movie_file`.
   *
   * **EXPERIMENTAL.** This method has important limitations:
   *
   * * The size must be specified; no auto-scaling such as that provided by :meth:`add_picture`
   *   is performed.
   * * The MIME type of the video file should be specified, e.g. 'video/mp4'. The provided
   *   video file is not interrogated for its type. The MIME type `video/unknown` is used by
   *   default (and works fine in tests as of this writing).
   * * A poster frame image must be provided, it cannot be automatically extracted from the
   *   video file. If no poster frame is provided, the default "media loudspeaker" image will
   *   be used.
   *
   * Return a newly added movie shape to the slide, positioned at (`left`, `top`), having size
   * (`width`, `height`), and containing `movie_file`. Before the video is started,
   * `poster_frame_image` is displayed as a placeholder for the video.
   */
  addMovie(movieFile: FileSource, left: number, top: number, width: number, height: number, posterFrameImage: FileSource | null = null, mimeType: string = CT.VIDEO): Movie {
    const moviePic = _MoviePicElementCreator.newMoviePic(this, this._nextShapeId, movieFile, left, top, width, height, posterFrameImage, mimeType);
    this._spTree.append(moviePic);
    this._addVideoTiming(moviePic);
    return this._shapeFactory(moviePic) as Movie;
  }

  /**
   * Add a |GraphicFrame| object containing a table.
   *
   * The table has the specified number of `rows` and `cols` and the specified position and
   * size. `width` is evenly distributed between the columns of the new table. Likewise,
   * `height` is evenly distributed between the rows. Note that the `.table` property on the
   * returned |GraphicFrame| shape must be used to access the enclosed |Table| object.
   */
  addTable(rows: number, cols: number, left: number, top: number, width: number, height: number): GraphicFrame {
    const graphicFrame = this._addGraphicFrameContainingTable(rows, cols, left, top, width, height);
    return this._shapeFactory(graphicFrame) as GraphicFrame;
  }

  /**
   * Add placeholder shapes based on those in `slide_layout`.
   *
   * Z-order of placeholders is preserved. Latent placeholders (date, slide number, and footer)
   * are not cloned.
   */
  cloneLayoutPlaceholders(slideLayout: SlideLayout): void {
    for (const placeholder of slideLayout.iterCloneablePlaceholders()) {
      this.clonePlaceholder(placeholder);
    }
  }

  /** Sequence of placeholder shapes in this slide. */
  get placeholders(): SlidePlaceholders {
    return this.parent.placeholders;
  }

  /**
   * The title placeholder shape on the slide.
   *
   * |None| if the slide has no title placeholder.
   */
  get title(): Shape | null {
    for (const elm of this._spTree.iterPhElms()) {
      if (elm.phIdx === 0) {
        return this._shapeFactory(elm) as Shape;
      }
    }
    return null;
  }

  /** Return a newly added `p:graphicFrame` element containing a table as specified. */
  _addGraphicFrameContainingTable(rows: number, cols: number, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const _id = this._nextShapeId;
    const name = `Table ${(_id - 1)}`;
    const graphicFrame = this._spTree.addTable(_id, name, rows, cols, x, y, cx, cy);
    return graphicFrame;
  }

  /**
   * Add a `p:video` element under `p:sld/p:timing`.
   *
   * The element will refer to the specified `pic` element by its shape id, and cause the video
   * play controls to appear for that video.
   */
  _addVideoTiming(pic: CT_Picture): void {
    const sld: CT_Slide = this._spTree.xpath("/p:sld")[0];
    const childTnLst = sld.getOrAddChildTnLst();
    childTnLst.addVideo(pic.shapeId);
  }

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return SlideShapeFactory(shapeElm, this);
  }
}

/**
 * Sequence of shapes appearing on a slide layout.
 *
 * The first shape in the sequence is the backmost in z-order and the last shape is topmost.
 * Supports indexed access, len(), index(), and iteration.
 */
export class LayoutShapes extends _BaseShapes {

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return _LayoutShapeFactory(shapeElm, this);
  }
}

/**
 * Sequence of shapes appearing on a slide master.
 *
 * The first shape in the sequence is the backmost in z-order and the last shape is topmost.
 * Supports indexed access, len(), and iteration.
 */
export class MasterShapes extends _BaseShapes {

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return _MasterShapeFactory(shapeElm, this);
  }
}

/**
 * Sequence of shapes appearing on a notes slide.
 *
 * The first shape in the sequence is the backmost in z-order and the last shape is topmost.
 * Supports indexed access, len(), index(), and iteration.
 */
export class NotesSlideShapes extends _BaseShapes {

  /**
   * Return the base name for a placeholder of `ph_type` in this shape collection.
   *
   * A notes slide uses a different name for the body placeholder and has some unique
   * placeholder types, so this method overrides the default in the base class.
   */
  phBasename(phType: PP_PLACEHOLDER): string {
    return lookup(NOTES_PH_BASENAMES, phType);
  }

  /** Return appropriate shape object for `shape_elm` appearing on a notes slide. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return _NotesSlideShapeFactory(shapeElm, this);
  }
}

/**
 * Base class for placeholder collections.
 *
 * Subclasses differentiate behaviors for a master, layout, and slide. By default, placeholder
 * shapes are constructed using |BaseShapeFactory|. Subclasses should override
 * :method:`_shape_factory` to use custom placeholder classes.
 */
export class BasePlaceholders<T extends BaseShape = BaseShape> extends _BaseShapes {
  declare [Symbol.iterator]: () => Generator<T>;

  /** True if `shape_elm` is a placeholder shape, False otherwise. */
  static _isMemberElm(shapeElm: ShapeElement): boolean {
    return shapeElm.hasPhElm;
  }
}

/** Sequence of |LayoutPlaceholder| instance for each placeholder shape on a slide layout. */
export class LayoutPlaceholders extends BasePlaceholders<LayoutPlaceholder> {

  /** The first placeholder shape with matching `idx` value, or `default` if not found. */
  get(idx: number, dflt: LayoutPlaceholder | null = null): LayoutPlaceholder | null {
    for (const placeholder of this) {
      if (placeholder.element.phIdx === idx) {
        return placeholder;
      }
    }
    return dflt;
  }

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(shapeElm: ShapeElement): BaseShape {
    return _LayoutShapeFactory(shapeElm, this);
  }
}

/** Sequence of MasterPlaceholder representing the placeholder shapes on a slide master. */
export class MasterPlaceholders<T extends BaseShape = MasterPlaceholder> extends BasePlaceholders<T> {

  /**
   * Return the first placeholder shape with type `ph_type` (e.g. 'body').
   *
   * Returns `default` if no such placeholder shape is present in the collection.
   */
  get(phType: PP_PLACEHOLDER, dflt: T | null = null): T | null {
    for (const placeholder of this) {
      if ((placeholder as any).phType === phType) {
        return placeholder;
      }
    }
    return dflt;
  }

  /** Return an instance of the appropriate shape proxy class for `shape_elm`. */
  _shapeFactory(placeholderElm: ShapeElement): T {
    return _MasterShapeFactory(placeholderElm, this) as T;
  }
}

/** Sequence of placeholder shapes on a notes slide. */
export class NotesSlidePlaceholders extends MasterPlaceholders<NotesSlidePlaceholder> {

  /** Return an instance of the appropriate placeholder proxy class for `placeholder_elm`. */
  _shapeFactory(placeholderElm: ShapeElement): NotesSlidePlaceholder {
    return _NotesSlideShapeFactory(placeholderElm, this) as NotesSlidePlaceholder;
  }
}

/**
 * Collection of placeholder shapes on a slide.
 *
 * Supports iteration, :func:`len`, and dictionary-style lookup on the `idx` value of the
 * placeholders it contains.
 */
export class SlidePlaceholders extends ParentedElementProxy implements Iterable<BaseShape> {
  declare _element: CT_GroupShape;

  /**
   * Access placeholder shape having `idx`.
   *
   * Note that while this looks like list access, idx is actually a dictionary key and will
   * raise |KeyError| if no placeholder with that idx value is in the collection.
   */
  get<T extends BaseShape = SlidePlaceholder>(idx: number): T {
    for (const e of this._element.iterPhElms()) {
      if (e.phIdx === idx) {
        return SlideShapeFactory(e, this) as T;
      }
    }
    throw new KeyError(`no placeholder on this slide with idx == ${idx}`);
  }

  [PY_GETITEM](idx: number): BaseShape {
    return this.get(idx);
  }

  /** Generate placeholder shapes in `idx` order. */
  *[Symbol.iterator](): Generator<BaseShape> {
    const phElms = [...this._element.iterPhElms()].sort((a, b) => a.phIdx - b.phIdx);
    for (const e of phElms) yield SlideShapeFactory(e, this);
  }

  /** Return count of placeholder shapes. */
  get length(): number {
    return [...this._element.iterPhElms()].length;
  }
}

/** Return an instance of the appropriate shape proxy class for `shape_elm`. */
export function BaseShapeFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  const tag = shapeElm.tag;
  if (shapeElm instanceof CT_Picture) {
    const videoFiles = shapeElm.xpath("./p:nvPicPr/p:nvPr/a:videoFile");
    if (videoFiles.length > 0) {
      return new Movie(shapeElm, parent);
    }
    return new Picture(shapeElm, parent);
  }
  const shapeCls = shapeClsFor(tag) ?? BaseShape;
  return new shapeCls(shapeElm as any, parent);
}

type ShapeCls = new (elm: any, parent: ProvidesPart) => BaseShape;
const TAG_CXNSP = qn("p:cxnSp");
const TAG_GRPSP = qn("p:grpSp");
const TAG_SP = qn("p:sp");
const TAG_GRAPHICFRAME = qn("p:graphicFrame");
/** The shape class for a clark tag (looked up on each call, like python's dict literal). */
function shapeClsFor(tag: string): ShapeCls | undefined {
  switch (tag) {
    case TAG_CXNSP:
      return Connector;
    case TAG_GRPSP:
      return GroupShape;
    case TAG_SP:
      return Shape;
    case TAG_GRAPHICFRAME:
      return GraphicFrame;
  }
  return undefined;
}

/** Return appropriate shape object for `shape_elm` on a slide layout. */
export function _LayoutShapeFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  if (shapeElm instanceof CT_Shape && shapeElm.hasPhElm) {
    return new LayoutPlaceholder(shapeElm, parent);
  }
  return BaseShapeFactory(shapeElm, parent);
}

/** Return appropriate shape object for `shape_elm` on a slide master. */
export function _MasterShapeFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  if (shapeElm instanceof CT_Shape && shapeElm.hasPhElm) {
    return new MasterPlaceholder(shapeElm, parent);
  }
  return BaseShapeFactory(shapeElm, parent);
}

/** Return appropriate shape object for `shape_elm` on a notes slide. */
export function _NotesSlideShapeFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  if (shapeElm instanceof CT_Shape && shapeElm.hasPhElm) {
    return new NotesSlidePlaceholder(shapeElm, parent);
  }
  return BaseShapeFactory(shapeElm, parent);
}

/** Return a placeholder shape of the appropriate type for `shape_elm`. */
export function _SlidePlaceholderFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  const tag = shapeElm.tag;
  if (tag === qn("p:sp")) {
    const phType = shapeElm.phType;
    const Constructor =
      phType === PP_PLACEHOLDER.BITMAP || phType === PP_PLACEHOLDER.PICTURE
        ? PicturePlaceholder
        : phType === PP_PLACEHOLDER.CHART
          ? ChartPlaceholder
          : phType === PP_PLACEHOLDER.TABLE
            ? TablePlaceholder
            : SlidePlaceholder;
    return new Constructor(shapeElm as CT_Shape, parent);
  }
  if (tag === qn("p:graphicFrame")) return new PlaceholderGraphicFrame(shapeElm as CT_GraphicalObjectFrame, parent);
  if (tag === qn("p:pic")) return new PlaceholderPicture(shapeElm as CT_Picture, parent);
  return BaseShapeFactory(shapeElm, parent);
}

/** Return appropriate shape object for `shape_elm` on a slide. */
export function SlideShapeFactory(shapeElm: ShapeElement, parent: ProvidesPart): BaseShape {
  if (shapeElm.hasPhElm) {
    return _SlidePlaceholderFactory(shapeElm, parent);
  }
  return BaseShapeFactory(shapeElm, parent);
}

/**
 * Functional service object for creating a new movie p:pic element.
 *
 * It's entire external interface is its :meth:`new_movie_pic` class method that returns a new
 * `p:pic` element containing the specified video. This class is not intended to be constructed
 * or an instance of it retained by the caller; it is a "one-shot" object, really a function
 * wrapped in a object such that its helper methods can be organized here.
 */
export class _MoviePicElementCreator {
  declare _shapes: SlideShapes;
  declare _shapeId: number;
  declare _movieFile: FileSource;
  declare _x: number;
  declare _y: number;
  declare _cx: number;
  declare _cy: number;
  declare _posterFrameFile: FileSource | null;
  declare _mimeType: string | null;

  constructor(shapes: SlideShapes, shapeId: number, movieFile: FileSource, x: number, y: number, cx: number, cy: number, posterFrameFile: FileSource | null, mimeType: string | null) {
    this._shapes = shapes;
    this._shapeId = shapeId;
    this._movieFile = movieFile;
    [this._x, this._y, this._cx, this._cy] = [x, y, cx, cy];
    this._posterFrameFile = posterFrameFile;
    this._mimeType = mimeType;
  }

  /**
   * Return a new `p:pic` element containing video in `movie_file`.
   *
   * If `mime_type` is None, 'video/unknown' is used. If `poster_frame_file` is None, the
   * default "media loudspeaker" image is used.
   */
  static newMoviePic(shapes: SlideShapes, shapeId: number, movieFile: FileSource, x: number, y: number, cx: number, cy: number, posterFrameImage: FileSource | null, mimeType: string | null): CT_Picture {
    return new this(shapes, shapeId, movieFile, x, y, cx, cy, posterFrameImage, mimeType)._pic;
  }

  /**
   * Return the rId of RT.MEDIA relationship to video part.
   *
   * For historical reasons, there are two relationships to the same part; one is the video rId
   * and the other is the media rId.
   */
  get _mediaRId(): string {
    return this._videoPartRIds[0];
  }

  /** Return the new `p:pic` element referencing the video. */
  get _pic(): CT_Picture {
    return lazy(this, "_pic", () => {
      return CT_Picture.newVideoPic(this._shapeId, this._shapeName, this._videoRId, this._mediaRId, this._posterFrameRId, this._x, this._y, this._cx, this._cy);
    });
  }

  /**
   * Return the image file for video placeholder image.
   *
   * If no poster frame file is provided, the default "media loudspeaker" image is used.
   */
  get _posterFrameImageFile(): FileSource {
    return lazy(this, "_posterFrameImageFile", () => {
      const posterFrameFile = this._posterFrameFile;
      if (posterFrameFile === null) {
        return new BytesIO(SPEAKER_IMAGE_BYTES);
      }
      return posterFrameFile;
    });
  }

  /**
   * Return the rId of relationship to poster frame image.
   *
   * The poster frame is the image used to represent the video before it's played.
   */
  get _posterFrameRId(): string {
    return lazy(this, "_posterFrameRId", () => {
      const [, posterFrameRId] = this._slidePart.getOrAddImagePart(this._posterFrameImageFile);
      return posterFrameRId;
    });
  }

  /**
   * Return the appropriate shape name for the p:pic shape.
   *
   * A movie shape is named with the base filename of the video.
   */
  get _shapeName(): string {
    return this._video.filename;
  }

  /** Return SlidePart object for slide containing this movie. */
  get _slidePart(): SlidePart {
    return this._shapes.part;
  }

  /** Return a |Video| object containing the movie file. */
  get _video(): Video {
    return lazy(this, "_video", () => {
      return Video.fromPathOrFileLike(this._movieFile, this._mimeType);
    });
  }

  /**
   * Return the rIds for relationships to media part for video.
   *
   * This is where the media part and its relationships to the slide are actually created.
   */
  get _videoPartRIds(): [string, string] {
    return lazy(this, "_videoPartRIds", () => {
      const [mediaRId, videoRId] = this._slidePart.getOrAddVideoMediaPart(this._video);
      return [mediaRId, videoRId];
    });
  }

  /**
   * Return the rId of RT.VIDEO relationship to video part.
   *
   * For historical reasons, there are two relationships to the same part; one is the video rId
   * and the other is the media rId.
   */
  get _videoRId(): string {
    return this._videoPartRIds[1];
  }
}

/**
 * Functional service object for creating a new OLE-object p:graphicFrame element.
 *
 * It's entire external interface is its :meth:`graphicFrame` class method that returns a new
 * `p:graphicFrame` element containing the specified embedded OLE-object shape. This class is not
 * intended to be constructed or an instance of it retained by the caller; it is a "one-shot"
 * object, really a function wrapped in a object such that its helper methods can be organized
 * here.
 */
export class _OleObjectElementCreator {
  declare _shapes: _BaseGroupShapes;
  declare _shapeId: number;
  declare _oleObjectFile: FileSource;
  declare _progIdArg: PROG_ID | string;
  declare _x: number;
  declare _y: number;
  declare _cxArg: number | null;
  declare _cyArg: number | null;
  declare _iconFileArg: FileSource | null;
  declare _iconWidthArg: number | null;
  declare _iconHeightArg: number | null;

  constructor(shapes: _BaseGroupShapes, shapeId: number, oleObjectFile: FileSource, progId: PROG_ID | string, x: number, y: number, cx: number | null, cy: number | null, iconFile: FileSource | null, iconWidth: number | null, iconHeight: number | null) {
    this._shapes = shapes;
    this._shapeId = shapeId;
    this._oleObjectFile = oleObjectFile;
    this._progIdArg = progId;
    this._x = x;
    this._y = y;
    this._cxArg = cx;
    this._cyArg = cy;
    this._iconFileArg = iconFile;
    this._iconWidthArg = iconWidth;
    this._iconHeightArg = iconHeight;
  }

  /** Return new `p:graphicFrame` element containing embedded `ole_object_file`. */
  static graphicFrame(shapes: _BaseGroupShapes, shapeId: number, oleObjectFile: FileSource, progId: PROG_ID | string, x: number, y: number, cx: number | null, cy: number | null, iconFile: FileSource | null, iconWidth: number | null, iconHeight: number | null): CT_GraphicalObjectFrame {
    return new this(shapes, shapeId, oleObjectFile, progId, x, y, cx, cy, iconFile, iconWidth, iconHeight)._graphicFrame;
  }

  /** Newly-created `p:graphicFrame` element referencing embedded OLE-object. */
  get _graphicFrame(): CT_GraphicalObjectFrame {
    return lazy(this, "_graphicFrame", () => {
      this._iconImageFile;
      return CT_GraphicalObjectFrame.newOleObjectGraphicFrame(this._shapeId, this._shapeName, this._oleObjectRId, this._progId, this._iconRId, this._x, this._y, this._cx, this._cy, this._iconWidth, this._iconHeight);
    });
  }

  /** Emu object specifying width of "show-as-icon" image for OLE shape. */
  get _cx(): number {
    return lazy(this, "_cx", () => {
      if (this._cxArg !== null) {
        return this._cxArg;
      }
      return this._progIdArg instanceof PROG_ID ? Emu(this._progIdArg.width) : Emu(965200);
    });
  }

  /** Emu object specifying height of "show-as-icon" image for OLE shape. */
  get _cy(): number {
    return lazy(this, "_cy", () => {
      if (this._cyArg !== null) {
        return this._cyArg;
      }
      return this._progIdArg instanceof PROG_ID ? Emu(this._progIdArg.height) : Emu(609600);
    });
  }

  /**
   * Vertical size of enclosed EMF icon within the OLE graphic-frame.
   *
   * This must be specified when a custom icon is used, to avoid stretching of the image and
   * possible undesired resizing by PowerPoint when the OLE shape is double-clicked to open it.
   *
   * The correct size can be determined by creating an example PPTX using PowerPoint and then
   * inspecting the XML of the OLE graphics-frame (p:oleObj.imgH).
   */
  get _iconHeight(): number {
    return lazy(this, "_iconHeight", () => {
      return (this._iconHeightArg !== null ? this._iconHeightArg : Emu(609600));
    });
  }

  /**
   * Reference to image file containing icon to show in lieu of this object.
   *
   * This can be either a str path or a file-like object (io.BytesIO typically).
   */
  get _iconImageFile(): FileSource {
    return lazy(this, "_iconImageFile", () => {
      if (this._iconFileArg !== null) {
        return this._iconFileArg;
      }
      // python opens pptx/templates/<icon>; the icons are bundled in src/templates/data.ts
      const iconFilename = this._progIdArg instanceof PROG_ID ? this._progIdArg.iconFilename : "generic-icon.emf";
      return new BytesIO(templateBytes(iconFilename));
    });
  }

  /** str rId like "rId7" of rel to icon (image) representing OLE-object part. */
  get _iconRId(): string {
    return lazy(this, "_iconRId", () => {
      const [, rId] = this._slidePart.getOrAddImagePart(this._iconImageFile);
      return rId;
    });
  }

  /**
   * Width of enclosed EMF icon within the OLE graphic-frame.
   *
   * This must be specified when a custom icon is used, to avoid stretching of the image and
   * possible undesired resizing by PowerPoint when the OLE shape is double-clicked to open it.
   */
  get _iconWidth(): number {
    return lazy(this, "_iconWidth", () => {
      return (this._iconWidthArg !== null ? this._iconWidthArg : Emu(965200));
    });
  }

  /**
   * str rId like "rId6" of relationship to embedded ole_object part.
   *
   * This is where the ole_object part and its relationship to the slide are actually created.
   */
  get _oleObjectRId(): string {
    return lazy(this, "_oleObjectRId", () => {
      return this._slidePart.addEmbeddedOleObjectPart(this._progIdArg, this._oleObjectFile);
    });
  }

  /**
   * str like "Excel.Sheet.12" identifying program used to open object.
   *
   * This value appears in the `progId` attribute of the `p:oleObj` element for the object.
   */
  get _progId(): string {
    return lazy(this, "_progId", () => {
      const progIdArg = this._progIdArg;
      return (progIdArg instanceof PROG_ID ? progIdArg.progId : progIdArg);
    });
  }

  /**
   * str name like "Object 1" for the embedded ole_object shape.
   *
   * The name is formed from the prefix "Object " and the shape-id decremented by 1.
   */
  get _shapeName(): string {
    return lazy(this, "_shapeName", () => {
      return `Object ${(this._shapeId - 1)}`;
    });
  }

  /** SlidePart object for this slide. */
  get _slidePart(): SlidePart {
    return lazy(this, "_slidePart", () => {
      return this._shapes.part;
    });
  }
}

const PH_BASENAMES = new Map<PP_PLACEHOLDER, string>([
  [PP_PLACEHOLDER.BITMAP, "ClipArt Placeholder"],
  [PP_PLACEHOLDER.BODY, "Text Placeholder"],
  [PP_PLACEHOLDER.CENTER_TITLE, "Title"],
  [PP_PLACEHOLDER.CHART, "Chart Placeholder"],
  [PP_PLACEHOLDER.DATE, "Date Placeholder"],
  [PP_PLACEHOLDER.FOOTER, "Footer Placeholder"],
  [PP_PLACEHOLDER.HEADER, "Header Placeholder"],
  [PP_PLACEHOLDER.MEDIA_CLIP, "Media Placeholder"],
  [PP_PLACEHOLDER.OBJECT, "Content Placeholder"],
  [PP_PLACEHOLDER.ORG_CHART, "SmartArt Placeholder"],
  [PP_PLACEHOLDER.PICTURE, "Picture Placeholder"],
  [PP_PLACEHOLDER.SLIDE_NUMBER, "Slide Number Placeholder"],
  [PP_PLACEHOLDER.SUBTITLE, "Subtitle"],
  [PP_PLACEHOLDER.TABLE, "Table Placeholder"],
  [PP_PLACEHOLDER.TITLE, "Title"],
]);

const NOTES_PH_BASENAMES = new Map<PP_PLACEHOLDER, string>([
  [PP_PLACEHOLDER.BODY, "Notes Placeholder"],
  [PP_PLACEHOLDER.DATE, "Date Placeholder"],
  [PP_PLACEHOLDER.FOOTER, "Footer Placeholder"],
  [PP_PLACEHOLDER.HEADER, "Header Placeholder"],
  [PP_PLACEHOLDER.SLIDE_IMAGE, "Slide Image Placeholder"],
  [PP_PLACEHOLDER.SLIDE_NUMBER, "Slide Number Placeholder"],
]);

/** python `d[key]` on a dict keyed by enum members: KeyError when absent. */
function lookup<K, V>(m: Map<K, V>, key: K): V {
  const v = m.get(key);
  if (v === undefined) throw new KeyError(String(key));
  return v;
}
