// Graphic Frame shape and related objects. (`pptx.shapes.graphfrm`)
import { ValueError, NotImplementedError } from "../errors.js";
import { lazy } from "../util.js";
import { MSO_SHAPE_TYPE } from "../enum/shapes.js";
import { BaseShape } from "./base.js";
import { ParentedElementProxy } from "../shared.js";
import { GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_OLEOBJ, GRAPHIC_DATA_URI_TABLE } from "../spec.js";
import { Table } from "../table.js";
// `pptx.chart.chart.Chart` is ported in stage 5 (gljeremy/pptx-ts#5)
type Chart = any;
import type { ShadowFormat } from "../dml/effect.js";
import type { CT_GraphicalObjectData, CT_GraphicalObjectFrame } from "../oxml/shapes/graphfrm.js";
import type { ChartPart } from "../parts/chart.js";
import type { BaseSlidePart } from "../parts/slide.js";
import type { ProvidesPart } from "../types.js";

/**
 * Container shape for table, chart, smart art, and media objects.
 *
 * Corresponds to a `p:graphicFrame` element in the shape tree.
 */
export class GraphicFrame extends BaseShape {
  declare _graphicFrame: any;

  constructor(graphicFrame: CT_GraphicalObjectFrame, parent: ProvidesPart) {
    super(graphicFrame, parent);
    this._graphicFrame = graphicFrame;
  }

  /**
   * The |Chart| object containing the chart in this graphic frame.
   *
   * Raises |ValueError| if this graphic frame does not contain a chart.
   */
  get chart(): Chart {
    if (!this.hasChart) {
      throw new ValueError("shape does not contain a chart");
    }
    return this.chartPart.chart;
  }

  /** The |ChartPart| object containing the chart in this graphic frame. */
  get chartPart(): ChartPart {
    const chartRId = this._graphicFrame.chartRId;
    if (chartRId === null) {
      throw new ValueError("this graphic frame does not contain a chart");
    }
    return this.part.relatedPart(chartRId);
  }

  /**
   * |True| if this graphic frame contains a chart object. |False| otherwise.
   *
   * When |True|, the chart object can be accessed using the `.chart` property.
   */
  get hasChart(): boolean {
    return this._graphicFrame.graphicDataUri === GRAPHIC_DATA_URI_CHART;
  }

  /**
   * |True| if this graphic frame contains a table object, |False| otherwise.
   *
   * When |True|, the table object can be accessed using the `.table` property.
   */
  get hasTable(): boolean {
    return this._graphicFrame.graphicDataUri === GRAPHIC_DATA_URI_TABLE;
  }

  /**
   * _OleFormat object for this graphic-frame shape.
   *
   * Raises `ValueError` on a GraphicFrame instance that does not contain an OLE object.
   *
   * An shape that contains an OLE object will have `.shape_type` of either
   * `EMBEDDED_OLE_OBJECT` or `LINKED_OLE_OBJECT`.
   */
  get oleFormat(): _OleFormat {
    if (!this._graphicFrame.hasOleobj) {
      throw new ValueError("not an OLE-object shape");
    }
    return new _OleFormat(this._graphicFrame.graphicData, this._parent);
  }

  /**
   * Unconditionally raises |NotImplementedError|.
   *
   * Access to the shadow effect for graphic-frame objects is content-specific (i.e. different
   * for charts, tables, etc.) and has not yet been implemented.
   */
  get shadow(): ShadowFormat {
    return lazy(this, "shadow", () => {
      throw new NotImplementedError("shadow property on GraphicFrame not yet supported");
    });
  }

  /**
   * Optional member of `MSO_SHAPE_TYPE` identifying the type of this shape.
   *
   * Possible values are `MSO_SHAPE_TYPE.CHART`, `MSO_SHAPE_TYPE.TABLE`,
   * `MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT`, `MSO_SHAPE_TYPE.LINKED_OLE_OBJECT`.
   *
   * This value is `None` when none of these four types apply, for example when the shape
   * contains SmartArt.
   */
  get shapeType(): MSO_SHAPE_TYPE | null {
    const graphicDataUri = this._graphicFrame.graphicDataUri;
    if (graphicDataUri === GRAPHIC_DATA_URI_CHART) {
      return MSO_SHAPE_TYPE.CHART;
    } else if (graphicDataUri === GRAPHIC_DATA_URI_TABLE) {
      return MSO_SHAPE_TYPE.TABLE;
    } else if (graphicDataUri === GRAPHIC_DATA_URI_OLEOBJ) {
      return (this._graphicFrame.isEmbeddedOleObj ? MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT : MSO_SHAPE_TYPE.LINKED_OLE_OBJECT);
    } else {
      return null;
    }
  }

  /**
   * The |Table| object contained in this graphic frame.
   *
   * Raises |ValueError| if this graphic frame does not contain a table.
   */
  get table(): Table {
    if (!this.hasTable) {
      throw new ValueError("shape does not contain a table");
    }
    const tbl = this._graphicFrame.graphic.graphicData.tbl;
    return new Table(tbl, this);
  }
}

/** Provides attributes on an embedded OLE object. */
export class _OleFormat extends ParentedElementProxy {
  get part(): BaseSlidePart {
    return this._parent.part;
  }
  declare _graphicData: any;

  constructor(graphicData: CT_GraphicalObjectData, parent: ProvidesPart) {
    super(graphicData, parent);
    this._graphicData = graphicData;
  }

  /**
   * Optional bytes of OLE object, suitable for loading or saving as a file.
   *
   * This value is `None` if the embedded object does not represent a "file".
   */
  get blob(): Uint8Array | null {
    const blobRId = this._graphicData.blobRId;
    if (blobRId === null) {
      return null;
    }
    return this.part.relatedPart(blobRId).blob;
  }

  /**
   * str "progId" attribute of this embedded OLE object.
   *
   * The progId is a str like "Excel.Sheet.12" that identifies the "file-type" of the embedded
   * object, or perhaps more precisely, the application (aka. "server" in OLE parlance) to be
   * used to open this object.
   */
  get progId(): string | null {
    return this._graphicData.progId;
  }

  /** True when OLE object should appear as an icon (rather than preview). */
  get showAsIcon(): boolean | null {
    return this._graphicData.showAsIcon;
  }
}
