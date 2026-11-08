// lxml custom element class for CT_GraphicalObjectFrame XML element. (`pptx.oxml.shapes.graphfrm`)
import { defineChildren } from "../xmlchemy.js";
import { parseXml } from "../parser.js";
import { CT_Chart } from "../chart/chart.js";
import { nsdecls } from "../ns.js";
import { BaseShapeElement } from "./shared.js";
import { XsdBoolean, XsdString } from "../simpletypes.js";
import { CT_Table } from "../table.js";
import { BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute, ZeroOrOne } from "../xmlchemy.js";
import { GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_OLEOBJ, GRAPHIC_DATA_URI_TABLE } from "../../spec.js";
import type { CT_ApplicationNonVisualDrawingProps, CT_NonVisualDrawingProps, CT_Transform2D } from "./shared.js";

/**
 * `a:graphic` element.
 *
 * The container for the reference to or definition of the framed graphical object (table, chart,
 * etc.).
 */
export class CT_GraphicalObject extends BaseOxmlElement {
  declare graphicData: CT_GraphicalObjectData;
  static {
    defineChildren(this, {
      graphicData: new OneAndOnlyOne("a:graphicData"),
    });
  }

  /** The `c:chart` grandchild element, or |None| if not present. */
  get chart(): CT_Chart | null {
    return this.graphicData.chart;
  }
}

/**
 * `p:graphicData` element.
 *
 * The direct container for a table, a chart, or another graphical object.
 */
export class CT_GraphicalObjectData extends BaseShapeElement {
  declare chart: CT_Chart | null;
  declare getOrAddChart: () => CT_Chart;
  declare _addChart: (attrs?: Record<string, any>) => CT_Chart;
  declare _insertChart: (child: CT_Chart) => CT_Chart;
  declare _newChart: () => CT_Chart;
  declare _removeChart: () => void;
  declare tbl: CT_Table | null;
  declare getOrAddTbl: () => CT_Table;
  declare _addTbl: (attrs?: Record<string, any>) => CT_Table;
  declare _insertTbl: (child: CT_Table) => CT_Table;
  declare _newTbl: () => CT_Table;
  declare _removeTbl: () => void;
  declare uri: string;
  static {
    defineChildren(this, {
      chart: new ZeroOrOne("c:chart"),
      tbl: new ZeroOrOne("a:tbl"),
      uri: new RequiredAttribute("uri", XsdString),
    });
  }

  /**
   * Optional `r:id` attribute value of `p:oleObj` descendent element.
   *
   * This value is `None` when this `p:graphicData` element does not enclose an OLE object.
   * This value could also be `None` if an enclosed OLE object does not specify this attribute
   * (it is specified optional in the schema) but so far, all OLE objects we've encountered
   * specify this value.
   */
  get blobRId(): string | null {
    return (this._oleObj === null ? null : this._oleObj.rId);
  }

  /**
   * Optional boolean indicating an embedded OLE object.
   *
   * Returns `None` when this `p:graphicData` element does not enclose an OLE object. `True`
   * indicates an embedded OLE object and `False` indicates a linked OLE object.
   */
  get isEmbeddedOleObj(): boolean | null {
    return (this._oleObj === null ? null : this._oleObj.isEmbedded);
  }

  /**
   * Optional str value of "progId" attribute of `p:oleObj` descendent.
   *
   * This value identifies the "type" of the embedded object in terms of the application used
   * to open it.
   *
   * This value is `None` when this `p:graphicData` element does not enclose an OLE object.
   * This could also be `None` if an enclosed OLE object does not specify this attribute (it is
   * specified optional in the schema) but so far, all OLE objects we've encountered specify
   * this value.
   */
  get progId(): string | null {
    return (this._oleObj === null ? null : this._oleObj.progId);
  }

  /**
   * Optional value of "showAsIcon" attribute value of `p:oleObj` descendent.
   *
   * This value is `None` when this `p:graphicData` element does not enclose an OLE object. It
   * is False when the `showAsIcon` attribute is omitted on the `p:oleObj` element.
   */
  get showAsIcon(): boolean | null {
    return (this._oleObj === null ? null : this._oleObj.showAsIcon);
  }

  /**
   * Optional `p:oleObj` element contained in this `p:graphicData' element.
   *
   * Returns `None` when this graphic-data element does not enclose an OLE object. Note that
   * this returns the last `p:oleObj` element found. There can be more than one `p:oleObj`
   * element because an `mc.AlternateContent` element may appear as the child of
   * `p:graphicData` and that alternate-content subtree can contain multiple compatibility
   * choices. The last one should suit best for reading purposes because it contains the lowest
   * common denominator.
   */
  get _oleObj(): CT_OleObject | null {
    const oleObjs = this.xpath(".//p:oleObj");
    return oleObjs.length ? oleObjs.at(-1)! : null;
  }
}

/**
 * `p:graphicFrame` element.
 *
 * A container for a table, a chart, or another graphical object.
 */
export class CT_GraphicalObjectFrame extends BaseShapeElement {
  declare nvGraphicFramePr: CT_GraphicalObjectFrameNonVisual;

  declare graphic: CT_GraphicalObject;
  static {
    defineChildren(this, {
      nvGraphicFramePr: new OneAndOnlyOne("p:nvGraphicFramePr"),
      xfrm: new OneAndOnlyOne("p:xfrm"),
      graphic: new OneAndOnlyOne("a:graphic"),
    });
  }

  /** The `c:chart` great-grandchild element, or |None| if not present. */
  get chart(): CT_Chart | null {
    return this.graphic.chart;
  }

  /**
   * The `rId` attribute of the `c:chart` great-grandchild element.
   *
   * |None| if not present.
   */
  get chartRId(): string | null {
    const chart = this.chart;
    if (chart === null) {
      return null;
    }
    return chart.rId;
  }

  /**
   * Return the required `p:xfrm` child element.
   *
   * Overrides version on BaseShapeElement.
   */
  getOrAddXfrm(): CT_Transform2D {
    return this.xfrm!;
  }

  /** `a:graphicData` grandchild of this graphic-frame element. */
  get graphicData(): CT_GraphicalObjectData {
    return this.graphic.graphicData;
  }

  /** str value of `uri` attribute of `a:graphicData` grandchild. */
  get graphicDataUri(): string {
    return this.graphic.graphicData.uri;
  }

  /** `True` for graphicFrame containing an OLE object, `False` otherwise. */
  get hasOleobj(): boolean {
    return this.graphicData.uri === GRAPHIC_DATA_URI_OLEOBJ;
  }

  /**
   * Optional boolean indicating an embedded OLE object.
   *
   * Returns `None` when this `p:graphicFrame` element does not enclose an OLE object. `True`
   * indicates an embedded OLE object and `False` indicates a linked OLE object.
   */
  get isEmbeddedOleObj(): boolean | null {
    return this.graphicData.isEmbeddedOleObj;
  }

  /** Return a `p:graphicFrame` element tree populated with a chart element. */
  static newChartGraphicFrame(id_: number, name: string, rId: string, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const graphicFrame = CT_GraphicalObjectFrame.newGraphicFrame(id_, name, x, y, cx, cy);
    const graphicData = graphicFrame.graphic.graphicData;
    graphicData.uri = GRAPHIC_DATA_URI_CHART;
    graphicData.append(CT_Chart.newChart(rId));
    return graphicFrame;
  }

  /**
   * Return a new `p:graphicFrame` element tree suitable for containing a table or chart.
   *
   * Note that a graphicFrame element is not a valid shape until it contains a graphical object
   * such as a table.
   */
  static newGraphicFrame(id_: number, name: string, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    return parseXml(`<p:graphicFrame ${nsdecls("a", "p")}>
  <p:nvGraphicFramePr>
    <p:cNvPr id="${id_}" name="${name}"/>
    <p:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1"/>
    </p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="${x}" y="${y}"/>
    <a:ext cx="${cx}" cy="${cy}"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData/>
  </a:graphic>
</p:graphicFrame>`);
  }

  /**
   * Return newly-created `p:graphicFrame` for embedded OLE-object.
   *
   * `ole_object_rId` identifies the relationship to the OLE-object part.
   *
   * `progId` is a str identifying the object-type in terms of the application (program) used
   * to open it. This becomes an attribute of the same name in the `p:oleObj` element.
   *
   * `icon_rId` identifies the relationship to an image part used to display the OLE-object as
   * an icon (vs. a preview).
   */
  static newOleObjectGraphicFrame(id_: number, name: string, oleObjectRId: string, progId: string, iconRId: string, x: number, y: number, cx: number, cy: number, imgW: number, imgH: number): CT_GraphicalObjectFrame {
    return parseXml(`<p:graphicFrame ${nsdecls("a", "p", "r")}>
  <p:nvGraphicFramePr>
    <p:cNvPr id="${id_}" name="${name}"/>
    <p:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1"/>
    </p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="${x}" y="${y}"/>
    <a:ext cx="${cx}" cy="${cy}"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData        uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
      <p:oleObj showAsIcon="1"                r:id="${oleObjectRId}"                imgW="${imgW}"                imgH="${imgH}"                progId="${progId}">
        <p:embed/>
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="0" name=""/>
            <p:cNvPicPr/>
            <p:nvPr/>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip r:embed="${iconRId}"/>
            <a:stretch>
              <a:fillRect/>
            </a:stretch>
          </p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="${x}" y="${y}"/>
              <a:ext cx="${cx}" cy="${cy}"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
          </p:spPr>
        </p:pic>
      </p:oleObj>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`);
  }

  /** Return a `p:graphicFrame` element tree populated with a table element. */
  static newTableGraphicFrame(id_: number, name: string, rows: number, cols: number, x: number, y: number, cx: number, cy: number): CT_GraphicalObjectFrame {
    const graphicFrame = this.newGraphicFrame(id_, name, x, y, cx, cy);
    graphicFrame.graphic.graphicData.uri = GRAPHIC_DATA_URI_TABLE;
    graphicFrame.graphic.graphicData.append(CT_Table.newTbl(rows, cols, cx, cy));
    return graphicFrame;
  }
}

/**
 * `p:nvGraphicFramePr` element.
 *
 * This contains the non-visual properties of a graphic frame, such as name, id, etc.
 */
export class CT_GraphicalObjectFrameNonVisual extends BaseOxmlElement {
  declare cNvPr: CT_NonVisualDrawingProps;
  declare nvPr: CT_ApplicationNonVisualDrawingProps;
  static {
    defineChildren(this, {
      cNvPr: new OneAndOnlyOne("p:cNvPr"),
      nvPr: new OneAndOnlyOne("p:nvPr"),
    });
  }
}

/**
 * `p:oleObj` element, container for an OLE object (e.g. Excel file).
 *
 * An OLE object can be either linked or embedded (hence the name).
 */
export class CT_OleObject extends BaseOxmlElement {
  declare progId: string | null;
  declare rId: string | null;
  declare showAsIcon: boolean;
  static {
    defineChildren(this, {
      progId: new OptionalAttribute("progId", XsdString),
      rId: new OptionalAttribute("r:id", XsdString),
      showAsIcon: new OptionalAttribute("showAsIcon", XsdBoolean, false),
    });
  }

  /** True when this OLE object is embedded, False when it is linked. */
  get isEmbedded(): boolean {
    return this.xpath("./p:embed").length > 0;
  }
}
