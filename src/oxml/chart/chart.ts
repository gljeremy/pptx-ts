// Custom element classes for top-level chart-related XML elements (`pptx.oxml.chart.chart`).
//
// Stage 5 (gljeremy/pptx-ts#5) ports the rest of this module and the other chart element
// classes; `CT_Chart`'s relationship id and constructor are here already because graphic
// frames reference charts by `c:chart/@r:id`.
import { nsdecls } from "../ns.js";
import { parseXml } from "../parser.js";
import { XsdString } from "../simpletypes.js";
import { BaseOxmlElement, defineChildren, RequiredAttribute } from "../xmlchemy.js";

/** `c:chart` custom element class. */
export class CT_Chart extends BaseOxmlElement {
  declare rId: string;
  static {
    defineChildren(this, {
      rId: new RequiredAttribute("r:id", XsdString),
    });
  }

  /** A new `c:chart` element. */
  static newChart(rId: string): CT_Chart {
    return parseXml(`<c:chart ${nsdecls("c")} ${nsdecls("r")} r:id="${rId}"/>`) as CT_Chart;
  }
}
