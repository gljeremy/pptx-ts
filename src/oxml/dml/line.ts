// lxml custom element classes for DrawingML line-related XML elements. (`pptx.oxml.dml.line`)
import { defineChildren } from "../xmlchemy.js";
import { MSO_LINE_DASH_STYLE } from "../../enum/dml.js";
import { BaseOxmlElement, OptionalAttribute } from "../xmlchemy.js";

/** `a:prstDash` custom element class */
export class CT_PresetLineDashProperties extends BaseOxmlElement {
  declare val: any;
  static {
    defineChildren(this, {
      val: new OptionalAttribute("val", MSO_LINE_DASH_STYLE),
    });
  }
}
