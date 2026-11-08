// lxml custom element classes for theme-related XML elements. (`pptx.oxml.theme`)
import { parseFromTemplate } from "./parser.js";
import { BaseOxmlElement } from "./xmlchemy.js";

/** ``<a:theme>`` element, root of a theme part */
export class CT_OfficeStyleSheet extends BaseOxmlElement {

  /**
   * Return a new ``<a:theme>`` element containing default settings
   * suitable for use with a notes master.
   */
  static newDefault(): CT_OfficeStyleSheet {
    return parseFromTemplate("theme");
  }
}
