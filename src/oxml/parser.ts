// The oxml parser: XML parsing with custom element class lookup (`pptx.oxml` functions).
// Element classes are registered in oxml/index.ts, which mirrors python-pptx's
// `pptx/oxml/__init__.py`.
import { NamespacePrefixedTag } from "./ns.js";
import { parseXml as rawParse, registerClass, XmlElement } from "./xml.js";
import { templateText } from "../templates/index.js";

/** Root element obtained by parsing the XML in `xml`, using the custom element classes. */
export function parseXml(xml: string | Uint8Array | ArrayBuffer): any {
  return rawParse(xml, { removeBlankText: true });
}

/** Element loaded from the XML template named `templateName` (e.g. "theme"). */
export function parseFromTemplate(templateName: string): any {
  return parseXml(templateText(templateName));
}

/** Register `cls` to be constructed when the parser encounters an element named `nsptagname`. */
export function registerElementCls(nsptagname: string, cls: new () => XmlElement): void {
  registerClass(new NamespacePrefixedTag(nsptagname).clarkName, cls);
}

/** lxml `oxml_parser.makeelement`. */
export function makeelement(clarkName: string, nsmap?: Record<string, string> | null): any {
  return XmlElement.make(clarkName, nsmap ?? null);
}
