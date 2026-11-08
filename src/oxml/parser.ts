// The oxml parser: XML parsing with custom element class lookup (`pptx.oxml` functions).
// Element classes are registered in oxml/index.ts, which mirrors python-pptx's
// `pptx/oxml/__init__.py`.
import { NamespacePrefixedTag } from "./ns.js";
import { ValueError } from "../errors.js";
import { encodeUtf8, parseXml as rawParse, registerClass, XmlElement } from "./xml.js";
import { templateText } from "../templates/index.js";

/**
 * Root element obtained by parsing the XML in `xml`, using the custom element classes.
 *
 * Like lxml, a string (rather than bytes) that starts with an XML declaration naming an
 * encoding is rejected with ValueError, since the text is already decoded.
 */
export function parseXml(xml: string | Uint8Array | ArrayBuffer): any {
  if (typeof xml === "string" && /^\s*<\?xml[^>]*\sencoding\s*=/.test(xml)) {
    throw new ValueError("Unicode strings with encoding declaration are not supported. Please use bytes input or XML fragments without declaration.");
  }
  return rawParse(xml, { removeBlankText: true });
}

/** Element loaded from the XML template named `templateName` (e.g. "theme"). */
export function parseFromTemplate(templateName: string): any {
  // python reads the template file as bytes
  return parseXml(encodeUtf8(templateText(templateName)));
}

/** Register `cls` to be constructed when the parser encounters an element named `nsptagname`. */
export function registerElementCls(nsptagname: string, cls: new () => XmlElement): void {
  registerClass(new NamespacePrefixedTag(nsptagname).clarkName, cls);
}

/** lxml `oxml_parser.makeelement`. */
export function makeelement(clarkName: string, nsmap?: Record<string, string> | null): any {
  return XmlElement.make(clarkName, nsmap ?? null);
}
