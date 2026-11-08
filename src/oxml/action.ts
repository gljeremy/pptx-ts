// lxml custom element classes for text-related XML elements. (`pptx.oxml.action`)
import { defineChildren } from "./xmlchemy.js";
import { XsdString } from "./simpletypes.js";
import { BaseOxmlElement, OptionalAttribute } from "./xmlchemy.js";

/** Custom element class for <a:hlinkClick> elements. */
export class CT_Hyperlink extends BaseOxmlElement {
  declare rId: string;
  declare action: string | null;
  static {
    defineChildren(this, {
      rId: new OptionalAttribute("r:id", XsdString),
      action: new OptionalAttribute("action", XsdString),
    });
  }

  /**
   * Query portion of the `ppaction://` URL as dict.
   *
   * For example `{'id':'0', 'return':'true'}` in 'ppaction://customshow?id=0&return=true'.
   *
   * Returns an empty dict if the URL contains no query string or if no action attribute is
   * present.
   */
  get actionFields(): Record<string, string> {
    const url = this.action;
    if (url === null) {
      return {};
    }
    const halves = url.split("?");
    if (halves.length === 1) {
      return {};
    }
    const keyValuePairs = halves[1].split("&");
    return Object.fromEntries(keyValuePairs.map((pair) => pair.split("=")));
  }

  /**
   * The host portion of the `ppaction://` URL contained in the action attribute.
   *
   * For example 'customshow' in 'ppaction://customshow?id=0&return=true'. Returns |None| if no
   * action attribute is present.
   */
  get actionVerb(): string | null {
    const url = this.action;
    if (url === null) {
      return null;
    }
    const protocolAndHost = url.split("?")[0];
    const host = protocolAndHost.slice(11);
    return host;
  }
}
