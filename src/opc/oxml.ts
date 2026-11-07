// OPC-local oxml: content types and relationship elements (`pptx.opc.oxml`).
import { NAMESPACE as NS, RELATIONSHIP_TARGET_MODE as RTM } from "./constants.js";
import { parseXml, registerElementCls } from "../oxml/parser.js";
import { ST_ContentType, ST_Extension, ST_TargetMode, XsdAnyUri, XsdId } from "../oxml/simpletypes.js";
import { BaseOxmlElement, defineChildren, OptionalAttribute, RequiredAttribute, ZeroOrMore } from "../oxml/xmlchemy.js";
import { encodeUtf8, tostring, XmlElement } from "../oxml/xml.js";
import type { PackURI } from "./packuri.js";

export const nsmap = {
  ct: NS.OPC_CONTENT_TYPES,
  pr: NS.OPC_RELATIONSHIPS,
  r: NS.OFC_RELATIONSHIPS,
};

/** Encoded XML bytes for `element` (lxml `etree.tostring` with an encoding). */
export function oxmlToEncodedBytes(
  element: XmlElement,
  encoding = "utf-8",
  prettyPrint = false,
  standalone: boolean | null = null,
): Uint8Array {
  return encodeUtf8(oxmlTostring(element, encoding, prettyPrint, standalone));
}

/** XML text for `elm`; a declaration is included when an encoding other than "unicode" is named. */
export function oxmlTostring(
  elm: XmlElement,
  encoding: string | null = null,
  prettyPrint = false,
  standalone: boolean | null = null,
): string {
  // lxml writes a declaration when `standalone` is given or the encoding isn't utf-8/ascii
  const decl = standalone !== null || (encoding !== null && !/^(utf-?8|unicode|ascii|us-ascii)$/i.test(encoding));
  return tostring(elm, {
    xmlDeclaration: decl,
    encoding: encoding ?? "ASCII",
    standalone,
    prettyPrint,
    ascii: encoding === null,
  });
}

/** XML-file bytes for `partElm` (with XML declaration), suitable for a `.xml` package item. */
export function serializePartXml(partElm: XmlElement): Uint8Array {
  return encodeUtf8(tostring(partElm, { xmlDeclaration: true, encoding: "UTF-8", standalone: true }));
}

/** `<Default>`: the default content type for parts with the specified extension. */
export class CT_Default extends BaseOxmlElement {
  declare extension: string;
  declare contentType: string;
  static {
    defineChildren(this, {
      extension: new RequiredAttribute("Extension", ST_Extension),
      contentType: new RequiredAttribute("ContentType", ST_ContentType),
    });
  }
}

/** `<Override>`: the content type for the part with the specified partname. */
export class CT_Override extends BaseOxmlElement {
  declare partName: string;
  declare contentType: string;
  static {
    defineChildren(this, {
      partName: new RequiredAttribute("PartName", XsdAnyUri),
      contentType: new RequiredAttribute("ContentType", ST_ContentType),
    });
  }
}

/** `<Relationship>`: a single relationship from a source to a target part. */
export class CT_Relationship extends BaseOxmlElement {
  declare rId: string;
  declare reltype: string;
  declare targetRef: string;
  declare targetMode: string;
  static {
    defineChildren(this, {
      rId: new RequiredAttribute("Id", XsdId),
      reltype: new RequiredAttribute("Type", XsdAnyUri),
      targetRef: new RequiredAttribute("Target", XsdAnyUri),
      targetMode: new OptionalAttribute("TargetMode", ST_TargetMode, RTM.INTERNAL),
    });
  }

  /** A new `<Relationship>` element; `targetRef` is either a partname or a URI. */
  static new(rId: string, reltype: string, targetRef: string, targetMode: string = RTM.INTERNAL): CT_Relationship {
    const relationship = parseXml(`<Relationship xmlns="${nsmap.pr}"/>`) as CT_Relationship;
    relationship.rId = rId;
    relationship.reltype = reltype;
    relationship.targetRef = targetRef;
    relationship.targetMode = targetMode;
    return relationship;
  }
}

/** `<Relationships>`, the root element in a .rels file. */
export class CT_Relationships extends BaseOxmlElement {
  declare relationshipLst: CT_Relationship[];
  declare _insertRelationship: (r: CT_Relationship) => CT_Relationship;
  static {
    defineChildren(this, { relationship: new ZeroOrMore("pr:Relationship") });
  }

  /** Add a child `<Relationship>` element with attributes set as specified. */
  addRel(rId: string, reltype: string, target: string, isExternal = false): CT_Relationship {
    const targetMode = isExternal ? RTM.EXTERNAL : RTM.INTERNAL;
    return this._insertRelationship(CT_Relationship.new(rId, reltype, target, targetMode));
  }

  static new(): CT_Relationships {
    return parseXml(`<Relationships xmlns="${nsmap.pr}"/>`) as CT_Relationships;
  }

  /** XML bytes, with XML declaration, for this `<Relationships>` element. */
  get xmlFileBytes(): Uint8Array {
    return oxmlToEncodedBytes(this, "UTF-8", false, true);
  }
}

/** `<Types>`, the container element for Default and Override elements in [Content_Types].xml. */
export class CT_Types extends BaseOxmlElement {
  declare defaultLst: CT_Default[];
  declare overrideLst: CT_Override[];
  declare _addDefault: (attrs?: Record<string, any>) => CT_Default;
  declare _addOverride: (attrs?: Record<string, any>) => CT_Override;
  static {
    defineChildren(this, {
      default: new ZeroOrMore("ct:Default"),
      override: new ZeroOrMore("ct:Override"),
    });
  }

  addDefault(ext: string, contentType: string): CT_Default {
    return this._addDefault({ extension: ext, contentType });
  }

  addOverride(partname: PackURI | string, contentType: string): CT_Override {
    return this._addOverride({ partName: String(partname), contentType });
  }

  static new(): CT_Types {
    return parseXml(`<Types xmlns="${nsmap.ct}"/>`) as CT_Types;
  }
}

registerElementCls("ct:Default", CT_Default);
registerElementCls("ct:Override", CT_Override);
registerElementCls("ct:Types", CT_Types);
registerElementCls("pr:Relationship", CT_Relationship);
registerElementCls("pr:Relationships", CT_Relationships);
