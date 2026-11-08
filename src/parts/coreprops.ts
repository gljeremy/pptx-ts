// Core properties part, corresponds to ``/docProps/core.xml`` part in package. (`pptx.parts.coreprops`)
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { XmlPart } from "../opc/package.js";
import { PackURI } from "../opc/packuri.js";
import { CT_CoreProperties } from "../oxml/coreprops.js";
import type { Package } from "../package.js";

/**
 * Corresponds to part named `/docProps/core.xml`.
 *
 * Contains the core document properties for this document package.
 */
export class CorePropertiesPart extends XmlPart {
  declare _element: CT_CoreProperties;

  /**
   * Return default new |CorePropertiesPart| instance suitable as starting point.
   *
   * This provides a base for adding core-properties to a package that doesn't yet
   * have any.
   */
  static default(package_: Package): CorePropertiesPart {
    const coreProps = this._new(package_);
    coreProps.title = "PowerPoint Presentation";
    coreProps.lastModifiedBy = "python-pptx";
    coreProps.revision = 1;
    // the current UTC time, as a naive datetime (see src/oxml/coreprops.ts)
    coreProps.modified = new Date(Math.floor(Date.now() / 1000) * 1000);
    return coreProps;
  }

  get author(): string {
    return this._element.authorText;
  }

  set author(value: string) {
    this._element.authorText = value;
  }

  get category(): string {
    return this._element.categoryText;
  }

  set category(value: string) {
    this._element.categoryText = value;
  }

  get comments(): string {
    return this._element.commentsText;
  }

  set comments(value: string) {
    this._element.commentsText = value;
  }

  get contentStatus(): string {
    return this._element.contentStatusText;
  }

  set contentStatus(value: string) {
    this._element.contentStatusText = value;
  }

  get created(): Date | null {
    return this._element.createdDatetime;
  }

  set created(value: Date) {
    this._element.createdDatetime = value;
  }

  get identifier(): string {
    return this._element.identifierText;
  }

  set identifier(value: string) {
    this._element.identifierText = value;
  }

  get keywords(): string {
    return this._element.keywordsText;
  }

  set keywords(value: string) {
    this._element.keywordsText = value;
  }

  get language(): string {
    return this._element.languageText;
  }

  set language(value: string) {
    this._element.languageText = value;
  }

  get lastModifiedBy(): string {
    return this._element.lastModifiedByText;
  }

  set lastModifiedBy(value: string) {
    this._element.lastModifiedByText = value;
  }

  get lastPrinted(): Date | null {
    return this._element.lastPrintedDatetime;
  }

  set lastPrinted(value: Date) {
    this._element.lastPrintedDatetime = value;
  }

  get modified(): Date | null {
    return this._element.modifiedDatetime;
  }

  set modified(value: Date) {
    this._element.modifiedDatetime = value;
  }

  get revision(): number {
    return this._element.revisionNumber;
  }

  set revision(value: number) {
    this._element.revisionNumber = value;
  }

  get subject(): string {
    return this._element.subjectText;
  }

  set subject(value: string) {
    this._element.subjectText = value;
  }

  get title(): string {
    return this._element.titleText;
  }

  set title(value: string) {
    this._element.titleText = value;
  }

  get version(): string {
    return this._element.versionText;
  }

  set version(value: string) {
    this._element.versionText = value;
  }

  /** Return new empty |CorePropertiesPart| instance. */
  static _new(package_: Package): CorePropertiesPart {
    return new CorePropertiesPart(new PackURI("/docProps/core.xml"), CT.OPC_CORE_PROPERTIES, package_, CT_CoreProperties.newCoreProperties());
  }
}
