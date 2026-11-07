// Fundamental Open Packaging Convention (OPC) objects (`pptx.opc.package`).
import { KeyError, NotImplementedError, ValueError } from "../errors.js";
import { bytesOf, type FileSource, type FileTarget } from "../fileio.js";
import { parseXml } from "../oxml/parser.js";
import { PY_CONTAINS, PY_GETITEM, PY_LEN, XmlElement } from "../oxml/xml.js";
import { lazy } from "../util.js";
import { RELATIONSHIP_TARGET_MODE as RTM, RELATIONSHIP_TYPE as RT } from "./constants.js";
import { CT_Relationship, CT_Relationships, CT_Types, serializePartXml } from "./oxml.js";
import { CONTENT_TYPES_URI, PACKAGE_URI, PackURI } from "./packuri.js";
import { PackageReader, PackageWriter } from "./serialized.js";
import { CaseInsensitiveDict } from "./shared.js";

/** {partname: Part} as loaded from a package (keys are partname strings). */
export type PartsDict = Record<string, Part>;

/** Relationship methods required by both the package and each part. */
export class _RelatableMixin {
  /** The (single) part related to this one by `reltype`; KeyError if none, ValueError if several. */
  partRelatedBy(reltype: string): any {
    return this._rels.partWithReltype(reltype);
  }

  /** rId of the relationship of `reltype` to `target`, adding the relationship if needed. */
  relateTo(target: Part | string, reltype: string, isExternal = false): string {
    if (typeof target === "string") {
      if (!isExternal) throw new Error("AssertionError");
      return this._rels.getOrAddExtRel(reltype, target);
    }
    return this._rels.getOrAdd(reltype, target);
  }

  /** Related part identified by `rId`. */
  relatedPart(rId: string): any {
    return this._rels.get(rId).targetPart;
  }

  /** URL contained in the target ref of the relationship identified by `rId`. */
  targetRef(rId: string): string {
    return this._rels.get(rId).targetRef;
  }

  get _rels(): _Relationships {
    throw new NotImplementedError(`\`${this.constructor.name}\` must implement \`.rels\``);
  }
}

/** Main API class for the OPC layer; construct with `OpcPackage.open(pkgFile)`. */
export class OpcPackage extends _RelatableMixin {
  _pkgFile: FileSource;

  constructor(pkgFile: FileSource) {
    super();
    this._pkgFile = pkgFile;
  }

  /** A package loaded with the contents of `pkgFile` (path, bytes or file-like object). */
  static open<T extends OpcPackage>(this: new (f: FileSource) => T, pkgFile: FileSource): T {
    return new this(pkgFile)._load() as T;
  }

  /** Remove the relationship identified by `rId`. */
  dropRel(rId: string): void {
    this._rels.pop(rId);
  }

  /** Generate exactly one reference to each part in the package. */
  *iterParts(): Generator<Part> {
    const visited = new Set<Part>();
    for (const rel of this.iterRels()) {
      if (rel.isExternal) continue;
      const part = rel.targetPart;
      if (visited.has(part)) continue;
      yield part;
      visited.add(part);
    }
  }

  /** Generate exactly one reference to each relationship in the package (depth-first). */
  *iterRels(): Generator<_Relationship> {
    const visited = new Set<Part>();
    function* walkRels(rels: _Relationships): Generator<_Relationship> {
      for (const rel of rels.values()) {
        yield rel;
        if (rel.isExternal) continue;
        const part = rel.targetPart;
        if (visited.has(part)) continue;
        visited.add(part);
        yield* walkRels(part.rels);
      }
    }
    yield* walkRels(this._rels);
  }

  /** The main document part for this package (the presentation part). */
  get mainDocumentPart(): any {
    return this.partRelatedBy(RT.OFFICE_DOCUMENT);
  }

  /** Next available partname matching printf-style `tmpl`, e.g. '/ppt/slides/slide%d.xml'. */
  nextPartname(tmpl: string): PackURI {
    const fill = (n: number) => tmpl.replace("%d", String(n));
    const prefix = tmpl.slice(0, fill(42).indexOf("42"));
    const partnames = new Set<string>();
    for (const p of this.iterParts()) if (String(p.partname).startsWith(prefix)) partnames.add(String(p.partname));
    for (let n = partnames.size + 1; n > 0; n--) {
      const candidate = fill(n);
      if (!partnames.has(candidate)) return new PackURI(candidate);
    }
    throw new Error("ProgrammingError: ran out of candidate_partnames");
  }

  /** Save this package to `pkgFile` (a path or a writable file-like object). */
  save(pkgFile: FileTarget): void {
    PackageWriter.write(pkgFile, this._rels, [...this.iterParts()]);
  }

  /** The package after loading all parts and relationships. */
  _load(): this {
    const [pkgXmlRels, parts] = _PackageLoader.load(this._pkgFile, this);
    this._rels.loadFromXml(PACKAGE_URI, pkgXmlRels, parts);
    return this;
  }

  get _rels(): _Relationships {
    return lazy(this, "_rels", () => new _Relationships(PACKAGE_URI.baseURI));
  }
}

/** Function-object that loads a package from disk (or other store). */
export class _PackageLoader {
  _pkgFile: FileSource;
  _package: any;

  constructor(pkgFile: FileSource, pkg: any) {
    this._pkgFile = pkgFile;
    this._package = pkg;
  }

  /** [pkgXmlRels, parts] resulting from loading `pkgFile`. */
  static load(pkgFile: FileSource, pkg: any): [CT_Relationships, PartsDict] {
    return new this(pkgFile, pkg)._load();
  }

  _load(): [CT_Relationships, PartsDict] {
    const parts = this._parts;
    const xmlRels = this._xmlRels;
    for (const [partname, part] of Object.entries(parts)) part.loadRelsFromXml(xmlRels[partname], parts);
    return [xmlRels[String(PACKAGE_URI)], parts];
  }

  get _contentTypes(): _ContentTypeMap {
    return lazy(this, "_contentTypes", () => _ContentTypeMap.fromXml(this._packageReader.get(CONTENT_TYPES_URI)));
  }

  get _packageReader(): PackageReader {
    return lazy(this, "_packageReader", () => new PackageReader(this._pkgFile));
  }

  /** {partname: Part} for every part reachable through relationships. */
  get _parts(): PartsDict {
    return lazy(this, "_parts", () => {
      const contentTypes = this._contentTypes;
      const pkg = this._package;
      const reader = this._packageReader;
      const out: PartsDict = {};
      for (const name of Object.keys(this._xmlRels)) {
        if (name === "/") continue;
        const partname = new PackURI(name);
        // invalid partnames can arise in some packages; ignore those rather than raise
        if (!reader.has(partname)) continue;
        out[name] = PartFactory(partname, contentTypes.get(partname), pkg, reader.get(partname));
      }
      return out;
    });
  }

  /** {partname: CT_Relationships} for the package and all its parts. */
  get _xmlRels(): Record<string, CT_Relationships> {
    return lazy(this, "_xmlRels", () => {
      const xmlRels: Record<string, CT_Relationships> = {};
      const visited = new Set<string>();
      const loadRels = (sourcePartname: PackURI, rels: CT_Relationships) => {
        xmlRels[String(sourcePartname)] = rels;
        visited.add(String(sourcePartname));
        const baseUri = sourcePartname.baseURI;
        for (const rel of rels.relationshipLst) {
          if (rel.targetMode === RTM.EXTERNAL) continue;
          const target = PackURI.fromRelRef(baseUri, rel.targetRef);
          if (visited.has(String(target))) continue;
          loadRels(target, this._xmlRelsFor(target));
        }
      };
      loadRels(PACKAGE_URI, this._xmlRelsFor(PACKAGE_URI));
      return xmlRels;
    });
  }

  /** Parsed rels XML for `partname`; an empty `<Relationships>` when it has none. */
  _xmlRelsFor(partname: PackURI): CT_Relationships {
    const relsXml = this._packageReader.relsXmlFor(partname);
    return relsXml === null ? CT_Relationships.new() : (parseXml(relsXml) as CT_Relationships);
  }
}

/** Base class for package parts; also the default class for parts without specific behavior. */
export class Part extends _RelatableMixin {
  _partname: PackURI;
  _contentType: string;
  _package: any;
  _blob: Uint8Array | null;

  constructor(partname: PackURI, contentType: string, pkg: any, blob: Uint8Array | null = null) {
    super();
    this._partname = partname;
    this._contentType = contentType;
    this._package = pkg;
    this._blob = blob;
  }

  /** Instance loaded from the arguments; subclasses may pre-process (see XmlPart). */
  static load(partname: PackURI, contentType: string, pkg: any, blob: Uint8Array): Part {
    return new this(partname, contentType, pkg, blob);
  }

  /** Contents of this package part as bytes. */
  get blob(): Uint8Array {
    return this._blob ?? new Uint8Array(0);
  }
  set blob(blob: Uint8Array) {
    this._blob = blob;
  }

  /** Content-type (MIME-type) of this part. */
  get contentType(): string {
    return lazy(this, "contentType", () => this._contentType);
  }

  /** Load this part's relationships from `xmlRels`, resolving targets through `parts`. */
  loadRelsFromXml(xmlRels: CT_Relationships, parts: PartsDict): void {
    this._rels.loadFromXml(this._partname.baseURI, xmlRels, parts);
  }

  /** Package this part belongs to. */
  get package(): any {
    return lazy(this, "package", () => this._package);
  }

  /** PackURI partname for this part, e.g. "/ppt/slides/slide1.xml". */
  get partname(): PackURI {
    return this._partname;
  }
  set partname(partname: PackURI) {
    if (!(partname instanceof PackURI)) throw new TypeError(`partname must be instance of PackURI, got '${(partname as any)?.constructor?.name ?? typeof partname}'`);
    this._partname = partname;
  }

  /** Collection of relationships from this part to other parts. */
  get rels(): _Relationships {
    return lazy(this, "rels", () => this._rels);
  }

  /** Bytes of `file`: a path, bytes, or a file-like object. */
  _blobFromFile(file: FileSource): Uint8Array {
    return bytesOf(file);
  }

  get _rels(): _Relationships {
    return lazy(this, "_rels", () => new _Relationships(this._partname.baseURI));
  }
}

/** Base class for package parts containing an XML payload, which is most of them. */
export class XmlPart extends Part {
  _element: any;

  constructor(partname: PackURI, contentType: string, pkg: any, element: XmlElement) {
    super(partname, contentType, pkg);
    this._element = element;
  }

  /** Instance loaded with parsed XML from `blob`. */
  static load(partname: PackURI, contentType: string, pkg: any, blob: Uint8Array): any {
    return new (this as any)(partname, contentType, pkg, parseXml(blob));
  }

  /** XML serialization of this part. */
  get blob(): Uint8Array {
    return serializePartXml(this._element);
  }
  set blob(_b: Uint8Array) {
    throw new Error("AttributeError: can't set attribute 'blob' on an XmlPart");
  }

  /** Remove relationship `rId` if its reference count is under 2. */
  dropRel(rId: string): void {
    if (this._relRefCount(rId) < 2) this._rels.pop(rId);
  }

  /** This part (the end of the parent-delegation chain). */
  get part(): this {
    return this;
  }

  /** Count of references in this part's XML to `rId`. */
  _relRefCount(rId: string): number {
    return (this._element.xpath("//@r:id") as string[]).filter((r) => r === rId).length;
  }
}

/**
 * Constructs a registered subtype of Part for a content type (python `PartFactory(...)`).
 * Register custom part classes in `PartFactory.partTypeFor[contentType]`.
 */
export const PartFactory = Object.assign(
  function PartFactory(partname: PackURI, contentType: string, pkg: any, blob: Uint8Array): any {
    const PartClass = (PartFactory as any)._partClsFor(contentType);
    return PartClass.load(partname, contentType, pkg, blob);
  },
  {
    partTypeFor: {} as Record<string, typeof Part>,
    /** The part class registered for `contentType`, or Part. */
    _partClsFor(contentType: string): typeof Part {
      return Object.prototype.hasOwnProperty.call(PartFactory.partTypeFor, contentType) ? PartFactory.partTypeFor[contentType] : Part;
    },
  },
);

/** Dict-like lookup of content type by partname. */
export class _ContentTypeMap {
  _overrides: CaseInsensitiveDict<string> | Record<string, string>;
  _defaults: CaseInsensitiveDict<string> | Record<string, string>;

  constructor(overrides: CaseInsensitiveDict<string> | Record<string, string>, defaults: CaseInsensitiveDict<string> | Record<string, string>) {
    this._overrides = overrides;
    this._defaults = defaults;
  }

  /** Content-type (MIME-type) for the part identified by `partname`. */
  get(partname: PackURI): string {
    if (!(partname instanceof PackURI))
      throw new TypeError(`_ContentTypeMap key must be <type 'PackURI'>, got ${typeof partname === "string" ? "str" : (partname as any)?.constructor?.name}`);
    const lookup = (d: CaseInsensitiveDict<string> | Record<string, string>, k: string): string | undefined =>
      d instanceof CaseInsensitiveDict ? (d.has(k) ? d.get(k) : undefined) : Object.prototype.hasOwnProperty.call(d, k) ? d[k] : undefined;
    const o = lookup(this._overrides, String(partname));
    if (o !== undefined) return o;
    const d = lookup(this._defaults, partname.ext);
    if (d !== undefined) return d;
    throw new KeyError(`no content-type for partname '${partname}' in [Content_Types].xml`);
  }

  [PY_GETITEM](k: PackURI) {
    return this.get(k);
  }

  static fromXml(contentTypesXml: Uint8Array | string): _ContentTypeMap {
    const typesElm = parseXml(contentTypesXml) as CT_Types;
    // all partnames in [Content_Types].xml are absolute
    const overrides = new CaseInsensitiveDict<string>(typesElm.overrideLst.map((o) => [o.partName.toLowerCase(), o.contentType] as [string, string]));
    const defaults = new CaseInsensitiveDict<string>(typesElm.defaultLst.map((d) => [d.extension.toLowerCase(), d.contentType] as [string, string]));
    return new this(overrides, defaults);
  }
}

/** Collection of `_Relationship` instances with dict semantics, keyed by rId. */
export class _Relationships implements Iterable<string> {
  _baseUri: string | String;

  constructor(baseUri: string | String) {
    this._baseUri = baseUri;
  }

  /** `"rId7" in relationships` */
  has(rId: unknown): boolean {
    return typeof rId === "string" && Object.prototype.hasOwnProperty.call(this._rels, rId);
  }

  /** Relationship lookup by rId, like rels[rId]. */
  get(rId: string): _Relationship {
    if (!this.has(rId)) throw new KeyError(`no relationship with key '${rId}'`);
    return this._rels[rId];
  }

  [Symbol.iterator](): Iterator<string> {
    return Object.keys(this._rels)[Symbol.iterator]();
  }

  /** Count of relationships in the collection. */
  get length(): number {
    return Object.keys(this._rels).length;
  }

  [PY_GETITEM](k: string) {
    return this.get(k);
  }
  [PY_LEN]() {
    return this.length;
  }
  [PY_CONTAINS](k: unknown) {
    return this.has(k);
  }

  keys(): string[] {
    return Object.keys(this._rels);
  }

  values(): _Relationship[] {
    return Object.values(this._rels);
  }

  items(): [string, _Relationship][] {
    return Object.entries(this._rels);
  }

  /** rId of `reltype` to `targetPart`, adding a relationship when there isn't one. */
  getOrAdd(reltype: string, targetPart: Part): string {
    const existing = this._getMatching(reltype, targetPart);
    return existing === null ? this._addRelationship(reltype, targetPart) : existing;
  }

  /** rId of an external relationship of `reltype` to `targetRef`, adding one when needed. */
  getOrAddExtRel(reltype: string, targetRef: string): string {
    const existing = this._getMatching(reltype, targetRef, true);
    return existing === null ? this._addRelationship(reltype, targetRef, true) : existing;
  }

  /** Replace any relationships in this collection with those from `xmlRels`. */
  loadFromXml(baseUri: string | String, xmlRels: CT_Relationships, parts: PartsDict): void {
    const valid: _Relationship[] = [];
    for (const relElm of xmlRels.relationshipLst) {
      // a client sometimes "removes" a relationship by voiding its target; skip those
      if (relElm.targetMode === RTM.INTERNAL) {
        const partname = PackURI.fromRelRef(baseUri, relElm.targetRef);
        if (!(String(partname) in parts)) continue;
      }
      valid.push(_Relationship.fromXml(baseUri, relElm, parts));
    }
    const rels = this._rels;
    for (const k of Object.keys(rels)) delete rels[k];
    for (const rel of valid) rels[rel.rId] = rel;
  }

  /** Target part of the relationship with matching `reltype`. */
  partWithReltype(reltype: string): any {
    const rels = getList(this._relsByReltype, reltype);
    if (rels.length === 0) throw new KeyError(`no relationship of type '${reltype}' in collection`);
    if (rels.length > 1) throw new ValueError(`multiple relationships of type '${reltype}' in collection`);
    return rels[0].targetPart;
  }

  /** Remove and return the relationship identified by `rId`. */
  pop(rId: string): _Relationship {
    const rels = this._rels;
    if (!Object.prototype.hasOwnProperty.call(rels, rId)) throw new KeyError(rId);
    const r = rels[rId];
    delete rels[rId];
    return r;
  }

  /** XML serialization of this collection, suitable for a .rels file (rIds in numerical order). */
  get xml(): Uint8Array {
    const relsElm = CT_Relationships.new();
    const sorted = this.keys()
      .map((rId) => [rId.startsWith("rId") && /^\d+$/.test(rId.slice(3)) ? parseInt(rId.slice(3), 10) : 0, rId] as [number, string])
      .sort((a, b) => a[0] - b[0] || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0));
    for (const [, rId] of sorted) {
      const rel = this.get(rId);
      relsElm.addRel(rel.rId, rel.reltype, rel.targetRef, rel.isExternal);
    }
    return relsElm.xmlFileBytes;
  }

  _addRelationship(reltype: string, target: Part | string, isExternal = false): string {
    const rId = this._nextRId;
    this._rels[rId] = new _Relationship(this._baseUri, rId, reltype, isExternal ? RTM.EXTERNAL : RTM.INTERNAL, target);
    return rId;
  }

  _getMatching(reltype: string, target: Part | string, isExternal = false): string | null {
    for (const rel of getList(this._relsByReltype, reltype)) {
      if (rel.isExternal !== isExternal) continue;
      const relTarget = rel.isExternal ? rel.targetRef : rel.targetPart;
      if (relTarget === target) return rel.rId;
    }
    return null;
  }

  /** First unused "rIdN", making use of any gaps in numbering. */
  get _nextRId(): string {
    const rels = this._rels;
    for (let n = this.length + 1; n > 0; n--) {
      const candidate = `rId${n}`;
      if (!Object.prototype.hasOwnProperty.call(rels, candidate)) return candidate;
    }
    throw new Error("ProgrammingError: Impossible to have more distinct rIds than relationships");
  }

  /** {rId: _Relationship} */
  get _rels(): Record<string, _Relationship> {
    return lazy(this, "_rels", () => ({}) as Record<string, _Relationship>);
  }

  /** {reltype: [rels]} for all relationships in the collection. */
  get _relsByReltype(): Record<string, _Relationship[]> {
    const d: Record<string, _Relationship[]> = {};
    for (const rel of this.values()) (d[rel.reltype] ??= []).push(rel);
    return d;
  }
}

/** d[k] of a python defaultdict(list) (also accepts a plain mapping). */
function getList<T>(d: Record<string, T[]>, k: string): T[] {
  return Object.prototype.hasOwnProperty.call(d, k) ? d[k] : [];
}

/** Value object describing a link from a part or package to another part. */
export class _Relationship {
  _baseUri: string | String;
  _rId: string;
  _reltype: string;
  _targetMode: string;
  _target: Part | string;

  constructor(baseUri: string | String, rId: string, reltype: string, targetMode: string, target: Part | string) {
    this._baseUri = baseUri;
    this._rId = rId;
    this._reltype = reltype;
    this._targetMode = targetMode;
    this._target = target;
  }

  static fromXml(baseUri: string | String, rel: CT_Relationship, parts: PartsDict): _Relationship {
    const target = rel.targetMode === RTM.EXTERNAL ? rel.targetRef : parts[String(PackURI.fromRelRef(baseUri, rel.targetRef))];
    return new this(baseUri, rel.rId, rel.reltype, rel.targetMode, target);
  }

  /** True if the target is outside the package, such as a URL. */
  get isExternal(): boolean {
    return lazy(this, "isExternal", () => this._targetMode === RTM.EXTERNAL);
  }

  /** Member of RELATIONSHIP_TYPE describing the relationship of target to source. */
  get reltype(): string {
    return lazy(this, "reltype", () => this._reltype);
  }

  /** Relationship id, like 'rId9'. */
  get rId(): string {
    return lazy(this, "rId", () => this._rId);
  }

  /** Part referred to by this relationship. */
  get targetPart(): any {
    return lazy(this, "targetPart", () => {
      if (this.isExternal) throw new ValueError("`.target_part` property on _Relationship is undefined when target-mode is external");
      return this._target as Part;
    });
  }

  /** Partname targeted by this relationship. */
  get targetPartname(): PackURI {
    return lazy(this, "targetPartname", () => {
      if (this.isExternal) throw new ValueError("`.target_partname` property on _Relationship is undefined when target-mode is external");
      return (this._target as Part).partname;
    });
  }

  /** Relative partname for an internal relationship; typically a URL for an external one. */
  get targetRef(): string {
    return lazy(this, "targetRef", () => {
      if (this.isExternal) return this._target as string;
      return this.targetPartname.relativeRef(this._baseUri);
    });
  }
}
