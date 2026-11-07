// Reading/writing serialized OPC packages (`pptx.opc.serialized`).
import { unzipSync, zipSync, type Zippable } from "fflate";
import { KeyError, NotImplementedError } from "../errors.js";
import { PackageNotFoundError } from "../exc.js";
import { bytesOf, isDir, isZipBytes, nodeFs, pathExists, readFileBytes, writeFileBytes, type FileSource, type FileTarget } from "../fileio.js";
import { PY_CONTAINS, PY_GETITEM } from "../oxml/xml.js";
import { CONTENT_TYPE as CT } from "./constants.js";
import { CT_Types, serializePartXml } from "./oxml.js";
import { CONTENT_TYPES_URI, PACKAGE_URI, PackURI, posixpath } from "./packuri.js";
import { CaseInsensitiveDict } from "./shared.js";
import { isDefaultContentType } from "./spec.js";
import { lazy, pyBool } from "../util.js";
import type { Part, _Relationships } from "./package.js";

/** Access to the package-parts of an OPC package with dict semantics. */
export class PackageReader {
  _pkgFile: FileSource;

  constructor(pkgFile: FileSource) {
    this._pkgFile = pkgFile;
  }

  /** True when the part identified by `packUri` is present in the package. */
  has(packUri: unknown): boolean {
    return this._blobReader.has(packUri);
  }

  /** Bytes for the part corresponding to `packUri`. */
  get(packUri: PackURI | string): Uint8Array {
    return this._blobReader.get(packUri);
  }

  [PY_GETITEM](k: PackURI) {
    return this.get(k);
  }
  [PY_CONTAINS](k: unknown) {
    return this.has(k);
  }

  /** Optional rels item XML for `partname`; null if none is present. */
  relsXmlFor(partname: PackURI): Uint8Array | null {
    const blobReader = this._blobReader;
    const uri = partname.relsUri;
    return blobReader.has(uri) ? blobReader.get(uri) : null;
  }

  get _blobReader(): _PhysPkgReader {
    return lazy(this, "_blobReader", () => _PhysPkgReader.factory(this._pkgFile));
  }
}

/** Writes a zip-format OPC package to `pkgFile`. */
export class PackageWriter {
  _pkgFile: FileTarget;
  _pkgRels: _Relationships;
  _parts: readonly Part[];

  constructor(pkgFile: FileTarget, pkgRels: _Relationships, parts: readonly Part[]) {
    this._pkgFile = pkgFile;
    this._pkgRels = pkgRels;
    this._parts = parts;
  }

  /** Write a physical package (.pptx file) to `pkgFile`. */
  static write(pkgFile: FileTarget, pkgRels: _Relationships, parts: readonly Part[]): void {
    new this(pkgFile, pkgRels, parts)._write();
  }

  _write(): void {
    const physWriter = _PhysPkgWriter.factory(this._pkgFile);
    try {
      this._writeContentTypesStream(physWriter);
      this._writePkgRels(physWriter);
      this._writeParts(physWriter);
    } finally {
      physWriter.close();
    }
  }

  _writeContentTypesStream(physWriter: _PhysPkgWriter): void {
    physWriter.write(CONTENT_TYPES_URI, serializePartXml(_ContentTypesItem.xmlFor(this._parts)));
  }

  _writeParts(physWriter: _PhysPkgWriter): void {
    for (const part of this._parts) {
      physWriter.write(part.partname, part.blob);
      if (pyBool(part._rels)) physWriter.write(part.partname.relsUri, part.rels.xml);
    }
  }

  _writePkgRels(physWriter: _PhysPkgWriter): void {
    physWriter.write(PACKAGE_URI.relsUri, this._pkgRels.xml);
  }
}

/** Base class for physical package reader objects. */
export class _PhysPkgReader {
  has(_item: unknown): boolean {
    throw new NotImplementedError(`\`${this.constructor.name}\` must implement \`.__contains__()\``);
  }

  get(_packUri: PackURI | string): Uint8Array {
    throw new NotImplementedError(`\`${this.constructor.name}\` must implement \`.__contains__()\``);
  }

  [PY_GETITEM](k: PackURI) {
    return this.get(k);
  }
  [PY_CONTAINS](k: unknown) {
    return this.has(k);
  }

  /** A reader appropriate for `pkgFile`: a directory, a zip path, zip bytes or a file-like object. */
  static factory(pkgFile: FileSource): _PhysPkgReader {
    if (typeof pkgFile !== "string") return new _ZipPkgReader(pkgFile);
    if (nodeFs() && isDir(pkgFile)) return new _DirPkgReader(pkgFile);
    if (nodeFs() && pathExists(pkgFile) && isZipBytes(readFileBytes(pkgFile))) return new _ZipPkgReader(pkgFile);
    throw new PackageNotFoundError(`Package not found at '${pkgFile}'`);
  }
}

/** Reader for an OPC package extracted into a directory. */
export class _DirPkgReader extends _PhysPkgReader {
  _path: string;

  constructor(path: string) {
    super();
    this._path = path.replace(/[\\/]+$/, "");
  }

  has(packUri: unknown): boolean {
    if (!(packUri instanceof PackURI)) return false;
    return pathExists(posixpath.join(this._path, packUri.membername));
  }

  get(packUri: PackURI | string): Uint8Array {
    const path = posixpath.join(this._path, new PackURI(packUri).membername);
    try {
      return readFileBytes(path);
    } catch {
      throw new KeyError(`no member '${packUri}' in package`);
    }
  }
}

/** Reader for a zip-file OPC package. */
export class _ZipPkgReader extends _PhysPkgReader {
  _pkgFile: FileSource;

  constructor(pkgFile: FileSource) {
    super();
    this._pkgFile = pkgFile;
  }

  has(packUri: unknown): boolean {
    return (typeof packUri === "string" || packUri instanceof String) && this._blobs.has(String(packUri));
  }

  get(packUri: PackURI | string): Uint8Array {
    const b = this._blobs.get(String(packUri));
    if (b === undefined) throw new KeyError(`no member '${packUri}' in package`);
    return b;
  }

  /** partname -> package part binaries. */
  get _blobs(): Map<string, Uint8Array> {
    return lazy(this, "_blobs", () => {
      const bytes = bytesOf(this._pkgFile);
      let files: Record<string, Uint8Array>;
      try {
        files = unzipSync(bytes);
      } catch (e) {
        throw new Error(`BadZipFile: File is not a zip file (${(e as Error).message})`);
      }
      const out = new Map<string, Uint8Array>();
      for (const [name, data] of Object.entries(files)) {
        if (name.endsWith("/")) continue; // directory entries
        out.set(`/${name}`, data);
      }
      return out;
    });
  }
}

/** Base class for physical package writer objects. */
export class _PhysPkgWriter {
  static factory(pkgFile: FileTarget): _ZipPkgWriter {
    return new _ZipPkgWriter(pkgFile);
  }

  write(_packUri: PackURI, _blob: Uint8Array): void {
    throw new NotImplementedError(`\`${this.constructor.name}\` must implement \`.write()\``);
  }

  close(): void {}
}

/** Writer for a zip-file (.pptx file) OPC package. */
export class _ZipPkgWriter extends _PhysPkgWriter {
  _pkgFile: FileTarget;
  _members: Zippable = {};

  constructor(pkgFile: FileTarget) {
    super();
    this._pkgFile = pkgFile;
  }

  /** Write `blob` to the package with membername corresponding to `packUri`. */
  write(packUri: PackURI, blob: Uint8Array): void {
    this._members[packUri.membername] = [blob, { level: 6 }];
  }

  /** Write the zip archive out. */
  close(): void {
    const zipped = zipSync(this._members, { mtime: new Date() });
    if (typeof this._pkgFile === "string") writeFileBytes(this._pkgFile, zipped);
    else this._pkgFile.write(zipped);
  }
}

/** Composes the content-types "part" ([Content_Types].xml) for a collection of parts. */
export class _ContentTypesItem {
  _parts: readonly Part[];

  constructor(parts: readonly Part[]) {
    this._parts = parts;
  }

  /** Content-types XML mapping each part in `parts` to a content-type. */
  static xmlFor(parts: readonly Part[]): CT_Types {
    return new this(parts)._xml;
  }

  /** Defaults sorted by extension and overrides sorted by partname. */
  get _xml(): CT_Types {
    return lazy(this, "_xml", () => {
      const [defaults, overrides] = this._defaultsAndOverrides;
      const typesElm = CT_Types.new();
      for (const [ext, ct] of defaults.items().sort(cmpFirst)) typesElm.addDefault(ext, ct);
      for (const [partname, ct] of Object.entries(overrides).sort(cmpFirst)) typesElm.addOverride(partname, ct);
      return typesElm;
    });
  }

  /** [defaults {ext: content_type}, overrides {partname: content_type}] */
  get _defaultsAndOverrides(): [CaseInsensitiveDict<string>, Record<string, string>] {
    return lazy(this, "_defaultsAndOverrides", () => {
      const defaults = new CaseInsensitiveDict<string>({ rels: CT.OPC_RELATIONSHIPS, xml: CT.XML });
      const overrides: Record<string, string> = {};
      for (const part of this._parts) {
        const partname = part.partname;
        const contentType = part.contentType;
        const ext = partname.ext;
        if (isDefaultContentType(ext.toLowerCase(), contentType)) defaults.set(ext, contentType);
        else overrides[String(partname)] = contentType;
      }
      return [defaults, overrides];
    });
  }
}

const cmpFirst = (a: [string, string], b: [string, string]) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0);
