// Overall .pptx package (`pptx.package`).
import { KeyError } from "./errors.js";
import type { FileSource } from "./fileio.js";
import type { Video } from "./media.js";
import { RELATIONSHIP_TYPE as RT } from "./opc/constants.js";
import { OpcPackage } from "./opc/package.js";
import { PackURI } from "./opc/packuri.js";
import { CorePropertiesPart } from "./parts/coreprops.js";
import { Image, ImagePart } from "./parts/image.js";
import { MediaPart } from "./parts/media.js";
import type { PresentationPart } from "./parts/presentation.js";
import { lazy } from "./util.js";

/** An overall .pptx package. */
export class Package extends OpcPackage {
  /**
   * Instance of `CorePropertiesPart` holding read/write Dublin Core doc properties.
   *
   * Creates a default core properties part if one is not present (not common).
   */
  get coreProperties(): CorePropertiesPart {
    return lazy(this, "coreProperties", () => {
      try {
        return this.partRelatedBy(RT.CORE_PROPERTIES);
      } catch (e) {
        if (!(e instanceof KeyError)) throw e;
        const coreProps = CorePropertiesPart.default(this);
        this.relateTo(coreProps, RT.CORE_PROPERTIES);
        return coreProps;
      }
    });
  }

  /**
   * An `ImagePart` containing the image in `imageFile`. If the image part already exists in
   * this package, it is reused, otherwise a new one is created.
   */
  getOrAddImagePart(imageFile: FileSource): ImagePart {
    return this._imageParts.getOrAddImagePart(imageFile);
  }

  /**
   * A `MediaPart` containing the media in `media`. If a media part for this media bytestream
   * ("file") is already present in this package, it is reused, otherwise a new one is created.
   */
  getOrAddMediaPart(media: Video): MediaPart {
    return this._mediaParts.getOrAddMediaPart(media);
  }

  /**
   * The next available image partname, with extension `ext`.
   *
   * Partname uses the next available sequence number, reusing gaps.
   */
  nextImagePartname(ext: string): PackURI {
    const imageIdxs = [...this.iterParts()]
      .filter((part) => part.partname.startsWith("/ppt/media/image") && part.partname.idx !== null)
      .map((part) => part.partname.idx as number)
      .sort((a, b) => a - b);
    return new PackURI(`/ppt/media/image${firstAvailableIdx(imageIdxs)}.${ext}`);
  }

  /**
   * The next available media partname, with extension `ext`.
   *
   * Partname is first available, starting at sequence number 1. Empty sequence numbers are
   * reused.
   */
  nextMediaPartname(ext: string): PackURI {
    const mediaIdxs = [...this.iterParts()]
      .filter((part) => part.partname.startsWith("/ppt/media/media"))
      .map((part) => part.partname.idx as number)
      .sort((a, b) => a - b);
    return new PackURI(`/ppt/media/media${firstAvailableIdx(mediaIdxs)}.${ext}`);
  }

  /** The presentation part (main document part) of this package. */
  get presentationPart(): PresentationPart {
    return this.mainDocumentPart;
  }

  /** `_ImageParts` object providing access to the image parts in this package. */
  get _imageParts(): _ImageParts {
    return lazy(this, "_imageParts", () => new _ImageParts(this));
  }

  /** `_MediaParts` object providing access to the media parts in this package. */
  get _mediaParts(): _MediaParts {
    return lazy(this, "_mediaParts", () => new _MediaParts(this));
  }
}

/** First sequence number not in sorted `idxs` (python's `first_available_*_idx()`). */
function firstAvailableIdx(idxs: number[]): number {
  for (const [i, usedIdx] of idxs.entries()) {
    const idx = i + 1;
    if (idx < usedIdx) return idx;
  }
  return idxs.length + 1;
}

/** Provides access to the image parts in a package. */
export class _ImageParts implements Iterable<ImagePart> {
  declare _package: Package;

  constructor(package_: Package) {
    this._package = package_;
  }

  /** Generate a reference to each `ImagePart` in the package. */
  *[Symbol.iterator](): Generator<ImagePart> {
    const imageParts = new Set<ImagePart>();
    for (const rel of this._package.iterRels()) {
      if (rel.isExternal) continue;
      if (rel.reltype !== RT.IMAGE) continue;
      const imagePart = rel.targetPart;
      if (imageParts.has(imagePart)) continue;
      imageParts.add(imagePart);
      yield imagePart;
    }
  }

  /**
   * An `ImagePart` containing the image in `imageFile` (a path, bytes or a file-like object).
   * If an image part containing this same image already exists, that instance is returned,
   * otherwise a new image part is created.
   */
  getOrAddImagePart(imageFile: FileSource): ImagePart {
    const image = Image.fromFile(imageFile);
    const imagePart = this._findBySha1(image.sha1);
    return imagePart !== null ? imagePart : ImagePart.new(this._package, image);
  }

  /**
   * The `ImagePart` in this package whose image has SHA1 hash digest `sha1`, or null when no
   * image part matches.
   */
  _findBySha1(sha1: string): ImagePart | null {
    for (const imagePart of this) {
      // a part related as an image may be a plain Part (e.g. an unrecognized format) without a sha1
      if (!("sha1" in imagePart)) continue;
      if (imagePart.sha1 === sha1) return imagePart;
    }
    return null;
  }
}

/** Provides access to the media parts in a package, looked up by the media's SHA1 hash. */
export class _MediaParts implements Iterable<MediaPart> {
  declare _package: Package;

  constructor(package_: Package) {
    this._package = package_;
  }

  /** Generate a reference to each `MediaPart` in the package. */
  *[Symbol.iterator](): Generator<MediaPart> {
    const mediaParts = new Set<MediaPart>();
    for (const rel of this._package.iterRels()) {
      if (rel.isExternal) continue;
      if (rel.reltype !== RT.MEDIA && rel.reltype !== RT.VIDEO) continue;
      const mediaPart = rel.targetPart;
      if (mediaParts.has(mediaPart)) continue;
      mediaParts.add(mediaPart);
      yield mediaPart;
    }
  }

  /**
   * A `MediaPart` containing the media in `media`. If this package already contains a media
   * part for the same bytestream, that instance is returned, otherwise one is created.
   */
  getOrAddMediaPart(media: Video): MediaPart {
    let mediaPart = this._findBySha1(media.sha1);
    if (mediaPart === null) mediaPart = MediaPart.new(this._package, media);
    return mediaPart;
  }

  /** The `MediaPart` having `sha1` hash, or null if not found. */
  _findBySha1(sha1: string): MediaPart | null {
    for (const mediaPart of this) {
      if (mediaPart.sha1 === sha1) return mediaPart;
    }
    return null;
  }
}
