// ImagePart and related objects (`pptx.parts.image`).
import { ValueError } from "../errors.js";
import { basename, bytesOf, type FileSource } from "../fileio.js";
import { readImageProps, type ImageProps, type RawDpi } from "../image/headers.js";
import { Part } from "../opc/package.js";
import type { PackURI } from "../opc/packuri.js";
import { imageContentTypes } from "../opc/spec.js";
import type { Package } from "../package.js";
import { sha1Hex } from "../sha1.js";
import { Emu, lazy, pyRound } from "../util.js";

/**
 * An image part.
 *
 * An image part generally has a partname matching the regex `ppt/media/image[1-9][0-9]*.*`.
 */
export class ImagePart extends Part {
  declare _filename: string | null;

  constructor(partname: PackURI, contentType: string, package_: Package, blob: Uint8Array, filename: string | null = null) {
    super(partname, contentType, package_, blob);
    this._blob = blob;
    this._filename = filename;
  }

  /** A new ImagePart containing `image`, an `Image` object. */
  static new(package_: Package, image: Image): ImagePart {
    return new this(package_.nextImagePartname(image.ext), image.contentType, package_, image.blob, image.filename);
  }

  /**
   * The filename associated with this image.
   *
   * Either the filename of the original image or a generic name of the form `image.ext` where
   * `ext` is appropriate to the image file format, e.g. `'jpg'`. An image created using a path
   * will have that filename; one created with a file-like object will have a generic name.
   */
  get desc(): string {
    // generic filename if original filename is unknown
    if (this._filename === null || this._filename === undefined) return `image.${this.ext}`;
    return this._filename;
  }

  /** File-name extension for this image e.g. `'png'`. */
  get ext(): string {
    return this.partname.ext;
  }

  /** An `Image` object containing the image in this image part (not a PIL image). */
  get image(): Image {
    return new Image(this._blob!, this.desc);
  }

  /**
   * Scaled image dimensions in EMU based on the combination of parameters supplied.
   *
   * If `scaledCx` and `scaledCy` are both null, the native image size is returned. If neither is
   * null, their values are returned unchanged. If only one is given, the other is calculated so
   * the image's aspect ratio is preserved.
   */
  scale(scaledCx: number | null, scaledCy: number | null): [number, number] {
    const [imageCx, imageCy] = this._nativeSize;
    if (scaledCx && scaledCy) return [scaledCx, scaledCy];
    if (scaledCx && !scaledCy) {
      const scalingFactor = scaledCx / imageCx;
      return [scaledCx, Math.trunc(pyRound(imageCy * scalingFactor))];
    }
    if (!scaledCx && scaledCy) {
      const scalingFactor = scaledCy / imageCy;
      return [Math.trunc(pyRound(imageCx * scalingFactor)), scaledCy];
    }
    // only remaining case is both `scaledCx` and `scaledCy` are null
    return [imageCx, imageCy];
  }

  /** The 40-character SHA1 hash digest for the image binary of this image part. */
  get sha1(): string {
    return lazy(this, "sha1", () => sha1Hex(this._blob!));
  }

  /** [horzDpi, vertDpi] resolution of this image. */
  get _dpi(): [number, number] {
    return Image.fromBlob(this._blob!).dpi;
  }

  /**
   * [width, height] native dimensions of the image in EMU, calculated from the image DPI value
   * (72 dpi when the image doesn't say).
   */
  get _nativeSize(): [number, number] {
    const EMU_PER_INCH = 914400;
    const [horzDpi, vertDpi] = this._dpi;
    const [widthPx, heightPx] = this._pxSize;
    return [Emu(Math.trunc((EMU_PER_INCH * widthPx) / horzDpi)), Emu(Math.trunc((EMU_PER_INCH * heightPx) / vertDpi))];
  }

  /** [width, height] dimensions of this image in pixels. */
  get _pxSize(): [number, number] {
    return Image.fromBlob(this._blob!).size;
  }
}

const EXT_MAP: Record<string, string> = { BMP: "bmp", GIF: "gif", JPEG: "jpg", PNG: "png", TIFF: "tiff", WMF: "wmf" };

/** Immutable value object representing an image such as a JPEG, PNG, or GIF. */
export class Image {
  declare _blob: Uint8Array;
  declare _filename: string | null;

  constructor(blob: Uint8Array, filename: string | null) {
    this._blob = blob;
    this._filename = filename;
  }

  /** A new Image loaded from the image binary in `blob`. */
  static fromBlob(blob: Uint8Array, filename: string | null = null): Image {
    return new this(blob, filename);
  }

  /** A new Image loaded from `imageFile`: a path, or bytes or a file-like object. */
  static fromFile(imageFile: FileSource): Image {
    if (typeof imageFile === "string") {
      // treat imageFile as a path
      return this.fromBlob(bytesOf(imageFile), basename(imageFile));
    }
    // a file-like object: reposition the file cursor if it has one
    const f = imageFile as any;
    if (f && typeof f.seek === "function") f.seek(0);
    return this.fromBlob(bytesOf(imageFile), null);
  }

  /** The binary image bytestream of this image. */
  get blob(): Uint8Array {
    return this._blob;
  }

  /** MIME-type of this image, e.g. `"image/jpeg"`. */
  get contentType(): string {
    return lazy(this, "contentType", () => imageContentTypes[this.ext]);
  }

  /** [horzDpi, vertDpi] dots-per-inch resolution of this image; (72, 72) when not specified. */
  get dpi(): [number, number] {
    return lazy(this, "dpi", () => {
      // 72 if `dpi` is null, non-numeric, less than 1 or greater than 2048
      const intDpi = (dpi: unknown): number => {
        const f = typeof dpi === "number" ? dpi : typeof dpi === "string" && dpi.trim() !== "" ? Number(dpi) : NaN;
        if (!Number.isFinite(f)) return 72;
        const n = Math.trunc(pyRound(f));
        return n < 1 || n > 2048 ? 72 : n;
      };
      const rawDpi = this._pilProps[2] as RawDpi | unknown;
      return Array.isArray(rawDpi) ? [intDpi(rawDpi[0]), intDpi(rawDpi[1])] : [72, 72];
    });
  }

  /**
   * Canonical file extension for this image e.g. `'png'`: lowercase, for the content type of
   * the image regardless of the extension used in its filename.
   */
  get ext(): string {
    return lazy(this, "ext", () => {
      const format = this._format;
      if (format === null || !(format in EXT_MAP))
        throw new ValueError(`unsupported image format, expected one of: dict_keys(['BMP', 'GIF', 'JPEG', 'PNG', 'TIFF', 'WMF']), got '${format}'`);
      return EXT_MAP[format];
    });
  }

  /** Filename from the path used to load this image; null when loaded from bytes or a stream. */
  get filename(): string | null {
    return this._filename;
  }

  /** SHA1 hash digest of the image blob. */
  get sha1(): string {
    return lazy(this, "sha1", () => sha1Hex(this._blob));
  }

  /** [width, height] dimensions of this image in pixels. */
  get size(): [number, number] {
    return lazy(this, "size", () => this._pilProps[1]);
  }

  /** The image format, as Pillow names it, e.g. 'PNG'. */
  get _format(): string | null {
    return this._pilProps[0];
  }

  /** [format, size, dpi] read from the image headers (python-pptx gets these from Pillow). */
  get _pilProps(): ImageProps {
    return lazy(this, "_pilProps", () => readImageProps(this._blob));
  }
}
