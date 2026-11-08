// Objects related to images, audio, and video (`pptx.media`).
import { basename, bytesOf, extname, type FileSource } from "./fileio.js";
import { CONTENT_TYPE as CT } from "./opc/constants.js";
import { sha1Hex } from "./sha1.js";
import { base64Decode } from "./templates/index.js";
import { lazy } from "./util.js";

const EXT_BY_MIME: Record<string, string> = {
  [CT.ASF]: "asf",
  [CT.AVI]: "avi",
  [CT.MOV]: "mov",
  [CT.MP4]: "mp4",
  [CT.MPG]: "mpg",
  [CT.MS_VIDEO]: "avi",
  [CT.SWF]: "swf",
  [CT.WMV]: "wmv",
  [CT.X_MS_VIDEO]: "avi",
};

/** Immutable value object representing a video such as MP4. */
export class Video {
  declare _blob: Uint8Array;
  declare _mimeType: string | null;
  declare _filename: string | null;

  constructor(blob: Uint8Array, mimeType: string | null, filename: string | null) {
    this._blob = blob;
    this._mimeType = mimeType;
    this._filename = filename;
  }

  /** A new Video loaded from the binary in `blob`. */
  static fromBlob(blob: Uint8Array, mimeType: string | null, filename: string | null = null): Video {
    return new this(blob, mimeType, filename);
  }

  /** A new Video loaded from a path, bytes or a file-like object. */
  static fromPathOrFileLike(movieFile: FileSource, mimeType: string | null): Video {
    if (typeof movieFile === "string") {
      // treat movieFile as a path
      return this.fromBlob(bytesOf(movieFile), mimeType, basename(movieFile));
    }
    // assume movieFile is a file-like object
    const blob = movieFile instanceof Uint8Array || movieFile instanceof ArrayBuffer ? bytesOf(movieFile) : readAll(movieFile);
    return this.fromBlob(blob, mimeType, null);
  }

  /** The bytestream of the media "file". */
  get blob(): Uint8Array {
    return this._blob;
  }

  /** MIME-type of this media, e.g. `'video/mp4'`. */
  get contentType(): string | null {
    return this._mimeType;
  }

  /**
   * The file extension for this video, e.g. 'mp4'.
   *
   * The extension is that from the actual filename if known. Otherwise it is the lowercase
   * canonical extension for the video's MIME type. 'vid' is used if the MIME type is not
   * recognized.
   */
  get ext(): string {
    if (this._filename) return extname(this._filename).replace(/^\.+/, "");
    return (this._mimeType !== null && EXT_BY_MIME[this._mimeType]) || "vid";
  }

  /**
   * File name of this video's binary, or a default name 'movie.{ext}' when it was not loaded
   * from a file.
   */
  get filename(): string {
    if (this._filename !== null) return this._filename;
    return `movie.${this.ext}`;
  }

  /**
   * The SHA1 hash digest for the binary "file" of this video, e.g.
   * '1be010ea47803b00e140b852765cdf84f491da47'. Used to dedupe media parts.
   */
  get sha1(): string {
    return lazy(this, "sha1", () => sha1Hex(this._blob));
  }
}

/** Read a file-like object from its current position (python `f.read()`, no seek). */
function readAll(f: { read(n?: number): Uint8Array | ArrayLike<number> }): Uint8Array {
  const b = f.read();
  return b instanceof Uint8Array ? b : new Uint8Array(b as ArrayLike<number>);
}

/** The default poster frame for a movie with none: a PNG of a speaker icon. */
export const SPEAKER_IMAGE_BYTES: Uint8Array = base64Decode(
  "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",
);
