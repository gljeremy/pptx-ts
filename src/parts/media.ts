// MediaPart and related objects. (`pptx.parts.media`)
import { lazy } from "../util.js";
import { sha1Hex } from "../sha1.js";
import type { Video } from "../media.js";
import type { Package } from "../package.js";
import { Part } from "../opc/package.js";

/**
 * A media part, containing an audio or video resource.
 *
 * A media part generally has a partname matching the regex
 * `ppt/media/media[1-9][0-9]*.*`.
 */
export class MediaPart extends Part {

  /**
   * Return new |MediaPart| instance containing `media`.
   *
   * `media` must be a |Media| object.
   */
  static new(package_: Package, media: Video): MediaPart {
    return new this(package_.nextMediaPartname(media.ext), media.contentType as string, package_, media.blob);
  }

  /**
   * The SHA1 hash digest for the media binary of this media part.
   *
   * Example: `'1be010ea47803b00e140b852765cdf84f491da47'`
   */
  get sha1(): string {
    return lazy(this, "sha1", () => {
      return sha1Hex(this.blob);
    });
  }
}
