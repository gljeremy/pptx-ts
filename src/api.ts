// Directly exposed API functions and classes, `Presentation()` for now (`pptx.api`).
import { ValueError } from "./errors.js";
import { BytesIO, type FileSource } from "./fileio.js";
import { CONTENT_TYPE as CT } from "./opc/constants.js";
import { Package } from "./package.js";
import type { PresentationPart } from "./parts/presentation.js";
import type { Presentation as PresentationObj } from "./presentation.js";
import { templateBytes } from "./templates/index.js";

/**
 * A `Presentation` object loaded from `pptx`, where `pptx` can be a path to a `.pptx` file
 * (Node only), its bytes, or a file-like object. If `pptx` is missing or null, the built-in
 * default presentation "template" is loaded.
 */
export function Presentation(pptx: FileSource | null = null): PresentationObj {
  if (pptx === null || pptx === undefined) pptx = _defaultPptxPath();
  const presentationPart = Package.open(pptx).mainDocumentPart as PresentationPart;
  if (!_isPptxPackage(presentationPart)) {
    throw new ValueError(`file '${pptx}' is not a PowerPoint file, content type is '${presentationPart.contentType}'`);
  }
  return presentationPart.presentation;
}

/**
 * The built-in default .pptx package (python returns its path; here it's an in-memory copy of
 * `templates/default.pptx`, so it also works in browsers).
 */
export function _defaultPptxPath(): BytesIO {
  return new BytesIO(templateBytes("default.pptx"));
}

/** True if `prsPart` is a valid main document part. */
export function _isPptxPackage(prsPart: PresentationPart): boolean {
  const validContentTypes: string[] = [CT.PML_PRESENTATION_MAIN, CT.PML_PRES_MACRO_MAIN];
  return validContentTypes.includes(prsPart.contentType);
}
