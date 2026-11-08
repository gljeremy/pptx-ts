// Embedded Package part objects. (`pptx.parts.embeddedpackage`)
import { PROG_ID } from "../enum/shapes.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { Part } from "../opc/package.js";
import type { Package } from "../package.js";

/**
 * A distinct OPC package, e.g. an Excel file, embedded in this PPTX package.
 *
 * Has a partname like: `ppt/embeddings/Microsoft_Excel_Sheet1.xlsx`.
 */
export class EmbeddedPackagePart extends Part {

  /**
   * Return a new |EmbeddedPackagePart| subclass instance added to *package*.
   *
   * The subclass is determined by `prog_id` which corresponds to the "application"
   * used to open the "file-type" of `object_blob`. The returned part contains the
   * bytes of `object_blob` and has the content-type also determined by `prog_id`.
   */
  static factory(progId: PROG_ID | string, objectBlob: Uint8Array, package_: Package): EmbeddedPackagePart {
    if (!(progId instanceof PROG_ID)) {
      return new this(package_.nextPartname("/ppt/embeddings/oleObject%d.bin"), CT.OFC_OLE_OBJECT, package_, objectBlob);
    }
    const EmbeddedPartCls = { DOCX: EmbeddedDocxPart, PPTX: EmbeddedPptxPart, XLSX: EmbeddedXlsxPart }[progId.value as "DOCX" | "PPTX" | "XLSX"];
    return EmbeddedPartCls.new(objectBlob, package_);
  }

  /**
   * Return new |EmbeddedPackagePart| subclass object.
   *
   * The returned part object contains `blob` and is added to `package`.
   */
  /** Partname template and content type are defined by each subclass. */
  static partnameTemplate: string;
  static contentType: string;

  static new(blob: Uint8Array, package_: Package): EmbeddedPackagePart {
    return new this(package_.nextPartname(this.partnameTemplate), this.contentType, package_, blob);
  }
}

/**
 * A Word .docx file stored in a part.
 *
 * This part-type arises when a Word document appears as an embedded OLE-object shape.
 */
export class EmbeddedDocxPart extends EmbeddedPackagePart {
  static partnameTemplate = "/ppt/embeddings/Microsoft_Word_Document%d.docx";
  static contentType = CT.WML_DOCUMENT;
}

/**
 * A PowerPoint file stored in a part.
 *
 * This part-type arises when a PowerPoint presentation (.pptx file) appears as an
 * embedded OLE-object shape.
 */
export class EmbeddedPptxPart extends EmbeddedPackagePart {
  static partnameTemplate = "/ppt/embeddings/Microsoft_PowerPoint_Presentation%d.pptx";
  static contentType = CT.PML_PRESENTATION;
}

/**
 * An Excel file stored in a part.
 *
 * This part-type arises as the data source for a chart, but may also be the OLE-object
 * for an embedded object shape.
 */
export class EmbeddedXlsxPart extends EmbeddedPackagePart {
  static partnameTemplate = "/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx";
  static contentType = CT.SML_SHEET;
}
