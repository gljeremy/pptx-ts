/** Exceptions used with pptx-ts (`pptx.exc`). The base exception class is PythonPptxError. */

export class PythonPptxError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "PythonPptxError";
  }
}

/** Raised when a package cannot be found at the specified path. */
export class PackageNotFoundError extends PythonPptxError {
  constructor(message = "") {
    super(message);
    this.name = "PackageNotFoundError";
  }
}

/** Raised when a value is encountered in the XML that is not valid according to the schema. */
export class InvalidXmlError extends PythonPptxError {
  constructor(message = "") {
    super(message);
    this.name = "InvalidXmlError";
  }
}
