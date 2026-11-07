// Python built-in exception types that python-pptx raises, so callers (and the
// ported tests) can tell them apart. JS's own TypeError stands in for Python's.

export class ValueError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "ValueError";
  }
}

export class KeyError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "KeyError";
  }
}

export class IndexError extends RangeError {
  constructor(message = "") {
    super(message);
    this.name = "IndexError";
  }
}

export class AttributeError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "AttributeError";
  }
}

export class NotImplementedError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "NotImplementedError";
  }
}

export class ZeroDivisionError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "ZeroDivisionError";
  }
}

export class FileNotFoundError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "FileNotFoundError";
  }
}
