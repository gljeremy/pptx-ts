// pptx-ts tests for the zip package writer, standing in for the upstream tests that exercise
// python's zipfile.ZipFile directly (see test/fixups.json).
import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { BytesIO } from "../../src/fileio.js";
import { PackURI } from "../../src/opc/packuri.js";
import { _ZipPkgReader, _ZipPkgWriter } from "../../src/opc/serialized.js";

describe("_ZipPkgWriter", () => {
  it("writes deflated members that read back unchanged", () => {
    const stream = new BytesIO();
    const writer = new _ZipPkgWriter(stream);
    const blob = new TextEncoder().encode("<a/>".repeat(1000));
    writer.write(new PackURI("/ppt/foo.xml"), blob);
    writer.write(new PackURI("/[Content_Types].xml"), new Uint8Array([1, 2, 3]));
    writer.close();
    const files = unzipSync(stream.getvalue());
    expect(Object.keys(files).sort()).toEqual(["[Content_Types].xml", "ppt/foo.xml"]);
    expect(files["ppt/foo.xml"]).toEqual(blob);
    expect(stream.getvalue().length).toBeLessThan(blob.length / 10);
  });

  it("round trips through _ZipPkgReader", () => {
    const stream = new BytesIO();
    const writer = new _ZipPkgWriter(stream);
    writer.write(new PackURI("/a/b.bin"), new Uint8Array([9, 8, 7]));
    writer.close();
    const reader = new _ZipPkgReader(stream.getvalue());
    expect(reader.has(new PackURI("/a/b.bin"))).toBe(true);
    expect(reader.has(new PackURI("/a/c.bin"))).toBe(false);
    expect(reader.get(new PackURI("/a/b.bin"))).toEqual(new Uint8Array([9, 8, 7]));
  });
});
