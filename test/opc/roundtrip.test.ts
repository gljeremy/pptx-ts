// pptx-ts end-to-end check of the OPC layer: open real packages, save them, reopen the
// result, and compare every part (partname, content type, bytes) and relationship.
import { describe, expect, it } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { BytesIO } from "../../src/fileio.js";
import { OpcPackage } from "../../src/opc/package.js";
import { templateBytes } from "../../src/templates/index.js";

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "test_files");

function snapshot(pkg: OpcPackage) {
  const parts = [...pkg.iterParts()].map((p) => ({
    partname: String(p.partname),
    contentType: p.contentType,
    blob: Buffer.from(p.blob).toString("base64"),
    rels: p.rels.items().map(([rId, r]) => [rId, r.reltype, r.targetRef, r.isExternal]).sort((x, y) => (x[0] < y[0] ? -1 : 1)),
  }));
  return parts.sort((a, b) => (a.partname < b.partname ? -1 : 1));
}

describe("OpcPackage round trip", () => {
  const cases: [string, () => Uint8Array][] = [
    ["default.pptx template", () => templateBytes("default.pptx")],
    ...fs
      .readdirSync(dir)
      .filter((f) => f.endsWith(".pptx"))
      .map((f) => [f, () => new Uint8Array(fs.readFileSync(path.join(dir, f)))] as [string, () => Uint8Array]),
  ];
  for (const [name, load] of cases) {
    it(name, () => {
      const pkg = OpcPackage.open(load());
      const before = snapshot(pkg);
      expect(before.length).toBeGreaterThan(0);
      const out = new BytesIO();
      pkg.save(out);
      const again = OpcPackage.open(out.getvalue());
      expect(snapshot(again)).toEqual(before);
    });
  }

  it("opens from a path and from a file-like object", () => {
    const p = path.join(dir, "test.pptx");
    const a = snapshot(OpcPackage.open(p));
    const b = snapshot(OpcPackage.open(new BytesIO(new Uint8Array(fs.readFileSync(p)))));
    expect(b).toEqual(a);
  });

  it("opens an expanded (unzipped) package directory", () => {
    const pkg = OpcPackage.open(path.join(dir, "expanded_pptx"));
    expect(String(pkg.mainDocumentPart.partname)).toBe("/ppt/presentation.xml");
  });
});
