// The OLE icons live outside the main entry (src/templates/icons.ts, `pptx-ts/ole-icons`); in
// Node they are read from templates/ when not registered. Both must give the files' bytes.
import { readFileSync } from "node:fs";
import { expect, it, vi } from "vitest";
import * as fileio from "../src/fileio.js";
import { base64Decode, templateBytes } from "../src/templates/index.js";
import { ICONS } from "../src/templates/icons.js";
import { BINARY } from "../src/templates/data.js";

const file = (name: string) => new Uint8Array(readFileSync(new URL(`../templates/${name}`, import.meta.url)));
const icons = ["docx-icon.emf", "generic-icon.emf", "pptx-icon.emf", "xlsx-icon.emf"];

it("keeps the OLE icons out of the main template data", () => {
  expect(Object.keys(BINARY)).toEqual(["default.pptx"]);
  expect(Object.keys(ICONS).sort()).toEqual(icons);
});

it("reads an unregistered OLE icon from the package's templates/ in Node", () => {
  for (const name of icons) expect(templateBytes(name)).toEqual(file(name));
});

it("serves the icons pptx-ts/ole-icons registers", async () => {
  const fs = vi.spyOn(fileio, "nodeFs").mockReturnValue(null);
  try {
    await import("../src/ole-icons.js");
    for (const name of icons) {
      expect(base64Decode(ICONS[name])).toEqual(file(name));
      expect(templateBytes(name)).toEqual(file(name));
    }
  } finally {
    fs.mockRestore();
  }
});

it("raises KeyError for a template file that doesn't exist", () => {
  expect(() => templateBytes("nope.emf")).toThrow(/is not loaded/);
  expect(() => templateBytes("nope.bin")).toThrow(/no template file named/);
});
