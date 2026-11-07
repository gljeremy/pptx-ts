// Mappings that embody aspects of the Open XML spec ISO/IEC 29500 (`pptx.opc.spec`).
import { CONTENT_TYPE as CT } from "./constants.js";

export const defaultContentTypes: ReadonlyArray<readonly [string, string]> = [
  ["bin", CT.PML_PRINTER_SETTINGS],
  ["bin", CT.SML_PRINTER_SETTINGS],
  ["bin", CT.WML_PRINTER_SETTINGS],
  ["bmp", CT.BMP],
  ["emf", CT.X_EMF],
  ["fntdata", CT.X_FONTDATA],
  ["gif", CT.GIF],
  ["jpe", CT.JPEG],
  ["jpeg", CT.JPEG],
  ["jpg", CT.JPEG],
  ["mov", CT.MOV],
  ["mp4", CT.MP4],
  ["mpg", CT.MPG],
  ["png", CT.PNG],
  ["rels", CT.OPC_RELATIONSHIPS],
  ["tif", CT.TIFF],
  ["tiff", CT.TIFF],
  ["vid", CT.VIDEO],
  ["wdp", CT.MS_PHOTO],
  ["wmf", CT.X_WMF],
  ["wmv", CT.WMV],
  ["xlsx", CT.SML_SHEET],
  ["xml", CT.XML],
];

const defaultSet = new Set(defaultContentTypes.map(([e, c]) => `${e}\0${c}`));

/** True when (`ext`, `contentType`) is one of the default content types. */
export function isDefaultContentType(ext: string, contentType: string): boolean {
  return defaultSet.has(`${ext}\0${contentType}`);
}

export const imageContentTypes: Record<string, string> = {
  bmp: CT.BMP,
  emf: CT.X_EMF,
  gif: CT.GIF,
  jpe: CT.JPEG,
  jpeg: CT.JPEG,
  jpg: CT.JPEG,
  png: CT.PNG,
  tif: CT.TIFF,
  tiff: CT.TIFF,
  wdp: CT.MS_PHOTO,
  wmf: CT.X_WMF,
};
