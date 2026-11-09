// `import "pptx-ts/ole-icons"`: bundles the icons python-pptx shows for embedded OLE objects
// (Excel, Word, PowerPoint and generic, about 260 KB of EMF), for `addOleObject()` in browsers.
// Node reads them from the package's files without this import.
import { ICONS } from "./templates/icons.js";
import { registerTemplateFiles } from "./templates/index.js";

registerTemplateFiles(ICONS);
