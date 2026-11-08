// Machine-translated from python-pptx tests/oxml/shapes/test_picture.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $kw, $m, $mod, pytestClass } from "../../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../../src/exc.js";
import { defineChildren } from "../../../src/oxml/xmlchemy.js";
import { nsdecls as nsdecls$imported } from "../../../src/oxml/ns.js";
import { CT_Picture as CT_Picture$imported } from "../../../src/oxml/shapes/picture.js";
// python binds imported names at import time; patches made later don't change them
const nsdecls = nsdecls$imported;
const CT_Picture = CT_Picture$imported;
void $py;


const DescribeCT_Picture: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_new_pic_element", args: ["desc", "xml_desc"], parametrize: [{ names: ["desc", "xml_desc"], values: [["kittens.jpg", "kittens.jpg"], ["bits&bobs.png", "bits&amp;bobs.png"], ["img&.png", "img&amp;.png"], ["im<ag>e.png", "im&lt;ag&gt;e.png"]] }], fn: (self: any, desc: any, xml_desc: any): any => {
        let pic: any;
        pic = $m(CT_Picture, "newPic", $kw({shape_id: 9, name: "Picture 8", desc: desc, rId: "rId42", x: 1, y: 2, cx: 3, cy: 4}));
        $assertEqual(pic.xml, $mod("<p:pic %s>\n  <p:nvPicPr>\n    <p:cNvPr id=\"9\" name=\"Picture 8\" descr=\"%s\"/>\n    <p:cNvPicPr>\n      <a:picLocks noChangeAspect=\"1\"/>\n    </p:cNvPicPr>\n    <p:nvPr/>\n  </p:nvPicPr>\n  <p:blipFill>\n    <a:blip r:embed=\"rId42\"/>\n    <a:stretch>\n      <a:fillRect/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a:xfrm>\n      <a:off x=\"1\" y=\"2\"/>\n      <a:ext cx=\"3\" cy=\"4\"/>\n    </a:xfrm>\n    <a:prstGeom prst=\"rect\">\n      <a:avLst/>\n    </a:prstGeom>\n  </p:spPr>\n</p:pic>\n", [nsdecls("a", "p", "r"), xml_desc]));
      } },
    { name: "it_can_create_a_new_video_pic_element", args: [], fn: (self: any): any => {
        let pic: any;
        pic = $m(CT_Picture, "newVideoPic", $kw({shape_id: 42, shape_name: "Media 41", video_rId: "rId1", media_rId: "rId2", poster_frame_rId: "rId3", x: 1, y: 2, cx: 3, cy: 4}));
        $assertEqual(pic.xml, $mod("<p:pic %s>\n  <p:nvPicPr>\n    <p:cNvPr id=\"42\" name=\"Media 41\">\n      <a:hlinkClick r:id=\"\" action=\"ppaction://media\"/>\n    </p:cNvPr>\n    <p:cNvPicPr>\n      <a:picLocks noChangeAspect=\"1\"/>\n    </p:cNvPicPr>\n    <p:nvPr>\n      <a:videoFile r:link=\"rId1\"/>\n      <p:extLst>\n        <p:ext uri=\"{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}\">\n          <p14:media xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" r:embed=\"rId2\"/>\n        </p:ext>\n      </p:extLst>\n    </p:nvPr>\n  </p:nvPicPr>\n  <p:blipFill>\n    <a:blip r:embed=\"rId3\"/>\n    <a:stretch>\n      <a:fillRect/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a:xfrm>\n      <a:off x=\"1\" y=\"2\"/>\n      <a:ext cx=\"3\" cy=\"4\"/>\n    </a:xfrm>\n    <a:prstGeom prst=\"rect\">\n      <a:avLst/>\n    </a:prstGeom>\n  </p:spPr>\n</p:pic>\n", nsdecls("a", "p", "r")));
      } },
  ],
};
pytestClass("DescribeCT_Picture", DescribeCT_Picture);

