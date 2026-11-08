// Machine-translated from python-pptx tests/parts/test_media.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $isinstance, $list, initializerMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Video as Video$imported } from "../../src/media.js";
import { Package as Package$imported } from "../../src/package.js";
import { MediaPart } from "../../src/parts/media.js";
// python binds imported names at import time; patches made later don't change them
const Video = Video$imported;
const Package = Package$imported;
void $py;


const DescribeMediaPart: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_a_media_object", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, media_: any, media_part: any, package_: any;
        media_ = instanceMock(request, Video);
        _init_ = initializerMock(request, MediaPart);
        package_ = instanceMock(request, Package);
        package_.nextMediaPartname.returnValue = "media42.mp4";
        [media_.blob, media_.contentType] = $list([$b("blob-bytes"), "video/mp4"]);
        media_part = MediaPart.new(package_, media_);
        package_.nextMediaPartname.assertCalledOnceWith(media_.ext);
        _init_.assertCalledOnceWith(media_part, "media42.mp4", media_.contentType, package_, media_.blob);
        $assert($isinstance(media_part, MediaPart));
      } },
    { name: "it_knows_the_sha1_hash_of_the_media", args: [], fn: (self: any): any => {
        $assertEqual($call(MediaPart, null, null, null, $b("blobish-bytes")).sha1, "61efc464c21e54cfc1382fb5b6ef7512e141ceae");
      } },
  ],
};
pytestClass("DescribeMediaPart", DescribeMediaPart);

