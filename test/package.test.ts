// Machine-translated from python-pptx tests/test_package.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $isinstance, $iter, $kw, $list, $mod, $path, $pptx, $slice, $t1, $truthy, call, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { Video as Video$imported } from "../src/media.js";
import { RELATIONSHIP_TYPE as RT$imported } from "../src/opc/constants.js";
import { Part as Part$imported, _Relationship as _Relationship$imported } from "../src/opc/package.js";
import { PackURI as PackURI$imported } from "../src/opc/packuri.js";
import { Package as Package$imported, _ImageParts as _ImageParts$imported, _MediaParts as _MediaParts$imported } from "../src/package.js";
import { CorePropertiesPart as CorePropertiesPart$imported } from "../src/parts/coreprops.js";
import { Image as Image$imported, ImagePart as ImagePart$imported } from "../src/parts/image.js";
import { MediaPart as MediaPart$imported } from "../src/parts/media.js";
// python binds imported names at import time; patches made later don't change them
const Video = Video$imported;
const RT = RT$imported;
const Part = Part$imported;
const _Relationship = _Relationship$imported;
const PackURI = PackURI$imported;
const Package = Package$imported;
const _ImageParts = _ImageParts$imported;
const _MediaParts = _MediaParts$imported;
const CorePropertiesPart = CorePropertiesPart$imported;
const Image = Image$imported;
const ImagePart = ImagePart$imported;
const MediaPart = MediaPart$imported;
void $py;


const DescribePackage: any = {
  fixtures: {
    image_part_fixture: { args: ["image_parts_", "image_part_", "_image_parts_prop_"], fn: (self: any, image_parts_: any, image_part_: any, _image_parts_prop_: any): any => {
        let image_file: any, package$: any;
        package$ = $call(Package, null);
        image_file = "foobar.png";
        _image_parts_prop_.returnValue = image_parts_;
        image_parts_.getOrAddImagePart.returnValue = image_part_;
        return [package$, image_file, image_parts_, image_part_];
      } },
    media_part_fixture: { args: ["media_", "media_part_", "_media_parts_prop_", "media_parts_"], fn: (self: any, media_: any, media_part_: any, _media_parts_prop_: any, media_parts_: any): any => {
        let package$: any;
        package$ = $call(Package, null);
        _media_parts_prop_.returnValue = media_parts_;
        media_parts_.getOrAddMediaPart.returnValue = media_part_;
        return [package$, media_, media_part_];
      } },
    m_parts_fixture: { args: ["_MediaParts_", "media_parts_"], fn: (self: any, _MediaParts_: any, media_parts_: any): any => {
        let package$: any;
        package$ = $call(Package, null);
        _MediaParts_.returnValue = media_parts_;
        return [package$, _MediaParts_, media_parts_];
      } },
    next_fixture: { args: ["request", "iter_parts_"], params: [[[3, 4, 2], 1], [[4, 2, 1], 3], [[2, 3, 1], 4]], fn: (self: any, request: any, iter_parts_: any): any => {
        let expected_value: any, ext: any, idx: any, idxs: any, package$: any;
        [idxs, idx] = $list(request.param);
        package$ = $call(Package, null);
        package$.iterParts.returnValue = self.iImageParts(request, idxs);
        ext = "foo";
        expected_value = $mod("/ppt/media/image%d.%s", [idx, ext]);
        return [package$, ext, expected_value];
      } },
    nmp_fixture: { args: ["request", "iter_parts_"], params: [[[3, 4, 2], 1], [[4, 2, 1], 3], [[2, 3, 1], 4]], fn: (self: any, request: any, iter_parts_: any): any => {
        let expected_value: any, ext: any, idx: any, idxs: any, package$: any;
        [idxs, idx] = $list(request.param);
        package$ = $call(Package, null);
        package$.iterParts.returnValue = self.iMediaParts(request, idxs);
        ext = "foo";
        expected_value = $mod("/ppt/media/media%d.%s", [idx, ext]);
        return [package$, ext, expected_value];
      } },
    image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ImagePart);
      } },
    image_parts_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _ImageParts);
      } },
    _image_parts_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Package, "_image_parts");
      } },
    iter_parts_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Package, "iter_parts");
      } },
    media_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Video);
      } },
    media_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MediaPart);
      } },
    _MediaParts_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.package._MediaParts");
      } },
    media_parts_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _MediaParts);
      } },
    _media_parts_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Package, "_media_parts");
      } },
  },
  helpers: {
    iImageParts: (self: any, request: any, idxs: any): any => {
        const part = (idx: any): any => {
          let partname: any;
          partname = $call(PackURI, $mod("/ppt/media/image%d.png", idx));
          return $call(instanceMock, request, Part, $kw({partname: partname}));
        };
        return $iter($list(idxs).map((idx: any) => part(idx)));
      },
    iMediaParts: (self: any, request: any, idxs: any): any => {
        const part = (idx: any): any => {
          let partname: any;
          partname = $call(PackURI, $mod("/ppt/media/media%d.mp4", idx));
          return $call(instanceMock, request, Part, $kw({partname: partname}));
        };
        return $iter($list(idxs).map((idx: any) => part(idx)));
      },
  },
  tests: [
    { name: "it_provides_access_to_its_core_properties_part", args: [], fn: (self: any): any => {
        let default_pptx: any, pkg: any;
        default_pptx = $path.abspath($path.join($get($path.split($pptx.__file__), 0), "templates", "default.pptx"));
        pkg = Package.open(default_pptx);
        $assert($isinstance(pkg.coreProperties, CorePropertiesPart));
      } },
    { name: "it_can_get_or_add_an_image_part", args: ["image_part_fixture"], fn: (self: any, image_part_fixture: any): any => {
        let image_file: any, image_part: any, image_part_: any, image_parts_: any, package$: any;
        [package$, image_file, image_parts_, image_part_] = $list(image_part_fixture);
        image_part = package$.getOrAddImagePart(image_file);
        image_parts_.getOrAddImagePart.assertCalledOnceWith(image_file);
        $assert($is(image_part, image_part_));
      } },
    { name: "it_can_get_or_add_a_media_part", args: ["media_part_fixture"], fn: (self: any, media_part_fixture: any): any => {
        let media: any, media_part: any, media_part_: any, package$: any;
        [package$, media, media_part_] = $list(media_part_fixture);
        media_part = package$.getOrAddMediaPart(media);
        package$._mediaParts.getOrAddMediaPart.assertCalledOnceWith(media);
        $assert($is(media_part, media_part_));
      } },
    { name: "it_knows_the_next_available_image_partname", args: ["next_fixture"], fn: (self: any, next_fixture: any): any => {
        let expected_value: any, ext: any, package$: any, partname: any;
        [package$, ext, expected_value] = $list(next_fixture);
        partname = package$.nextImagePartname(ext);
        $assertEqual(partname, expected_value);
      } },
    { name: "it_knows_the_next_available_media_partname", args: ["nmp_fixture"], fn: (self: any, nmp_fixture: any): any => {
        let expected_value: any, ext: any, package$: any, partname: any;
        [package$, ext, expected_value] = $list(nmp_fixture);
        partname = package$.nextMediaPartname(ext);
        $assertEqual(partname, expected_value);
      } },
    { name: "it_provides_access_to_its_MediaParts_object", args: ["m_parts_fixture"], fn: (self: any, m_parts_fixture: any): any => {
        let _MediaParts_: any, media_parts: any, media_parts_: any, package$: any;
        [package$, _MediaParts_, media_parts_] = $list(m_parts_fixture);
        media_parts = package$._mediaParts;
        _MediaParts_.assertCalledOnceWith(package$);
        $assert($is(media_parts, media_parts_));
      } },
  ],
};
pytestClass("DescribePackage", DescribePackage);

const Describe_ImageParts: any = {
  fixtures: {
    find_fixture: { args: ["request", "_iter_", "image_part_"], params: [true, false], fn: (self: any, request: any, _iter_: any, image_part_: any): any => {
        let expected_value: any, image_part_is_present: any, image_parts: any, sha1: any;
        image_part_is_present = request.param;
        image_parts = $call(_ImageParts, null);
        _iter_.returnValue = $iter([image_part_]);
        sha1 = "foobar";
        if ($truthy(image_part_is_present)) {
          image_part_.sha1 = "foobar";
          expected_value = image_part_;
        } else {
          image_part_.sha1 = "barfoo";
          expected_value = null;
        }
        return [image_parts, sha1, expected_value];
      } },
    iter_fixture: { args: ["request", "package_"], fn: (self: any, request: any, package_: any): any => {
        let expected_parts: any, image_parts: any, rels: any;
        const rel = (is_external: any, reltype: any): any => {
          let part: any;
          part = instanceMock(request, Part);
          return $call(instanceMock, request, _Relationship, $kw({is_external: is_external, reltype: reltype, target_part: part}));
        };
        rels = [rel(true, RT.IMAGE), rel(false, RT.SLIDE), rel(false, RT.IMAGE)];
        package_.iterRels.returnValue = $iter([$get(rels, 0), $get(rels, 1), $get(rels, 2), $get(rels, 2)]);
        image_parts = $call(_ImageParts, package_);
        expected_parts = [$get(rels, 2).targetPart];
        return [image_parts, expected_parts];
      } },
    _find_by_sha1_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _ImageParts, "_find_by_sha1");
      } },
    Image_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.package.Image");
      } },
    image_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Image);
      } },
    ImagePart_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.package.ImagePart");
      } },
    image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ImagePart);
      } },
    _iter_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _ImageParts, "__iter__");
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_iterate_over_the_package_image_parts", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let expected_parts: any, image_parts: any;
        [image_parts, expected_parts] = $list(iter_fixture);
        $assertEqual($list(image_parts), expected_parts);
      } },
    { name: "it_can_get_a_matching_image_part", args: ["Image_", "image_", "image_part_", "_find_by_sha1_"], fn: (self: any, Image_: any, image_: any, image_part_: any, _find_by_sha1_: any): any => {
        let image_part: any, image_parts: any;
        Image_.fromFile.returnValue = image_;
        _find_by_sha1_.returnValue = image_part_;
        image_parts = $call(_ImageParts, null);
        image_part = image_parts.getOrAddImagePart("image.png");
        Image_.fromFile.assertCalledOnceWith("image.png");
        _find_by_sha1_.assertCalledOnceWith(image_parts, image_.sha1);
        $assert($is(image_part, image_part_));
      } },
    { name: "it_can_add_an_image_part", args: ["package_", "Image_", "image_", "_find_by_sha1_", "ImagePart_", "image_part_"], fn: (self: any, package_: any, Image_: any, image_: any, _find_by_sha1_: any, ImagePart_: any, image_part_: any): any => {
        let image_part: any, image_parts: any;
        Image_.fromFile.returnValue = image_;
        _find_by_sha1_.returnValue = null;
        ImagePart_.new.returnValue = image_part_;
        image_parts = $call(_ImageParts, package_);
        image_part = image_parts.getOrAddImagePart("image.png");
        Image_.fromFile.assertCalledOnceWith("image.png");
        _find_by_sha1_.assertCalledOnceWith(image_parts, image_.sha1);
        ImagePart_.new.assertCalledOnceWith(package_, image_);
        $assert($is(image_part, image_part_));
      } },
    { name: "it_can_find_an_image_part_by_sha1_hash", args: ["find_fixture"], fn: (self: any, find_fixture: any): any => {
        let expected_value: any, image_part: any, image_parts: any, sha1: any;
        [image_parts, sha1, expected_value] = $list(find_fixture);
        image_part = image_parts._findBySha1(sha1);
        $assert($is(image_part, expected_value));
      } },
    { name: "but_it_skips_unsupported_image_types", args: ["request", "_iter_"], fn: (self: any, request: any, _iter_: any): any => {
        let image_parts: any, png_part_: any, result: any, sha1: any, svg_part_: any;
        sha1 = "f00beed";
        svg_part_ = $call(instanceMock, request, Part, $kw({name: "svg_part_"}));
        png_part_ = $call(instanceMock, request, ImagePart, $kw({name: "png_part_", sha1: sha1}));
        _iter_.returnValue = $iter([svg_part_, png_part_]);
        image_parts = $call(_ImageParts, null);
        result = image_parts._findBySha1(sha1);
        $assertEqual(result, png_part_);
      } },
  ],
};
pytestClass("Describe_ImageParts", Describe_ImageParts);

const Describe_MediaParts: any = {
  fixtures: {
    find_fixture: { args: ["request", "_iter_", "media_part_"], params: [true, false], fn: (self: any, request: any, _iter_: any, media_part_: any): any => {
        let expected_value: any, media_part_is_present: any, media_parts: any, sha1: any;
        media_part_is_present = request.param;
        media_parts = $call(_MediaParts, null);
        _iter_.returnValue = $iter([media_part_]);
        sha1 = "foobar";
        if ($truthy(media_part_is_present)) {
          media_part_.sha1 = "foobar";
          expected_value = media_part_;
        } else {
          media_part_.sha1 = "barfoo";
          expected_value = null;
        }
        return [media_parts, sha1, expected_value];
      } },
    get_or_add_fixture: { args: ["request", "package_", "media_", "MediaPart_", "media_part_", "_find_by_sha1_"], params: [true, false], fn: (self: any, request: any, package_: any, media_: any, MediaPart_: any, media_part_: any, _find_by_sha1_: any): any => {
        let calls: any, media_parts: any, media_present: any, sha1: any;
        media_present = request.param;
        media_parts = $call(_MediaParts, package_);
        const $t1 = "2468";
        media_.sha1 = $t1;
        sha1 = $t1;
        calls = ($truthy(media_present) ? [] : [call(package_, media_)]);
        _find_by_sha1_.returnValue = ($truthy(media_present) ? media_part_ : null);
        MediaPart_.new.returnValue = ($truthy(media_present) ? null : media_part_);
        return [media_parts, media_, sha1, MediaPart_, calls, media_part_];
      } },
    iter_fixture: { args: ["request", "package_"], fn: (self: any, request: any, package_: any): any => {
        let expected_parts: any, media_parts: any, part_mocks: any, rels: any;
        const rel = (is_external: any, reltype: any, part: any): any => {
          return $call(instanceMock, request, _Relationship, $kw({is_external: is_external, reltype: reltype, target_part: part}));
        };
        part_mocks = [$call(instanceMock, request, Part, $kw({name: "linked-media"})), $call(instanceMock, request, Part, $kw({name: "slide"})), $call(instanceMock, request, Part, $kw({name: "embeded-media"}))];
        rels = [rel(true, RT.MEDIA, $get(part_mocks, 0)), rel(true, RT.VIDEO, $get(part_mocks, 0)), rel(false, RT.SLIDE, $get(part_mocks, 1)), rel(false, RT.MEDIA, $get(part_mocks, 2)), rel(false, RT.VIDEO, $get(part_mocks, 2))];
        package_.iterRels.returnValue = $iter(rels);
        media_parts = $call(_MediaParts, package_);
        expected_parts = [$get(part_mocks, 2)];
        return [media_parts, expected_parts];
      } },
    _find_by_sha1_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _MediaParts, "_find_by_sha1", $kw({autospec: true}));
      } },
    _iter_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _MediaParts, "__iter__");
      } },
    media_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Video);
      } },
    MediaPart_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.package.MediaPart");
      } },
    media_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MediaPart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_iterate_the_media_parts_in_the_package", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let expected_parts: any, media_parts: any;
        [media_parts, expected_parts] = $list(iter_fixture);
        $assertEqual($list(media_parts), expected_parts);
      } },
    { name: "it_can_get_or_add_a_media_part", args: ["get_or_add_fixture"], fn: (self: any, get_or_add_fixture: any): any => {
        let MediaPart_: any, calls: any, media_: any, media_part: any, media_part_: any, media_parts: any, sha1: any;
        [media_parts, media_, sha1, MediaPart_, calls] = $list($slice(get_or_add_fixture, null, 5, null));
        media_part_ = $get(get_or_add_fixture, 5);
        media_part = media_parts.getOrAddMediaPart(media_);
        media_parts._findBySha1.assertCalledOnceWith(media_parts, sha1);
        $assertEqual(MediaPart_.new.callArgsList, calls);
        $assert($is(media_part, media_part_));
      } },
    { name: "it_can_find_a_media_part_by_sha1", args: ["find_fixture"], fn: (self: any, find_fixture: any): any => {
        let expected_value: any, media_part: any, media_parts: any, sha1: any;
        [media_parts, sha1, expected_value] = $list(find_fixture);
        media_part = media_parts._findBySha1(sha1);
        $assert($is(media_part, expected_value));
      } },
  ],
};
pytestClass("Describe_MediaParts", Describe_MediaParts);

