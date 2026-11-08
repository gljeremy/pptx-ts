// Machine-translated from python-pptx tests/parts/test_image.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $is, $isinstance, $kw, $list, $open, $read, BytesIO, classMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Package as Package$imported } from "../../src/package.js";
import { Image, ImagePart } from "../../src/parts/image.js";
import { Emu as Emu$imported } from "../../src/util.js";
import { absjoin, testFileDir } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const Package = Package$imported;
const Emu = Emu$imported;
void $py;


let images_pptx_path = absjoin(testFileDir, "with_images.pptx");

let test_image_path = absjoin(testFileDir, "python-icon.jpeg");

let test_eps_path = absjoin(testFileDir, "cdw-logo.eps");

let new_image_path = absjoin(testFileDir, "monty-truth.png");

const DescribeImagePart: any = {
  fixtures: {
    image_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Image);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_an_image_object", args: ["request", "image_"], fn: (self: any, request: any, image_: any): any => {
        let _init_: any, image_part: any, package_: any, partname_: any;
        package_ = instanceMock(request, Package);
        _init_ = initializerMock(request, ImagePart);
        partname_ = package_.nextImagePartname.returnValue;
        image_part = ImagePart.new(package_, image_);
        package_.nextImagePartname.assertCalledOnceWith(image_.ext);
        _init_.assertCalledOnceWith(image_part, partname_, image_.contentType, package_, image_.blob, image_.filename);
        $assert($isinstance(image_part, ImagePart));
      } },
    { name: "it_provides_access_to_its_image", args: ["request", "image_"], fn: (self: any, request: any, image_: any): any => {
        let Image_: any, image: any, image_part: any;
        Image_ = classMock(request, "pptx.parts.image.Image");
        Image_.returnValue = image_;
        $call(propertyMock, request, ImagePart, "desc", $kw({return_value: "foobar.png"}));
        image_part = $call(ImagePart, null, null, null, $b("blob"), null);
        image = image_part.image;
        Image_.assertCalledOnceWith($b("blob"), "foobar.png");
        $assert($is(image, image_));
      } },
    { name: "it_can_scale_its_dimensions", args: ["width", "height", "expected_width", "expected_height"], parametrize: [{ names: ["width", "height", "expected_width", "expected_height"], values: [[null, null, $call(Emu, 2590800), $call(Emu, 2590800)], [1000, null, 1000, 1000], [null, 3000, 3000, 3000], [3337, 9999, 3337, 9999]] }], fn: (self: any, width: any, height: any, expected_width: any, expected_height: any): any => {
        let blob: any, f: any, image_part: any;
        f = $open(test_image_path, "rb");
        blob = $read(f);
        image_part = $call(ImagePart, null, null, null, blob);
        $assertEqual(image_part.scale(width, height), [expected_width, expected_height]);
      } },
    { name: "it_knows_its_pixel_dimensions_to_help", args: [], fn: (self: any): any => {
        let blob: any, f: any, image_part: any;
        f = $open(test_image_path, "rb");
        blob = $read(f);
        image_part = $call(ImagePart, null, null, null, blob);
        $assertEqual(image_part._pxSize, [204, 204]);
      } },
  ],
};
pytestClass("DescribeImagePart", DescribeImagePart);

const DescribeImage: any = {
  fixtures: {
    blob_fixture: { args: [], fn: (self: any): any => {
        let blob: any, image: any;
        blob = $b("foobar");
        image = $call(Image, blob, null);
        return [image, blob];
      } },
    content_type_fixture: { args: ["request", "_format_"], params: [["BMP", "image/bmp"], ["GIF", "image/gif"], ["JPEG", "image/jpeg"], ["PNG", "image/png"], ["TIFF", "image/tiff"], ["WMF", "image/x-wmf"]], fn: (self: any, request: any, _format_: any): any => {
        let expected_value: any, format: any, image: any;
        [format, expected_value] = $list(request.param);
        image = $call(Image, null, null);
        _format_.returnValue = format;
        return [image, expected_value];
      } },
    dpi_fixture: { args: ["request", "_pil_props_"], params: [[[42, 24], [42, 24]], [[42.1, 23.6], [42, 24]], [null, [72, 72]], [[3047, 2388], [72, 72]], ["foobar", [72, 72]]], fn: (self: any, request: any, _pil_props_: any): any => {
        let expected_dpi: any, image: any, raw_dpi: any;
        [raw_dpi, expected_dpi] = $list(request.param);
        image = $call(Image, null, null);
        _pil_props_.returnValue = [null, null, raw_dpi];
        return [image, expected_dpi];
      } },
    ext_fixture: { args: ["request", "_format_"], params: [["BMP", "bmp"], ["GIF", "gif"], ["JPEG", "jpg"], ["PNG", "png"], ["TIFF", "tiff"], ["WMF", "wmf"]], fn: (self: any, request: any, _format_: any): any => {
        let expected_value: any, format: any, image: any;
        [format, expected_value] = $list(request.param);
        image = $call(Image, null, null);
        _format_.returnValue = format;
        return [image, expected_value];
      } },
    filename_fixture: { args: ["request"], params: ["foo.bar", null], fn: (self: any, request: any): any => {
        let filename: any, image: any;
        filename = request.param;
        image = $call(Image, null, filename);
        return [image, filename];
      } },
    from_stream_fixture: { args: ["from_blob_", "image_"], fn: (self: any, from_blob_: any, image_: any): any => {
        let blob: any, f: any, image_file: any;
        f = $open(test_image_path, "rb");
        blob = $read(f);
        image_file = new BytesIO(blob);
        from_blob_.returnValue = image_;
        return [image_file, blob, image_];
      } },
    pil_fixture: { args: [], fn: (self: any): any => {
        let blob: any, dpi: any, f: any, format: any, image: any, size: any;
        f = $open(test_image_path, "rb");
        blob = $read(f);
        image = $call(Image, blob, null);
        size = [204, 204];
        format = "JPEG";
        dpi = [72, 72];
        return [image, size, format, dpi];
      } },
    _format_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Image, "_format");
      } },
    from_blob_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Image, "from_blob", $kw({autospec: false}));
      } },
    image_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Image);
      } },
    _init_: { args: ["request"], fn: (self: any, request: any): any => {
        return initializerMock(request, Image);
      } },
    _pil_props_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Image, "_pil_props");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_a_path", args: ["from_blob_", "image_"], fn: (self: any, from_blob_: any, image_: any): any => {
        let blob: any, f: any, image: any;
        f = $open(test_image_path, "rb");
        blob = $read(f);
        from_blob_.returnValue = image_;
        image = Image.fromFile(test_image_path);
        Image.fromBlob.assertCalledOnceWith(blob, "python-icon.jpeg");
        $assert($is(image, image_));
      } },
    { name: "it_can_construct_from_a_stream", args: ["from_stream_fixture"], fn: (self: any, from_stream_fixture: any): any => {
        let blob: any, image: any, image_: any, image_file: any;
        [image_file, blob, image_] = $list(from_stream_fixture);
        image = Image.fromFile(image_file);
        Image.fromBlob.assertCalledOnceWith(blob, null);
        $assert($is(image, image_));
      } },
    { name: "it_can_construct_from_a_blob", args: ["_init_"], fn: (self: any, _init_: any): any => {
        let image: any;
        image = Image.fromBlob($b("blob"), "foo.png");
        _init_.assertCalledOnceWith(image, $b("blob"), "foo.png");
        $assert($isinstance(image, Image));
      } },
    { name: "it_knows_its_blob", args: ["blob_fixture"], fn: (self: any, blob_fixture: any): any => {
        let expected_value: any, image: any;
        [image, expected_value] = $list(blob_fixture);
        $assertEqual(image.blob, expected_value);
      } },
    { name: "it_knows_its_content_type", args: ["content_type_fixture"], fn: (self: any, content_type_fixture: any): any => {
        let expected_value: any, image: any;
        [image, expected_value] = $list(content_type_fixture);
        $assertEqual(image.contentType, expected_value);
      } },
    { name: "it_knows_its_canonical_filename_extension", args: ["ext_fixture"], fn: (self: any, ext_fixture: any): any => {
        let expected_value: any, image: any;
        [image, expected_value] = $list(ext_fixture);
        $assertEqual(image.ext, expected_value);
      } },
    { name: "it_knows_its_dpi", args: ["dpi_fixture"], fn: (self: any, dpi_fixture: any): any => {
        let expected_value: any, image: any;
        [image, expected_value] = $list(dpi_fixture);
        $assertEqual(image.dpi, expected_value);
      } },
    { name: "it_knows_its_filename", args: ["filename_fixture"], fn: (self: any, filename_fixture: any): any => {
        let expected_value: any, image: any;
        [image, expected_value] = $list(filename_fixture);
        $assertEqual(image.filename, expected_value);
      } },
    { name: "it_knows_its_sha1_hash", args: [], fn: (self: any): any => {
        let image: any;
        image = $call(Image, $b("foobar"), null);
        $assertEqual(image.sha1, "8843d7f92416211de9ebb963ff4ce28125932878");
      } },
    { name: "it_knows_its_PIL_properties_to_help", args: ["pil_fixture"], fn: (self: any, pil_fixture: any): any => {
        let dpi: any, format: any, image: any, size: any;
        [image, size, format, dpi] = $list(pil_fixture);
        $assertEqual(image.size, size);
        $assertEqual(image._format, format);
        $assertEqual(image.dpi, dpi);
        $assertEqual(image._pilProps, [format, size, null]);
      } },
  ],
};
pytestClass("DescribeImage", DescribeImage);

