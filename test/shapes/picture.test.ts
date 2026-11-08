// Machine-translated from python-pptx tests/shapes/test_picture.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $abs, $append, $assert, $assertEqual, $call, $cmp, $getattr, $is, $isinstance, $kw, $list, $mod, $setattr, $t1, $truthy, call, classMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { LineFormat as LineFormat$imported } from "../../src/dml/line.js";
import { MSO_SHAPE as MSO_SHAPE$imported, MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported, PP_MEDIA_TYPE as PP_MEDIA_TYPE$imported } from "../../src/enum/shapes.js";
import { Image as Image$imported } from "../../src/parts/image.js";
import { SlidePart as SlidePart$imported } from "../../src/parts/slide.js";
import { Movie as Movie$imported, Picture as Picture$imported, _BasePicture as _BasePicture$imported, _MediaFormat as _MediaFormat$imported } from "../../src/shapes/picture.js";
import { Pt as Pt$imported } from "../../src/util.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const LineFormat = LineFormat$imported;
const MSO_SHAPE = MSO_SHAPE$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const PP_MEDIA_TYPE = PP_MEDIA_TYPE$imported;
const Image = Image$imported;
const SlidePart = SlidePart$imported;
const Movie = Movie$imported;
const Picture = Picture$imported;
const _BasePicture = _BasePicture$imported;
const _MediaFormat = _MediaFormat$imported;
const Pt = Pt$imported;
void $py;


const Describe_BasePicture: any = {
  fixtures: {
    crop_get_fixture: { args: ["request"], params: [["p:pic/p:blipFill", "left", 0.0], ["p:pic/p:blipFill/a:srcRect", "top", 0.0], ["p:pic/p:blipFill/a:srcRect{l=99999}", "bottom", 0.0], ["p:pic/p:blipFill/a:srcRect{l=42424}", "left", 0.42424], ["p:pic/p:blipFill/a:srcRect{t=-10000}", "top", -0.1], ["p:pic/p:blipFill/a:srcRect{r=250000}", "right", 2.5], ["p:pic/p:blipFill/a:srcRect{b=33333}", "bottom", 0.33333]], fn: (self: any, request: any): any => {
        let expected_value: any, pic_cxml: any, picture: any, prop_name: any, side: any;
        [pic_cxml, side, expected_value] = $list(request.param);
        picture = $call(Picture, element(pic_cxml), null);
        prop_name = $mod("crop_%s", side);
        return [picture, prop_name, expected_value];
      } },
    crop_set_fixture: { args: ["request"], params: [["p:pic{a:b=c}/p:blipFill", "left", 0.11, "p:pic{a:b=c}/p:blipFill/a:srcRect{l=11000}"], ["p:pic{a:b=c}/p:blipFill", "top", 0.21, "p:pic{a:b=c}/p:blipFill/a:srcRect{t=21000}"], ["p:pic{a:b=c}/p:blipFill", "right", 0.31, "p:pic{a:b=c}/p:blipFill/a:srcRect{r=31000}"], ["p:pic{a:b=c}/p:blipFill", "bottom", 0.41, "p:pic{a:b=c}/p:blipFill/a:srcRect{b=41000}"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{l=80000}", "left", 0.21, "p:pic{a:b=c}/p:blipFill/a:srcRect{l=21000}"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{t=70000}", "top", 0.22, "p:pic{a:b=c}/p:blipFill/a:srcRect{t=22000}"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{r=60000}", "right", 0.23, "p:pic{a:b=c}/p:blipFill/a:srcRect{r=23000}"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{b=50000}", "bottom", 0.24, "p:pic{a:b=c}/p:blipFill/a:srcRect{b=24000}"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{l=90000}", "left", 0, "p:pic{a:b=c}/p:blipFill/a:srcRect"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{t=91000}", "top", 0.0, "p:pic{a:b=c}/p:blipFill/a:srcRect"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{r=92000}", "right", 0, "p:pic{a:b=c}/p:blipFill/a:srcRect"], ["p:pic{a:b=c}/p:blipFill/a:srcRect{b=93000}", "bottom", 0.0, "p:pic{a:b=c}/p:blipFill/a:srcRect"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, pic: any, pic_cxml: any, picture: any, prop_name: any, side: any, value: any;
        [pic_cxml, side, value, expected_cxml] = $list(request.param);
        pic = element(pic_cxml);
        picture = $call(Picture, pic, null);
        prop_name = $mod("crop_%s", side);
        expected_xml = xml(expected_cxml);
        return [picture, prop_name, value, expected_xml];
      } },
    line_fixture: { args: [], fn: (self: any): any => {
        return $call(_BasePicture, element("p:pic/p:spPr"), null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_cropping", args: ["crop_get_fixture"], fn: (self: any, crop_get_fixture: any): any => {
        let crop: any, expected_value: any, picture: any, prop_name: any;
        [picture, prop_name, expected_value] = $list(crop_get_fixture);
        crop = $getattr(picture, prop_name);
        $assert($cmp($abs((crop - expected_value)), "<", 1e-06));
      } },
    { name: "it_can_change_its_cropping", args: ["crop_set_fixture"], fn: (self: any, crop_set_fixture: any): any => {
        let expected_xml: any, picture: any, prop_name: any, value: any;
        [picture, prop_name, value, expected_xml] = $list(crop_set_fixture);
        $setattr(picture, prop_name, value);
        $assertEqual(picture._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_outline", args: ["line_fixture"], fn: (self: any, line_fixture: any): any => {
        let line: any, picture: any;
        picture = line_fixture;
        line = picture.line;
        $assert($isinstance(line, LineFormat));
        picture.line.width = $call(Pt, 1);
        $assertEqual(picture.line.width, $call(Pt, 1));
      } },
  ],
};
pytestClass("Describe_BasePicture", Describe_BasePicture);

const DescribeMovie: any = {
  fixtures: {
    format_fixture: { args: ["_MediaFormat_", "media_format_"], fn: (self: any, _MediaFormat_: any, media_format_: any): any => {
        let movie: any, parent: any, pic: any;
        pic = element("p:pic");
        const $t1 = $call(Movie, pic, null);
        parent = $t1;
        movie = $t1;
        return [movie, _MediaFormat_, pic, parent, media_format_];
      } },
    media_type_fixture: { args: [], fn: (self: any): any => {
        return $call(Movie, null, null);
      } },
    pfrm_fixture: { args: ["request", "slide_part_", "image_", "part_prop_"], params: [["p:pic/p:blipFill/a:blip{r:embed=rId42}", true], ["p:pic/p:blipFill/a:blip", false]], fn: (self: any, request: any, slide_part_: any, image_: any, part_prop_: any): any => {
        let calls: any, cxml: any, expected_value: any, is_present: any, movie: any;
        [cxml, is_present] = $list(request.param);
        movie = $call(Movie, element(cxml), null);
        [calls, expected_value] = $list([[], null]);
        part_prop_.returnValue = slide_part_;
        slide_part_.getImage.returnValue = image_;
        if ($truthy(is_present)) {
          $append(calls, call("rId42"));
          expected_value = image_;
        }
        return [movie, slide_part_, calls, expected_value];
      } },
    shape_type_fixture: { args: [], fn: (self: any): any => {
        return $call(Movie, null, null);
      } },
    image_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Image);
      } },
    _MediaFormat_: { args: ["request", "media_format_"], fn: (self: any, request: any, media_format_: any): any => {
        return $call(classMock, request, "pptx.shapes.picture._MediaFormat", $kw({return_value: media_format_}));
      } },
    media_format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _MediaFormat);
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Movie, "part");
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_shape_type", args: ["shape_type_fixture"], fn: (self: any, shape_type_fixture: any): any => {
        let movie: any;
        movie = shape_type_fixture;
        $assertEqual(movie.shapeType, MSO_SHAPE_TYPE.MEDIA);
      } },
    { name: "it_knows_its_media_type", args: ["media_type_fixture"], fn: (self: any, media_type_fixture: any): any => {
        let movie: any;
        movie = media_type_fixture;
        $assertEqual(movie.mediaType, PP_MEDIA_TYPE.MOVIE);
      } },
    { name: "it_provides_access_to_its_media_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let MediaFormat_: any, media_format: any, media_format_: any, movie: any, parent: any, pic: any;
        [movie, MediaFormat_, pic, parent, media_format_] = $list(format_fixture);
        media_format = movie.mediaFormat;
        MediaFormat_.assertCalledOnceWith(pic, parent);
        $assert($is(media_format, media_format_));
      } },
    { name: "it_provides_access_to_its_poster_frame_image", args: ["pfrm_fixture"], fn: (self: any, pfrm_fixture: any): any => {
        let calls: any, expected_value: any, movie: any, poster_frame: any, slide_part_: any;
        [movie, slide_part_, calls, expected_value] = $list(pfrm_fixture);
        poster_frame = movie.posterFrame;
        $assertEqual(slide_part_.getImage.callArgsList, calls);
        $assertEqual(poster_frame, expected_value);
      } },
  ],
};
pytestClass("DescribeMovie", DescribeMovie);

const DescribePicture: any = {
  fixtures: {
    autoshape_get_fixture: { args: ["request"], params: [["", null], ["/a:prstGeom{prst=rect}", MSO_SHAPE.RECTANGLE], ["/a:prstGeom{prst=hexagon}", MSO_SHAPE.HEXAGON]], fn: (self: any, request: any): any => {
        let expected_value: any, pic_cxml: any, picture: any, prstGeom_cxml: any;
        [prstGeom_cxml, expected_value] = $list(request.param);
        pic_cxml = $mod("p:pic/p:spPr%s", prstGeom_cxml);
        picture = $call(Picture, element(pic_cxml), null);
        return [picture, expected_value];
      } },
    autoshape_set_fixture: { args: ["request"], params: [["p:pic/p:spPr/a:custGeom", MSO_SHAPE.RECTANGLE, "p:pic/p:spPr/a:prstGeom{prst=rect}"], ["p:pic/p:spPr/a:prstGeom{prst=rect}", MSO_SHAPE.OVAL, "p:pic/p:spPr/a:prstGeom{prst=ellipse}"], ["p:pic/p:spPr/a:prstGeom{prst=ellipse}", MSO_SHAPE.PIE, "p:pic/p:spPr/a:prstGeom{prst=pie}"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, new_value: any, pic_cxml: any, picture: any;
        [pic_cxml, new_value, expected_cxml] = $list(request.param);
        picture = $call(Picture, element(pic_cxml), null);
        expected_xml = xml(expected_cxml);
        return [picture, new_value, expected_xml];
      } },
    image_fixture: { args: ["part_prop_", "slide_part_", "image_"], fn: (self: any, part_prop_: any, slide_part_: any, image_: any): any => {
        let pic_cxml: any, picture: any, rId: any;
        [pic_cxml, rId] = $list(["p:pic/p:blipFill/a:blip{r:embed=rId42}", "rId42"]);
        picture = $call(Picture, element(pic_cxml), null);
        slide_part_.getImage.returnValue = image_;
        return [picture, slide_part_, rId, image_];
      } },
    shape_type_fixture: { args: [], fn: (self: any): any => {
        return $call(Picture, null, null);
      } },
    image_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Image);
      } },
    part_prop_: { args: ["request", "slide_part_"], fn: (self: any, request: any, slide_part_: any): any => {
        return $call(propertyMock, request, Picture, "part", $kw({return_value: slide_part_}));
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_masking_shape", args: ["autoshape_get_fixture"], fn: (self: any, autoshape_get_fixture: any): any => {
        let auto_shape_type: any, expected_value: any, picture: any;
        [picture, expected_value] = $list(autoshape_get_fixture);
        auto_shape_type = picture.autoShapeType;
        $assertEqual(auto_shape_type, expected_value);
      } },
    { name: "it_can_change_its_masking_shape", args: ["autoshape_set_fixture"], fn: (self: any, autoshape_set_fixture: any): any => {
        let expected_xml: any, new_value: any, picture: any;
        [picture, new_value, expected_xml] = $list(autoshape_set_fixture);
        picture.autoShapeType = new_value;
        $assertEqual(picture._element.xml, expected_xml);
      } },
    { name: "it_knows_its_shape_type", args: ["shape_type_fixture"], fn: (self: any, shape_type_fixture: any): any => {
        let picture: any;
        picture = shape_type_fixture;
        $assertEqual(picture.shapeType, MSO_SHAPE_TYPE.PICTURE);
      } },
    { name: "it_provides_access_to_its_image", args: ["image_fixture"], fn: (self: any, image_fixture: any): any => {
        let image: any, image_: any, picture: any, rId: any, slide_part_: any;
        [picture, slide_part_, rId, image_] = $list(image_fixture);
        image = picture.image;
        slide_part_.getImage.assertCalledOnceWith(rId);
        $assert($is(image, image_));
      } },
  ],
};
pytestClass("DescribePicture", DescribePicture);

