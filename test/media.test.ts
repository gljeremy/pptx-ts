// Machine-translated from python-pptx tests/test_media.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $assertEqual, $b, $call, $is, $isinstance, $kw, $list, $open, $read, BytesIO, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { Video } from "../src/media.js";
import { absjoin, testFileDir } from "./unitutil/file.js";
void $py;


let TEST_VIDEO_PATH = absjoin(testFileDir, "dummy.mp4");

const DescribeVideo: any = {
  fixtures: {
    blob_fixture: { args: [], fn: (self: any): any => {
        let blob: any, expected_value: any, video: any;
        blob = $b("blob-bytes");
        video = $call(Video, blob, null, null);
        expected_value = blob;
        return [video, expected_value];
      } },
    content_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, mime_type: any, video: any;
        mime_type = "video/mp4";
        video = $call(Video, null, mime_type, null);
        expected_value = mime_type;
        return [video, expected_value];
      } },
    ext_fixture: { args: ["request"], params: [[null, "foo.bar", "bar"], ["video/mp4", null, "mp4"], ["video/xyz", null, "vid"]], fn: (self: any, request: any): any => {
        let expected_value: any, filename: any, mime_type: any, video: any;
        [mime_type, filename, expected_value] = $list(request.param);
        video = $call(Video, null, mime_type, filename);
        return [video, expected_value];
      } },
    filename_fixture: { args: ["request", "ext_prop_"], params: [["foobar.mp4", null, "foobar.mp4"], [null, "vid", "movie.vid"]], fn: (self: any, request: any, ext_prop_: any): any => {
        let expected_value: any, ext: any, filename: any, video: any;
        [filename, ext, expected_value] = $list(request.param);
        video = $call(Video, null, null, filename);
        ext_prop_.returnValue = ext;
        return [video, expected_value];
      } },
    from_blob_fixture: { args: ["Video_init_"], fn: (self: any, Video_init_: any): any => {
        let blob: any, filename: any, mime_type: any;
        [blob, mime_type, filename] = $list(["01234", "video/mp4", "movie.mp4"]);
        return [blob, mime_type, filename, Video_init_];
      } },
    from_stream_fixture: { args: ["video_", "from_blob_"], fn: (self: any, video_: any, from_blob_: any): any => {
        let blob: any, f: any, mime_type: any, movie_stream: any;
        f = $open(TEST_VIDEO_PATH, "rb");
        blob = $read(f);
        movie_stream = new BytesIO(blob);
        mime_type = "video/mp4";
        from_blob_.returnValue = video_;
        return [movie_stream, mime_type, blob, video_];
      } },
    sha1_fixture: { args: [], fn: (self: any): any => {
        let blob: any, expected_value: any, video: any;
        blob = $b("blobish");
        video = $call(Video, blob, null, null);
        expected_value = "de731a6eed12f427642325193b8e57af3c624d62";
        return [video, expected_value];
      } },
    ext_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Video, "ext");
      } },
    from_blob_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Video, "from_blob", $kw({autospec: false}));
      } },
    video_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Video);
      } },
    Video_init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, Video, $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_a_path", args: ["video_", "from_blob_"], fn: (self: any, video_: any, from_blob_: any): any => {
        let blob: any, f: any, video: any;
        f = $open(TEST_VIDEO_PATH, "rb");
        blob = $read(f);
        from_blob_.returnValue = video_;
        video = Video.fromPathOrFileLike(TEST_VIDEO_PATH, "video/mp4");
        Video.fromBlob.assertCalledOnceWith(blob, "video/mp4", "dummy.mp4");
        $assert($is(video, video_));
      } },
    { name: "it_can_construct_from_a_stream", args: ["from_stream_fixture"], fn: (self: any, from_stream_fixture: any): any => {
        let blob: any, mime_type: any, movie_stream: any, video: any, video_: any;
        [movie_stream, mime_type, blob, video_] = $list(from_stream_fixture);
        video = Video.fromPathOrFileLike(movie_stream, mime_type);
        Video.fromBlob.assertCalledOnceWith(blob, mime_type, null);
        $assert($is(video, video_));
      } },
    { name: "it_can_construct_from_a_blob", args: ["from_blob_fixture"], fn: (self: any, from_blob_fixture: any): any => {
        let Video_init_: any, blob: any, filename: any, mime_type: any, video: any;
        [blob, mime_type, filename, Video_init_] = $list(from_blob_fixture);
        video = Video.fromBlob(blob, mime_type, filename);
        Video_init_.assertCalledOnceWith(video, blob, mime_type, filename);
        $assert($isinstance(video, Video));
      } },
    { name: "it_provides_access_to_the_video_bytestream", args: ["blob_fixture"], fn: (self: any, blob_fixture: any): any => {
        let expected_value: any, video: any;
        [video, expected_value] = $list(blob_fixture);
        $assertEqual(video.blob, expected_value);
      } },
    { name: "it_knows_its_content_type", args: ["content_type_fixture"], fn: (self: any, content_type_fixture: any): any => {
        let expected_value: any, video: any;
        [video, expected_value] = $list(content_type_fixture);
        $assertEqual(video.contentType, expected_value);
      } },
    { name: "it_knows_a_filename_for_the_video", args: ["filename_fixture"], fn: (self: any, filename_fixture: any): any => {
        let expected_value: any, video: any;
        [video, expected_value] = $list(filename_fixture);
        $assertEqual(video.filename, expected_value);
      } },
    { name: "it_knows_an_extension_for_the_video", args: ["ext_fixture"], fn: (self: any, ext_fixture: any): any => {
        let expected_value: any, video: any;
        [video, expected_value] = $list(ext_fixture);
        $assertEqual(video.ext, expected_value);
      } },
    { name: "it_knows_its_sha1_hash", args: ["sha1_fixture"], fn: (self: any, sha1_fixture: any): any => {
        let expected_value: any, video: any;
        [video, expected_value] = $list(sha1_fixture);
        $assertEqual(video.sha1, expected_value);
      } },
  ],
};
pytestClass("DescribeVideo", DescribeVideo);

