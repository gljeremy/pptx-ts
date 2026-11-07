// Machine-translated from python-pptx tests/opc/test_serialized.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $d, $get, $hashlib, $in, $is, $kw, $len, $list, $mod, $raises, $read, $str, $strip, $zipfile, ANY, BytesIO, call, classMock, functionMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { PackageNotFoundError } from "../../src/exc.js";
import { CONTENT_TYPE as CT } from "../../src/opc/constants.js";
import { Part, _Relationships } from "../../src/opc/package.js";
import { CONTENT_TYPES_URI, PackURI } from "../../src/opc/packuri.js";
import { PackageReader, PackageWriter, _ContentTypesItem, _DirPkgReader, _PhysPkgReader, _PhysPkgWriter, _ZipPkgReader, _ZipPkgWriter } from "../../src/opc/serialized.js";
import { absjoin, snippetText, testFileDir } from "../unitutil/file.js";
void $py;


let test_pptx_path = absjoin(testFileDir, "test.pptx");

let dir_pkg_path = absjoin(testFileDir, "expanded_pptx");

let zip_pkg_path = test_pptx_path;

const DescribePackageReader: any = {
  fixtures: {
    _blob_reader_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, PackageReader, "_blob_reader");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_contains_a_partname", args: ["_blob_reader_prop_"], fn: (self: any, _blob_reader_prop_: any): any => {
        let package_reader: any;
        _blob_reader_prop_.returnValue = new Set(["/ppt", "/docProps"]);
        package_reader = $call(PackageReader, "");
        $assert($in("/ppt", package_reader));
        $assert(!$in("/xyz", package_reader));
      } },
    { name: "it_can_get_a_blob_by_partname", args: ["_blob_reader_prop_"], fn: (self: any, _blob_reader_prop_: any): any => {
        let package_reader: any;
        _blob_reader_prop_.returnValue = $d({"/ppt/slides/slide1.xml": $b("blob")});
        package_reader = $call(PackageReader, "");
        $assertEqual($get(package_reader, $call(PackURI, "/ppt/slides/slide1.xml")), $b("blob"));
      } },
    { name: "it_can_get_the_rels_xml_for_a_partname", args: ["_blob_reader_prop_"], fn: (self: any, _blob_reader_prop_: any): any => {
        let package_reader: any;
        _blob_reader_prop_.returnValue = $d({"/ppt/_rels/presentation.xml.rels": $b("blob")});
        package_reader = $call(PackageReader, "");
        $assertEqual(package_reader.relsXmlFor($call(PackURI, "/ppt/presentation.xml")), $b("blob"));
      } },
    { name: "but_it_returns_None_when_the_part_has_no_rels", args: ["_blob_reader_prop_"], fn: (self: any, _blob_reader_prop_: any): any => {
        let package_reader: any;
        _blob_reader_prop_.returnValue = $d({"/ppt/_rels/presentation.xml.rels": $b("blob")});
        package_reader = $call(PackageReader, "");
        $assert($is(package_reader.relsXmlFor($call(PackURI, "/ppt/slides.slide1.xml")), null));
      } },
    { name: "it_constructs_its_blob_reader_to_help", args: ["request"], fn: (self: any, request: any): any => {
        let _PhysPkgReader_: any, blob_reader: any, package_reader: any, phys_pkg_reader_: any;
        phys_pkg_reader_ = instanceMock(request, _PhysPkgReader);
        _PhysPkgReader_ = classMock(request, "pptx.opc.serialized._PhysPkgReader");
        _PhysPkgReader_.factory.returnValue = phys_pkg_reader_;
        package_reader = $call(PackageReader, "prs.pptx");
        blob_reader = package_reader._blobReader;
        _PhysPkgReader_.factory.assertCalledOnceWith("prs.pptx");
        $assert($is(blob_reader, phys_pkg_reader_));
      } },
  ],
};
pytestClass("DescribePackageReader", DescribePackageReader);

const DescribePackageWriter: any = {
  fixtures: {
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
    phys_writer_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _ZipPkgWriter);
      } },
    relationships_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationships);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_write_interface_classmethod", args: ["request", "relationships_", "part_"], fn: (self: any, request: any, relationships_: any, part_: any): any => {
        let _init_: any, _write_: any;
        _init_ = initializerMock(request, PackageWriter);
        _write_ = methodMock(request, PackageWriter, "_write");
        PackageWriter.write("prs.pptx", relationships_, [part_, part_]);
        _init_.assertCalledOnceWith(ANY, "prs.pptx", relationships_, [part_, part_]);
        _write_.assertCalledOnceWith(ANY);
      } },
    { name: "it_can_write_a_package", args: ["request", "phys_writer_", "relationships_"], fn: (self: any, request: any, phys_writer_: any, relationships_: any): any => {
        let _PhysPkgWriter_: any, _write_content_types_stream_: any, _write_parts_: any, _write_pkg_rels_: any, package_writer: any;
        _PhysPkgWriter_ = classMock(request, "pptx.opc.serialized._PhysPkgWriter");
        phys_writer_.__enter__.returnValue = phys_writer_;
        _PhysPkgWriter_.factory.returnValue = phys_writer_;
        _write_content_types_stream_ = methodMock(request, PackageWriter, "_write_content_types_stream");
        _write_pkg_rels_ = methodMock(request, PackageWriter, "_write_pkg_rels");
        _write_parts_ = methodMock(request, PackageWriter, "_write_parts");
        package_writer = $call(PackageWriter, "prs.pptx", relationships_, []);
        package_writer._write();
        _PhysPkgWriter_.factory.assertCalledOnceWith("prs.pptx");
        _write_content_types_stream_.assertCalledOnceWith(package_writer, phys_writer_);
        _write_pkg_rels_.assertCalledOnceWith(package_writer, phys_writer_);
        _write_parts_.assertCalledOnceWith(package_writer, phys_writer_);
      } },
    { name: "it_can_write_a_content_types_stream", args: ["request", "phys_writer_", "relationships_", "part_"], fn: (self: any, request: any, phys_writer_: any, relationships_: any, part_: any): any => {
        let _ContentTypesItem_: any, package_writer: any, serialize_part_xml_: any;
        _ContentTypesItem_ = classMock(request, "pptx.opc.serialized._ContentTypesItem");
        _ContentTypesItem_.xmlFor.returnValue = "part_xml";
        serialize_part_xml_ = $call(functionMock, request, "pptx.opc.serialized.serialize_part_xml", $kw({return_value: $b("xml")}));
        package_writer = $call(PackageWriter, "", relationships_, [part_, part_]);
        package_writer._writeContentTypesStream(phys_writer_);
        _ContentTypesItem_.xmlFor.assertCalledOnceWith([part_, part_]);
        serialize_part_xml_.assertCalledOnceWith("part_xml");
        phys_writer_.write.assertCalledOnceWith(CONTENT_TYPES_URI, $b("xml"));
      } },
    { name: "it_can_write_a_sequence_of_parts", args: ["request", "relationships_", "phys_writer_"], fn: (self: any, request: any, relationships_: any, phys_writer_: any): any => {
        let package_writer: any, parts_: any;
        parts_ = $list(["a", "b", "c"]).map((x: any) => $call(instanceMock, request, Part, $kw({partname: $call(PackURI, $mod("/ppt/%s.xml", x)), blob: $mod("blob_%s", x), rels: $call(instanceMock, request, _Relationships, $kw({xml: $mod("rels_xml_%s", x)}))})));
        package_writer = $call(PackageWriter, "", relationships_, parts_);
        package_writer._writeParts(phys_writer_);
        $assertEqual(phys_writer_.write.callArgsList, [call("/ppt/a.xml", "blob_a"), call("/ppt/_rels/a.xml.rels", "rels_xml_a"), call("/ppt/b.xml", "blob_b"), call("/ppt/_rels/b.xml.rels", "rels_xml_b"), call("/ppt/c.xml", "blob_c"), call("/ppt/_rels/c.xml.rels", "rels_xml_c")]);
      } },
    { name: "it_can_write_a_pkg_rels_item", args: ["phys_writer_", "relationships_"], fn: (self: any, phys_writer_: any, relationships_: any): any => {
        let package_writer: any;
        relationships_.xml = $b("pkg-rels-xml");
        package_writer = $call(PackageWriter, "", relationships_, []);
        package_writer._writePkgRels(phys_writer_);
        phys_writer_.write.assertCalledOnceWith("/_rels/.rels", $b("pkg-rels-xml"));
      } },
  ],
};
pytestClass("DescribePackageWriter", DescribePackageWriter);

const Describe_PhysPkgReader: any = {
  fixtures: {
    zip_pkg_reader_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _ZipPkgReader);
      } },
    _ZipPkgReader_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.opc.serialized._ZipPkgReader");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_ZipPkgReader_when_pkg_is_file_like", args: ["_ZipPkgReader_", "zip_pkg_reader_"], fn: (self: any, _ZipPkgReader_: any, zip_pkg_reader_: any): any => {
        let file_like_pkg: any, phys_reader: any;
        _ZipPkgReader_.returnValue = zip_pkg_reader_;
        file_like_pkg = new BytesIO($b("pkg-bytes"));
        phys_reader = _PhysPkgReader.factory(file_like_pkg);
        _ZipPkgReader_.assertCalledOnceWith(file_like_pkg);
        $assert($is(phys_reader, zip_pkg_reader_));
      } },
    { name: "and_it_constructs_DirPkgReader_when_pkg_is_a_dir", args: ["request"], fn: (self: any, request: any): any => {
        let _DirPkgReader_: any, dir_pkg_reader_: any, phys_reader: any;
        dir_pkg_reader_ = instanceMock(request, _DirPkgReader);
        _DirPkgReader_ = $call(classMock, request, "pptx.opc.serialized._DirPkgReader", $kw({return_value: dir_pkg_reader_}));
        phys_reader = _PhysPkgReader.factory(dir_pkg_path);
        _DirPkgReader_.assertCalledOnceWith(dir_pkg_path);
        $assert($is(phys_reader, dir_pkg_reader_));
      } },
    { name: "and_it_constructs_ZipPkgReader_when_pkg_is_a_zip_file_path", args: ["_ZipPkgReader_", "zip_pkg_reader_"], fn: (self: any, _ZipPkgReader_: any, zip_pkg_reader_: any): any => {
        let phys_reader: any, pkg_file_path: any;
        _ZipPkgReader_.returnValue = zip_pkg_reader_;
        pkg_file_path = test_pptx_path;
        phys_reader = _PhysPkgReader.factory(pkg_file_path);
        _ZipPkgReader_.assertCalledOnceWith(pkg_file_path);
        $assert($is(phys_reader, zip_pkg_reader_));
      } },
    { name: "but_it_raises_when_pkg_path_is_not_a_package", args: [], fn: (self: any): any => {
        let e: any;
        e = $raises(PackageNotFoundError, () => {
          _PhysPkgReader.factory("foobar");
        });
        $assertEqual($str(e.value), "Package not found at 'foobar'");
      } },
  ],
};
pytestClass("Describe_PhysPkgReader", Describe_PhysPkgReader);

const Describe_DirPkgReader: any = {
  fixtures: {
    dir_pkg_reader: { args: [], scope: "class", fn: (self: any): any => {
        return $call(_DirPkgReader, dir_pkg_path);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_contains_a_partname", args: ["dir_pkg_reader"], fn: (self: any, dir_pkg_reader: any): any => {
        $assert($in($call(PackURI, "/ppt/presentation.xml"), dir_pkg_reader));
        $assert(!$in($call(PackURI, "/ppt/foobar.xml"), dir_pkg_reader));
      } },
    { name: "it_can_retrieve_the_blob_for_a_pack_uri", args: ["dir_pkg_reader"], fn: (self: any, dir_pkg_reader: any): any => {
        let blob: any;
        blob = $get(dir_pkg_reader, $call(PackURI, "/ppt/presentation.xml"));
        $assertEqual($hashlib.sha1(blob).hexdigest(), "51b78f4dabc0af2419d4e044ab73028c4bef53aa");
      } },
    { name: "but_it_raises_KeyError_when_requested_member_is_not_present", args: ["dir_pkg_reader"], fn: (self: any, dir_pkg_reader: any): any => {
        let e: any;
        e = $raises(KeyError, () => {
          $get(dir_pkg_reader, $call(PackURI, "/ppt/foobar.xml"));
        });
        $assertEqual($str(e.value), "\"no member '/ppt/foobar.xml' in package\"");
      } },
  ],
};
pytestClass("Describe_DirPkgReader", Describe_DirPkgReader);

const Describe_ZipPkgReader: any = {
  fixtures: {
    zip_pkg_reader: { args: [], scope: "class", fn: (self: any): any => {
        return $call(_ZipPkgReader, zip_pkg_path);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_contains_a_partname", args: ["zip_pkg_reader"], fn: (self: any, zip_pkg_reader: any): any => {
        $assert($in($call(PackURI, "/ppt/presentation.xml"), zip_pkg_reader));
        $assert(!$in($call(PackURI, "/ppt/foobar.xml"), zip_pkg_reader));
      } },
    { name: "it_can_get_a_blob_by_partname", args: ["zip_pkg_reader"], fn: (self: any, zip_pkg_reader: any): any => {
        let blob: any;
        blob = $get(zip_pkg_reader, $call(PackURI, "/ppt/presentation.xml"));
        $assertEqual($hashlib.sha1(blob).hexdigest(), "efa7bee0ac72464903a67a6744c1169035d52a54");
      } },
    { name: "but_it_raises_KeyError_when_requested_member_is_not_present", args: ["zip_pkg_reader"], fn: (self: any, zip_pkg_reader: any): any => {
        let e: any;
        e = $raises(KeyError, () => {
          $get(zip_pkg_reader, $call(PackURI, "/ppt/foobar.xml"));
        });
        $assertEqual($str(e.value), "\"no member '/ppt/foobar.xml' in package\"");
      } },
    { name: "it_loads_the_package_blobs_on_first_access_to_help", args: ["zip_pkg_reader"], fn: (self: any, zip_pkg_reader: any): any => {
        let blobs: any;
        blobs = zip_pkg_reader._blobs;
        $assertEqual($len(blobs), 38);
        $assert($in("/ppt/presentation.xml", blobs));
        $assert($in("/ppt/_rels/presentation.xml.rels", blobs));
      } },
  ],
};
pytestClass("Describe_ZipPkgReader", Describe_ZipPkgReader);

const Describe_PhysPkgWriter: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_ZipPkgWriter_unconditionally", args: ["request"], fn: (self: any, request: any): any => {
        let _ZipPkgWriter_: any, phys_writer: any, zip_pkg_writer_: any;
        zip_pkg_writer_ = instanceMock(request, _ZipPkgWriter);
        _ZipPkgWriter_ = $call(classMock, request, "pptx.opc.serialized._ZipPkgWriter", $kw({return_value: zip_pkg_writer_}));
        phys_writer = _PhysPkgWriter.factory("prs.pptx");
        _ZipPkgWriter_.assertCalledOnceWith("prs.pptx");
        $assert($is(phys_writer, zip_pkg_writer_));
      } },
  ],
};
pytestClass("Describe_PhysPkgWriter", Describe_PhysPkgWriter);

const Describe_ZipPkgWriter: any = {
  fixtures: {
    _zipf_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _ZipPkgWriter, "_zipf");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_has_an__enter__method_for_context_management", skip: "python context-manager protocol; the TS writer is closed by PackageWriter._write()", args: [], fn: (self: any): any => {
        let pkg_writer: any;
        pkg_writer = $call(_ZipPkgWriter, "");
        $assert($is(pkg_writer.__enter__(), pkg_writer));
      } },
    { name: "and_it_closes_the_zip_archive_on_context__exit__", skip: "python context-manager protocol", args: ["_zipf_prop_"], fn: (self: any, _zipf_prop_: any): any => {
        $call(_ZipPkgWriter, "").__exit__();
        _zipf_prop_.returnValue.close.assertCalledOnceWith();
      } },
    { name: "it_can_write_a_blob", skip: "exercises python's zipfile.ZipFile; covered by test/opc/zip_writer.test.ts", args: ["_zipf_prop_"], fn: (self: any, _zipf_prop_: any): any => {
        let members: any, pack_uri: any, pkg_writer: any, zipf: any;
        pack_uri = $call(PackURI, "/part/name.xml");
        _zipf_prop_.returnValue = $zipfile.ZipFile(new BytesIO(), "w");
        zipf = $zipfile.ZipFile(new BytesIO(), "w");
        pkg_writer = $call(_ZipPkgWriter, "");
        pkg_writer.write(pack_uri, $b("blob"));
        members = $d(Object.fromEntries($list(zipf.namelist()).map((name: any) => [$call(PackURI, $mod("/%s", name)), $read(zipf, name)])));
        $assertEqual($len(members), 1);
        $assertEqual($get(members, pack_uri), $b("blob"));
      } },
    { name: "it_provides_access_to_the_open_zip_file_to_help", skip: "python's zipfile.ZipFile; the TS writer uses fflate", args: ["request"], fn: (self: any, request: any): any => {
        let ZipFile_: any, pkg_writer: any, zipf: any;
        ZipFile_ = classMock(request, "pptx.opc.serialized.zipfile.ZipFile");
        pkg_writer = $call(_ZipPkgWriter, "prs.pptx");
        zipf = pkg_writer._zipf;
        ZipFile_.assertCalledOnceWith("prs.pptx", "w", $kw({compression: $zipfile.ZIP_DEFLATED, strict_timestamps: false}));
        $assert($is(zipf, ZipFile_.returnValue));
      } },
  ],
};
pytestClass("Describe_ZipPkgWriter", Describe_ZipPkgWriter);

const Describe_ContentTypesItem: any = {
  fixtures: {
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_an_interface_classmethod", args: ["request", "part_"], fn: (self: any, request: any, part_: any): any => {
        let _init_: any, xml: any;
        _init_ = initializerMock(request, _ContentTypesItem);
        $call(propertyMock, request, _ContentTypesItem, "_xml", $kw({return_value: $b("xml")}));
        xml = _ContentTypesItem.xmlFor([part_, part_]);
        _init_.assertCalledOnceWith(ANY, [part_, part_]);
        $assertEqual(xml, $b("xml"));
      } },
    { name: "it_can_compose_content_types_xml", args: ["request"], fn: (self: any, request: any): any => {
        let content_types: any, defaults: any, overrides: any;
        defaults = $d({png: CT.PNG, xml: CT.XML, rels: CT.OPC_RELATIONSHIPS});
        overrides = $d({"/docProps/core.xml": "app/vnd.core", "/ppt/slides/slide1.xml": "app/vnd.ct_sld", "/zebra/foo.bar": "app/vnd.foobar"});
        $call(propertyMock, request, _ContentTypesItem, "_defaults_and_overrides", $kw({return_value: [defaults, overrides]}));
        content_types = $call(_ContentTypesItem, [])._xml;
        $assertEqual(content_types.xml, $strip(snippetText("content-types-xml")));
      } },
    { name: "it_computes_defaults_and_overrides_to_help", args: ["request"], fn: (self: any, request: any): any => {
        let content_types: any, defaults: any, overrides: any, parts: any;
        parts = $list([["/media/image1.png", CT.PNG], ["/ppt/slides/slide1.xml", CT.PML_SLIDE], ["/foo/bar.xml", CT.XML], ["/docProps/core.xml", CT.OPC_CORE_PROPERTIES]]).map(([partname, content_type]: any) => $call(instanceMock, request, Part, $kw({partname: $call(PackURI, partname), content_type: content_type})));
        content_types = $call(_ContentTypesItem, parts);
        [defaults, overrides] = $list(content_types._defaultsAndOverrides);
        $assertEqual(defaults, $d({png: CT.PNG, rels: CT.OPC_RELATIONSHIPS, xml: CT.XML}));
        $assertEqual(overrides, $d({"/ppt/slides/slide1.xml": CT.PML_SLIDE, "/docProps/core.xml": CT.OPC_CORE_PROPERTIES}));
      } },
  ],
};
pytestClass("Describe_ContentTypesItem", Describe_ContentTypesItem);

