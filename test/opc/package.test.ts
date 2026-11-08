// Machine-translated from python-pptx tests/opc/test_package.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $cm1, $collections, $d, $dict, $dictget, $enter, $enumerate, $eq, $exit, $get, $in, $is, $isinstance, $iter, $itertools, $kw, $len, $list, $m, $mod, $open, $raises, $range, $read, $remove, $set, $set_, $slice, $str, $truthy, $values, ANY, BytesIO, call, classMock, functionMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CONTENT_TYPE as CT$imported, RELATIONSHIP_TARGET_MODE as RTM$imported, RELATIONSHIP_TYPE as RT$imported } from "../../src/opc/constants.js";
import { CT_Relationship as CT_Relationship$imported, CT_Relationships as CT_Relationships$imported } from "../../src/opc/oxml.js";
import { OpcPackage, Part, PartFactory as PartFactory$imported, XmlPart, _ContentTypeMap, _PackageLoader, _RelatableMixin as _RelatableMixin$imported, _Relationship, _Relationships as _Relationships$imported } from "../../src/opc/package.js";
import { PACKAGE_URI as PACKAGE_URI$imported, PackURI as PackURI$imported } from "../../src/opc/packuri.js";
import { parseXml as parseXml$imported } from "../../src/oxml/index.js";
import { PresentationPart as PresentationPart$imported } from "../../src/parts/presentation.js";
import { element } from "../unitutil/cxml.js";
import { absjoin, snippetBytes, testFileDir, testfileBytes } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const CT = CT$imported;
const RTM = RTM$imported;
const RT = RT$imported;
const CT_Relationship = CT_Relationship$imported;
const CT_Relationships = CT_Relationships$imported;
const PartFactory = PartFactory$imported;
const _RelatableMixin = _RelatableMixin$imported;
const _Relationships = _Relationships$imported;
const PACKAGE_URI = PACKAGE_URI$imported;
const PackURI = PackURI$imported;
const parseXml = parseXml$imported;
const PresentationPart = PresentationPart$imported;
void $py;


const Describe_RelatableMixin: any = {
  fixtures: {
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
    relationship_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationship);
      } },
    relationships_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationships);
      } },
    _rels_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _RelatableMixin, "_rels");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_find_a_part_related_by_reltype", args: ["_rels_prop_", "relationships_", "part_"], fn: (self: any, _rels_prop_: any, relationships_: any, part_: any): any => {
        let mixin: any, related_part: any;
        relationships_.partWithReltype.returnValue = part_;
        _rels_prop_.returnValue = relationships_;
        mixin = $call(_RelatableMixin);
        related_part = mixin.partRelatedBy(RT.CHART);
        relationships_.partWithReltype.assertCalledOnceWith(RT.CHART);
        $assert($is(related_part, part_));
      } },
    { name: "it_can_establish_a_relationship_to_another_part", args: ["_rels_prop_", "relationships_", "part_"], fn: (self: any, _rels_prop_: any, relationships_: any, part_: any): any => {
        let mixin: any, rId: any;
        relationships_.getOrAdd.returnValue = "rId42";
        _rels_prop_.returnValue = relationships_;
        mixin = $call(_RelatableMixin);
        rId = mixin.relateTo(part_, RT.SLIDE);
        relationships_.getOrAdd.assertCalledOnceWith(RT.SLIDE, part_);
        $assertEqual(rId, "rId42");
      } },
    { name: "and_it_can_establish_a_relationship_to_an_external_link", args: ["request", "_rels_prop_", "relationships_"], fn: (self: any, request: any, _rels_prop_: any, relationships_: any): any => {
        let mixin: any, rId: any;
        relationships_.getOrAddExtRel.returnValue = "rId24";
        _rels_prop_.returnValue = relationships_;
        mixin = $call(_RelatableMixin);
        rId = $m(mixin, "relateTo", "http://url", RT.HYPERLINK, $kw({is_external: true}));
        relationships_.getOrAddExtRel.assertCalledOnceWith(RT.HYPERLINK, "http://url");
        $assertEqual(rId, "rId24");
      } },
    { name: "it_can_find_a_related_part_by_rId", args: ["request", "_rels_prop_", "relationships_", "relationship_", "part_"], fn: (self: any, request: any, _rels_prop_: any, relationships_: any, relationship_: any, part_: any): any => {
        let mixin: any, related_part: any;
        _rels_prop_.returnValue = relationships_;
        relationships_.__getitem__.returnValue = relationship_;
        relationship_.targetPart = part_;
        mixin = $call(_RelatableMixin);
        related_part = mixin.relatedPart("rId17");
        relationships_.__getitem__.assertCalledOnceWith("rId17");
        $assert($is(related_part, part_));
      } },
    { name: "it_can_find_a_target_ref_URI_by_rId", args: ["request", "_rels_prop_", "relationships_", "relationship_"], fn: (self: any, request: any, _rels_prop_: any, relationships_: any, relationship_: any): any => {
        let mixin: any, target_ref: any;
        _rels_prop_.returnValue = relationships_;
        relationships_.__getitem__.returnValue = relationship_;
        relationship_.targetRef = "http://url";
        mixin = $call(_RelatableMixin);
        target_ref = mixin.targetRef("rId9");
        relationships_.__getitem__.assertCalledOnceWith("rId9");
        $assertEqual(target_ref, "http://url");
      } },
  ],
};
pytestClass("Describe_RelatableMixin", Describe_RelatableMixin);

const DescribeOpcPackage: any = {
  fixtures: {
    relationships_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationships);
      } },
    _rels_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, OpcPackage, "_rels");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_open_a_pkg_file", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, _load_: any, package$: any, package_: any;
        package_ = instanceMock(request, OpcPackage);
        _init_ = initializerMock(request, OpcPackage);
        _load_ = $call(methodMock, request, OpcPackage, "_load", $kw({return_value: package_}));
        package$ = OpcPackage.open("package.pptx");
        _init_.assertCalledOnceWith(ANY, "package.pptx");
        _load_.assertCalledOnceWith(ANY);
        $assert($is(package$, package_));
      } },
    { name: "it_can_drop_a_relationship", args: ["_rels_prop_", "relationships_"], fn: (self: any, _rels_prop_: any, relationships_: any): any => {
        _rels_prop_.returnValue = relationships_;
        $call(OpcPackage, null).dropRel("rId42");
        relationships_.pop.assertCalledOnceWith("rId42");
      } },
    { name: "it_can_iterate_over_its_parts", args: ["request"], fn: (self: any, request: any): any => {
        let package$: any, part_: any, part_2_: any, rels_iter: any;
        [part_, part_2_] = $list($list($range(2)).map((i: any) => $call(instanceMock, request, Part, $kw({name: $mod("part_%d", i)}))));
        rels_iter = $list([[true, "http://some/url/"], [false, part_], [false, part_], [false, part_2_], [false, part_], [false, part_2_]]).map(([is_external, target]: any) => $call(instanceMock, request, _Relationship, $kw({is_external: is_external, target_part: target})));
        $call(methodMock, request, OpcPackage, "iter_rels", $kw({return_value: rels_iter}));
        package$ = $call(OpcPackage, null);
        $assertEqual($list(package$.iterParts()), [part_, part_2_]);
      } },
    { name: "it_can_iterate_over_its_relationships", args: ["request", "_rels_prop_"], fn: (self: any, request: any, _rels_prop_: any): any => {
        let all_rels: any, package$: any, part_0_: any, part_1_: any, rels: any;
        [part_0_, part_1_] = $list($list($range(2)).map((i: any) => $call(instanceMock, request, Part, $kw({name: $mod("part_%d", i)}))));
        all_rels = $list($list($enumerate([[false, part_0_], [false, part_1_], [true, null], [false, part_1_], [false, part_0_]])).map(([i, [ext, part]]: any) => $call(instanceMock, request, _Relationship, $kw({name: $mod("r%d", i), is_external: ext, target_part: part}))));
        _rels_prop_.returnValue = $d(Object.fromEntries($list($slice(all_rels, null, 3, null)).map((r: any) => [r.rId, r])));
        part_0_.rels = $d(Object.fromEntries($list($slice(all_rels, 3, 4, null)).map((r: any) => [r.rId, r])));
        part_1_.rels = $d(Object.fromEntries($list($slice(all_rels, 4, null, null)).map((r: any) => [r.rId, r])));
        package$ = $call(OpcPackage, null);
        rels = $set_(package$.iterRels());
        $assertEqual(rels, $set_(all_rels));
      } },
    { name: "it_provides_access_to_the_main_document_part", args: ["request"], fn: (self: any, request: any): any => {
        let package$: any, part_related_by_: any, presentation_part: any, presentation_part_: any;
        presentation_part_ = instanceMock(request, PresentationPart);
        part_related_by_ = $call(methodMock, request, OpcPackage, "part_related_by", $kw({return_value: presentation_part_}));
        package$ = $call(OpcPackage, null);
        presentation_part = package$.mainDocumentPart;
        part_related_by_.assertCalledOnceWith(package$, RT.OFFICE_DOCUMENT);
        $assert($is(presentation_part, presentation_part_));
      } },
    { name: "it_can_find_the_next_available_partname", args: ["request", "ns", "expected_n"], parametrize: [{ names: ["ns", "expected_n"], values: [[[], 1], [[1], 2], [[1, 2], 3], [[2, 4], 3], [[1, 4], 3]] }], fn: (self: any, request: any, ns: any, expected_n: any): any => {
        let PackURI_: any, next_partname: any, package$: any, partname: any, tmpl: any;
        tmpl = "/x%d.xml";
        $call(methodMock, request, OpcPackage, "iter_parts", $kw({return_value: $list(ns).map((n: any) => $call(instanceMock, request, Part, $kw({partname: $mod(tmpl, n)})))}));
        next_partname = $mod(tmpl, expected_n);
        PackURI_ = $call(classMock, request, "pptx.opc.package.PackURI", $kw({return_value: $call(PackURI, next_partname)}));
        package$ = $call(OpcPackage, null);
        partname = package$.nextPartname(tmpl);
        PackURI_.assertCalledOnceWith(next_partname);
        $assertEqual(partname, next_partname);
      } },
    { name: "it_can_save_to_a_pkg_file", args: ["request", "_rels_prop_", "relationships_"], fn: (self: any, request: any, _rels_prop_: any, relationships_: any): any => {
        let PackageWriter_: any, package$: any, parts_: any;
        _rels_prop_.returnValue = relationships_;
        parts_ = $list($list($range(3)).map((_: any) => instanceMock(request, Part)));
        $call(methodMock, request, OpcPackage, "iter_parts", $kw({return_value: $iter(parts_)}));
        PackageWriter_ = classMock(request, "pptx.opc.package.PackageWriter");
        package$ = $call(OpcPackage, null);
        package$.save("prs.pptx");
        PackageWriter_.write.assertCalledOnceWith("prs.pptx", relationships_, parts_);
      } },
    { name: "it_loads_the_pkg_file_to_help", args: ["request", "_rels_prop_", "relationships_"], fn: (self: any, request: any, _rels_prop_: any, relationships_: any): any => {
        let _PackageLoader_: any, package$: any, return_value: any;
        _PackageLoader_ = classMock(request, "pptx.opc.package._PackageLoader");
        _PackageLoader_.load.returnValue = ["pkg-rels-xml", $d({partname: "part"})];
        _rels_prop_.returnValue = relationships_;
        package$ = $call(OpcPackage, "prs.pptx");
        return_value = package$._load();
        _PackageLoader_.load.assertCalledOnceWith("prs.pptx", package$);
        relationships_.loadFromXml.assertCalledOnceWith(PACKAGE_URI, "pkg-rels-xml", $d({partname: "part"}));
        $assert($is(return_value, package$));
      } },
    { name: "it_constructs_its_relationships_object_to_help", args: ["request", "relationships_"], fn: (self: any, request: any, relationships_: any): any => {
        let _Relationships_: any, package$: any, rels: any;
        _Relationships_ = $call(classMock, request, "pptx.opc.package._Relationships", $kw({return_value: relationships_}));
        package$ = $call(OpcPackage, null);
        rels = package$._rels;
        _Relationships_.assertCalledOnceWith(PACKAGE_URI.baseURI);
        $assert($is(rels, relationships_));
      } },
  ],
};
pytestClass("DescribeOpcPackage", DescribeOpcPackage);

const Describe_PackageLoader: any = {
  fixtures: {
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, OpcPackage);
      } },
    _xml_rels_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _PackageLoader, "_xml_rels");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_load_interface_classmethod", args: ["request", "package_"], fn: (self: any, request: any, package_: any): any => {
        let _init_: any, _load_: any, parts: any, pkg_xml_rels: any, pkg_xml_rels_: any;
        _init_ = initializerMock(request, _PackageLoader);
        pkg_xml_rels_ = element("r:Relationships");
        _load_ = $call(methodMock, request, _PackageLoader, "_load", $kw({return_value: [pkg_xml_rels_, $d({partname: "part"})]}));
        [pkg_xml_rels, parts] = $list(_PackageLoader.load("prs.pptx", package_));
        _init_.assertCalledOnceWith(ANY, "prs.pptx", package_);
        _load_.assertCalledOnceWith(ANY);
        $assert($is(pkg_xml_rels, pkg_xml_rels_));
        $assertEqual(parts, $d({partname: "part"}));
      } },
    { name: "it_loads_the_package_to_help", args: ["request", "_xml_rels_prop_"], fn: (self: any, request: any, _xml_rels_prop_: any): any => {
        let package_loader: any, parts: any, parts_: any, pkg_xml_rels: any, rels_: any;
        parts_ = $d(Object.fromEntries($list($range(1, 4)).map((n: any) => [$mod("partname_%d", n), $call(instanceMock, request, Part, $kw({partname: $mod("partname_%d", n)}))])));
        $call(propertyMock, request, _PackageLoader, "_parts", $kw({return_value: parts_}));
        rels_ = $dict($itertools.chain([["/", instanceMock(request, _Relationships)]], $list($range(1, 4)).map((n: any) => [$mod("partname_%d", n), instanceMock(request, _Relationships)])));
        _xml_rels_prop_.returnValue = rels_;
        package_loader = $call(_PackageLoader, null, null);
        [pkg_xml_rels, parts] = $list(package_loader._load());
        for (const part_ of $list($values(parts_))) {
          part_.loadRelsFromXml.assertCalledOnceWith($get(rels_, part_.partname), parts_);
        }
        $assert($is(pkg_xml_rels, $get(rels_, "/")));
        $assert($is(parts, parts_));
      } },
    { name: "it_loads_the_xml_relationships_from_the_package_to_help", args: ["request"], fn: (self: any, request: any): any => {
        let _xml_rels_for_: any, core_xml_rels: any, package_loader: any, pkg_xml_rels: any, prs_xml_rels: any, slide_xml_rels: any, thumbnail_xml_rels: any, xml_rels: any;
        pkg_xml_rels = parseXml(snippetBytes("package-rels-xml"));
        prs_xml_rels = parseXml(snippetBytes("presentation-rels-xml"));
        slide_xml_rels = CT_Relationships.new();
        thumbnail_xml_rels = CT_Relationships.new();
        core_xml_rels = CT_Relationships.new();
        _xml_rels_for_ = $call(methodMock, request, _PackageLoader, "_xml_rels_for", $kw({side_effect: $iter([pkg_xml_rels, prs_xml_rels, slide_xml_rels, thumbnail_xml_rels, core_xml_rels])}));
        package_loader = $call(_PackageLoader, null, null);
        xml_rels = package_loader._xmlRels;
        $assertEqual(_xml_rels_for_.callArgsList, [call(package_loader, "/"), call(package_loader, "/ppt/presentation.xml"), call(package_loader, "/ppt/slides/slide1.xml"), call(package_loader, "/docProps/thumbnail.jpeg"), call(package_loader, "/docProps/core.xml")]);
        $assertEqual(xml_rels, $d({"/": pkg_xml_rels, "/ppt/presentation.xml": prs_xml_rels, "/ppt/slides/slide1.xml": slide_xml_rels, "/docProps/thumbnail.jpeg": thumbnail_xml_rels, "/docProps/core.xml": core_xml_rels}));
      } },
  ],
};
pytestClass("Describe_PackageLoader", Describe_PackageLoader);

const DescribePart: any = {
  fixtures: {
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, OpcPackage);
      } },
    relationships_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationships);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_be_constructed_by_PartFactory", args: ["request", "package_"], fn: (self: any, request: any, package_: any): any => {
        let _init_: any, part: any, partname_: any;
        partname_ = instanceMock(request, PackURI);
        _init_ = initializerMock(request, Part);
        part = Part.load(partname_, CT.PML_SLIDE, package_, $b("blob"));
        _init_.assertCalledOnceWith(part, partname_, CT.PML_SLIDE, package_, $b("blob"));
        $assert($isinstance(part, Part));
      } },
    { name: "it_uses_the_load_blob_as_its_blob", args: [], fn: (self: any): any => {
        $assertEqual($call(Part, null, null, null, $b("blob")).blob, $b("blob"));
      } },
    { name: "it_can_change_its_blob", args: [], fn: (self: any): any => {
        let part: any;
        part = $call(Part, null, null, null, $b("old-blob"));
        part.blob = $b("new-blob");
        $assertEqual(part.blob, $b("new-blob"));
      } },
    { name: "it_knows_its_content_type", args: [], fn: (self: any): any => {
        $assertEqual($call(Part, null, CT.PML_SLIDE, null).contentType, CT.PML_SLIDE);
      } },
    { name: "it_knows_the_package_it_belongs_to", args: ["package_"], fn: (self: any, package_: any): any => {
        $assert($is($call(Part, null, null, package_).package, package_));
      } },
    { name: "it_knows_its_partname", args: [], fn: (self: any): any => {
        $assertEqual($call(Part, $call(PackURI, "/part/name"), null, null).partname, $call(PackURI, "/part/name"));
      } },
    { name: "it_can_change_its_partname", args: [], fn: (self: any): any => {
        let part: any;
        part = $call(Part, $call(PackURI, "/old/part/name"), null, null);
        part.partname = $call(PackURI, "/new/part/name");
        $assertEqual(part.partname, $call(PackURI, "/new/part/name"));
      } },
    { name: "it_provides_access_to_its_relationships_for_traversal", args: ["request", "relationships_"], fn: (self: any, request: any, relationships_: any): any => {
        $call(propertyMock, request, Part, "_rels", $kw({return_value: relationships_}));
        $assert($is($call(Part, null, null, null).rels, relationships_));
      } },
    { name: "it_can_load_a_blob_from_a_file_path_to_help", args: [], fn: (self: any): any => {
        let f: any, file_bytes: any, part: any, path: any;
        path = absjoin(testFileDir, "minimal.pptx");
        const $cm1 = $open(path, "rb");
        f = $enter($cm1);
        try {
          file_bytes = $read(f);
        } finally {
          $exit($cm1);
        }
        part = $call(Part, null, null, null, null);
        $assertEqual(part._blobFromFile(path), file_bytes);
      } },
    { name: "it_can_load_a_blob_from_a_file_like_object_to_help", args: [], fn: (self: any): any => {
        let part: any;
        part = $call(Part, null, null, null, null);
        $assertEqual(part._blobFromFile(new BytesIO($b("012345"))), $b("012345"));
      } },
    { name: "it_constructs_its_relationships_object_to_help", args: ["request", "relationships_"], fn: (self: any, request: any, relationships_: any): any => {
        let _Relationships_: any, part: any, rels: any;
        _Relationships_ = $call(classMock, request, "pptx.opc.package._Relationships", $kw({return_value: relationships_}));
        part = $call(Part, $call(PackURI, "/ppt/slides/slide1.xml"), null, null);
        rels = part._rels;
        _Relationships_.assertCalledOnceWith("/ppt/slides");
        $assert($is(rels, relationships_));
      } },
  ],
};
pytestClass("DescribePart", DescribePart);

const DescribeXmlPart: any = {
  fixtures: {
    relationships_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationships);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_be_constructed_by_PartFactory", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, element_: any, package_: any, parse_xml_: any, part: any, partname: any;
        partname = $call(PackURI, "/ppt/slides/slide1.xml");
        element_ = element("p:sld");
        package_ = instanceMock(request, OpcPackage);
        parse_xml_ = $call(functionMock, request, "pptx.opc.package.parse_xml", $kw({return_value: element_}));
        _init_ = initializerMock(request, XmlPart);
        part = XmlPart.load(partname, CT.PML_SLIDE, package_, $b("blob"));
        parse_xml_.assertCalledOnceWith($b("blob"));
        _init_.assertCalledOnceWith(part, partname, CT.PML_SLIDE, package_, element_);
        $assert($isinstance(part, XmlPart));
      } },
    { name: "it_can_serialize_to_xml", args: ["request"], fn: (self: any, request: any): any => {
        let blob: any, element_: any, serialize_part_xml_: any, xml_part: any;
        element_ = element("p:sld");
        serialize_part_xml_ = functionMock(request, "pptx.opc.package.serialize_part_xml");
        xml_part = $call(XmlPart, null, null, null, element_);
        blob = xml_part.blob;
        serialize_part_xml_.assertCalledOnceWith(element_);
        $assert($is(blob, serialize_part_xml_.returnValue));
      } },
    { name: "it_can_drop_a_relationship", args: ["request", "relationships_", "ref_count", "calls"], parametrize: [{ names: ["ref_count", "calls"], values: [[2, []], [1, [call("rId42")]]] }], fn: (self: any, request: any, relationships_: any, ref_count: any, calls: any): any => {
        let _rel_ref_count_: any, part: any;
        _rel_ref_count_ = $call(methodMock, request, XmlPart, "_rel_ref_count", $kw({return_value: ref_count}));
        $call(propertyMock, request, XmlPart, "_rels", $kw({return_value: relationships_}));
        part = $call(XmlPart, null, null, null, null);
        part.dropRel("rId42");
        _rel_ref_count_.assertCalledOnceWith(part, "rId42");
        $assertEqual(relationships_.pop.callArgsList, calls);
      } },
    { name: "it_knows_it_is_the_part_for_its_child_objects", args: [], fn: (self: any): any => {
        let xml_part: any;
        xml_part = $call(XmlPart, null, null, null, null);
        $assert($is(xml_part.part, xml_part));
      } },
  ],
};
pytestClass("DescribeXmlPart", DescribeXmlPart);

const DescribePartFactory: any = {
  fixtures: {
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, OpcPackage);
      } },
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_custom_part_type_for_registered_content_types", args: ["request", "package_", "part_"], fn: (self: any, request: any, package_: any, part_: any): any => {
        let SlidePart_: any, part: any, partname: any;
        SlidePart_ = classMock(request, "pptx.opc.package.XmlPart");
        SlidePart_.load.returnValue = part_;
        partname = $call(PackURI, "/ppt/slides/slide7.xml");
        $set(PartFactory.partTypeFor, CT.PML_SLIDE, SlidePart_);
        part = $call(PartFactory, partname, CT.PML_SLIDE, package_, $b("blob"));
        SlidePart_.load.assertCalledOnceWith(partname, CT.PML_SLIDE, package_, $b("blob"));
        $assert($is(part, part_));
      } },
    { name: "it_constructs_part_using_default_class_when_no_custom_registered", args: ["request", "package_", "part_"], fn: (self: any, request: any, package_: any, part_: any): any => {
        let Part_: any, part: any, partname: any;
        Part_ = classMock(request, "pptx.opc.package.Part");
        Part_.load.returnValue = part_;
        partname = $call(PackURI, "/bar/foo.xml");
        part = $call(PartFactory, partname, CT.OFC_VML_DRAWING, package_, $b("blob"));
        Part_.load.assertCalledOnceWith(partname, CT.OFC_VML_DRAWING, package_, $b("blob"));
        $assert($is(part, part_));
      } },
  ],
};
pytestClass("DescribePartFactory", DescribePartFactory);

const Describe_ContentTypeMap: any = {
  fixtures: {
    content_type_map: { args: [], scope: "class", fn: (self: any): any => {
        return _ContentTypeMap.fromXml(testfileBytes("expanded_pptx", "[Content_Types].xml"));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_content_types_xml", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, content_types_xml: any, ct_map: any;
        _init_ = initializerMock(request, _ContentTypeMap);
        content_types_xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Default Extension=\"PNG\" ContentType=\"image/png\"/>\n  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>\n</Types>\n";
        ct_map = _ContentTypeMap.fromXml(content_types_xml);
        _init_.assertCalledOnceWith(ct_map, $d({"/ppt/presentation.xml": CT.PML_PRESENTATION_MAIN}), $d({png: CT.PNG, xml: CT.XML}));
      } },
    { name: "it_matches_an_override_on_case_insensitive_partname", args: ["content_type_map", "partname", "expected_value"], parametrize: [{ names: ["partname", "expected_value"], values: [["/docProps/core.xml", CT.OPC_CORE_PROPERTIES], ["/ppt/presentation.xml", CT.PML_PRESENTATION_MAIN], ["/PPT/Presentation.XML", CT.PML_PRESENTATION_MAIN], ["/ppt/viewprops.xml", CT.PML_VIEW_PROPS]] }], fn: (self: any, content_type_map: any, partname: any, expected_value: any): any => {
        $assertEqual($get(content_type_map, $call(PackURI, partname)), expected_value);
      } },
    { name: "it_falls_back_to_case_insensitive_extension_default_match", args: ["content_type_map", "partname", "expected_value"], parametrize: [{ names: ["partname", "expected_value"], values: [["/foo/bar.xml", CT.XML], ["/FOO/BAR.Rels", CT.OPC_RELATIONSHIPS], ["/foo/bar.jpeg", CT.JPEG]] }], fn: (self: any, content_type_map: any, partname: any, expected_value: any): any => {
        $assertEqual($get(content_type_map, $call(PackURI, partname)), expected_value);
      } },
    { name: "it_raises_KeyError_on_partname_not_found", args: ["content_type_map"], fn: (self: any, content_type_map: any): any => {
        let e: any;
        e = $raises(KeyError, () => {
          $get(content_type_map, $call(PackURI, "/!blat/rhumba.1x&"));
        });
        $assertEqual($str(e.value), "\"no content-type for partname '/!blat/rhumba.1x&' in [Content_Types].xml\"");
      } },
    { name: "it_raises_TypeError_on_key_not_instance_of_PackURI", args: ["content_type_map"], fn: (self: any, content_type_map: any): any => {
        let e: any;
        e = $raises(TypeError, () => {
          $get(content_type_map, "/part/name1.xml");
        });
        $assertEqual($str(e.value), "_ContentTypeMap key must be <type 'PackURI'>, got str");
      } },
  ],
};
pytestClass("Describe_ContentTypeMap", Describe_ContentTypeMap);

const Describe_Relationships: any = {
  fixtures: {
    _add_relationship_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _Relationships, "_add_relationship");
      } },
    _get_matching_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _Relationships, "_get_matching");
      } },
    _next_rId_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Relationships, "_next_rId");
      } },
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
    _Relationship_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.opc.package._Relationship");
      } },
    relationship_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Relationship);
      } },
    _rels_by_reltype_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Relationships, "_rels_by_reltype");
      } },
    _rels_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Relationships, "_rels");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_contains_a_relationship_with_rId", args: ["_rels_prop_", "rId", "expected_value"], parametrize: [{ names: ["rId", "expected_value"], values: [["rId1", true], ["rId2", false]] }], fn: (self: any, _rels_prop_: any, rId: any, expected_value: any): any => {
        _rels_prop_.returnValue = $d({rId1: null});
        $assert($is($in(rId, $call(_Relationships, null)), expected_value));
      } },
    { name: "it_has_dict_style_lookup_of_rel_by_rId", args: ["_rels_prop_", "relationship_"], fn: (self: any, _rels_prop_: any, relationship_: any): any => {
        _rels_prop_.returnValue = $d({rId17: relationship_});
        $assert($is($get($call(_Relationships, null), "rId17"), relationship_));
      } },
    { name: "but_it_raises_KeyError_when_no_relationship_has_rId", args: ["_rels_prop_"], fn: (self: any, _rels_prop_: any): any => {
        let e: any;
        _rels_prop_.returnValue = $d({});
        e = $raises(KeyError, () => {
          $get($call(_Relationships, null), "rId6");
        });
        $assertEqual($str(e.value), "\"no relationship with key 'rId6'\"");
      } },
    { name: "it_can_iterate_the_rIds_of_the_relationships_it_contains", args: ["request", "_rels_prop_"], fn: (self: any, request: any, _rels_prop_: any): any => {
        let relationships: any, rels_: any;
        rels_ = $set_($list($range(5)).map((n: any) => instanceMock(request, _Relationship)));
        _rels_prop_.returnValue = $d(Object.fromEntries($list($enumerate(rels_)).map(([i, r]: any) => [$mod("rId%d", (i + 1)), r])));
        relationships = $call(_Relationships, null);
        for (const rId of $list(relationships)) {
          $remove(rels_, $get(relationships, rId));
        }
        $assertEqual($len(rels_), 0);
      } },
    { name: "it_has_a_len", args: ["_rels_prop_"], fn: (self: any, _rels_prop_: any): any => {
        _rels_prop_.returnValue = $d({a: 0, b: 1});
        $assertEqual($len($call(_Relationships, null)), 2);
      } },
    { name: "it_can_add_a_relationship_to_a_target_part", args: ["part_", "_get_matching_", "_add_relationship_"], fn: (self: any, part_: any, _get_matching_: any, _add_relationship_: any): any => {
        let rId: any, relationships: any;
        _get_matching_.returnValue = null;
        _add_relationship_.returnValue = "rId7";
        relationships = $call(_Relationships, null);
        rId = relationships.getOrAdd(RT.IMAGE, part_);
        _get_matching_.assertCalledOnceWith(relationships, RT.IMAGE, part_);
        _add_relationship_.assertCalledOnceWith(relationships, RT.IMAGE, part_);
        $assertEqual(rId, "rId7");
      } },
    { name: "but_it_returns_an_existing_relationship_if_it_matches", args: ["part_", "_get_matching_"], fn: (self: any, part_: any, _get_matching_: any): any => {
        let rId: any, relationships: any;
        _get_matching_.returnValue = "rId3";
        relationships = $call(_Relationships, null);
        rId = relationships.getOrAdd(RT.IMAGE, part_);
        _get_matching_.assertCalledOnceWith(relationships, RT.IMAGE, part_);
        $assertEqual(rId, "rId3");
      } },
    { name: "it_can_add_an_external_relationship_to_a_URI", args: ["_get_matching_", "_add_relationship_"], fn: (self: any, _get_matching_: any, _add_relationship_: any): any => {
        let rId: any, relationships: any;
        _get_matching_.returnValue = null;
        _add_relationship_.returnValue = "rId2";
        relationships = $call(_Relationships, null);
        rId = relationships.getOrAddExtRel(RT.HYPERLINK, "http://url");
        _get_matching_.assertCalledOnceWith(relationships, RT.HYPERLINK, "http://url", $kw({is_external: true}));
        _add_relationship_.assertCalledOnceWith(relationships, RT.HYPERLINK, "http://url", $kw({is_external: true}));
        $assertEqual(rId, "rId2");
      } },
    { name: "but_it_returns_an_existing_external_relationship_if_it_matches", args: ["part_", "_get_matching_"], fn: (self: any, part_: any, _get_matching_: any): any => {
        let rId: any, relationships: any;
        _get_matching_.returnValue = "rId10";
        relationships = $call(_Relationships, null);
        rId = relationships.getOrAddExtRel(RT.HYPERLINK, "http://url");
        _get_matching_.assertCalledOnceWith(relationships, RT.HYPERLINK, "http://url", $kw({is_external: true}));
        $assertEqual(rId, "rId10");
      } },
    { name: "it_can_load_from_the_xml_in_a_rels_part", args: ["request", "_Relationship_", "part_"], fn: (self: any, request: any, _Relationship_: any, part_: any): any => {
        let parts: any, relationships: any, rels_: any, xml_rels: any;
        rels_ = $list($list($range(5)).map((i: any) => $call(instanceMock, request, _Relationship, $kw({rId: $mod("rId%d", (i + 1))}))));
        _Relationship_.fromXml.sideEffect = $iter(rels_);
        parts = $d({"/ppt/slideLayouts/slideLayout1.xml": part_});
        xml_rels = parseXml(snippetBytes("rels-load-from-xml"));
        relationships = $call(_Relationships, null);
        relationships.loadFromXml("/ppt/slides", xml_rels, parts);
        $assertEqual(_Relationship_.fromXml.callArgsList, [call("/ppt/slides", $get(xml_rels, 0), parts), call("/ppt/slides", $get(xml_rels, 1), parts)]);
        $assertEqual(relationships._rels, $d({rId1: $get(rels_, 0), rId2: $get(rels_, 1)}));
      } },
    { name: "it_can_find_a_part_with_reltype", args: ["_rels_by_reltype_prop_", "relationship_", "part_"], fn: (self: any, _rels_by_reltype_prop_: any, relationship_: any, part_: any): any => {
        let relationships: any;
        relationship_.targetPart = part_;
        _rels_by_reltype_prop_.returnValue = $collections.defaultdict(Array, [[RT.SLIDE_LAYOUT, [relationship_]]]);
        relationships = $call(_Relationships, null);
        $assert($is(relationships.partWithReltype(RT.SLIDE_LAYOUT), part_));
      } },
    { name: "but_it_raises_KeyError_when_there_is_no_such_part", args: ["_rels_by_reltype_prop_"], fn: (self: any, _rels_by_reltype_prop_: any): any => {
        let e: any, relationships: any;
        _rels_by_reltype_prop_.returnValue = $collections.defaultdict(Array);
        relationships = $call(_Relationships, null);
        e = $raises(KeyError, () => {
          relationships.partWithReltype(RT.SLIDE_LAYOUT);
        });
        $assertEqual($str(e.value), "\"no relationship of type 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' in collection\"");
      } },
    { name: "and_it_raises_ValueError_when_there_is_more_than_one_part_with_reltype", args: ["_rels_by_reltype_prop_", "relationship_", "part_"], fn: (self: any, _rels_by_reltype_prop_: any, relationship_: any, part_: any): any => {
        let e: any, relationships: any;
        relationship_.targetPart = part_;
        _rels_by_reltype_prop_.returnValue = $collections.defaultdict(Array, [[RT.SLIDE_LAYOUT, [relationship_, relationship_]]]);
        relationships = $call(_Relationships, null);
        e = $raises(ValueError, () => {
          relationships.partWithReltype(RT.SLIDE_LAYOUT);
        });
        $assertEqual($str(e.value), "multiple relationships of type 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' in collection");
      } },
    { name: "it_can_pop_a_relationship_to_remove_it_from_the_collection", args: ["_rels_prop_", "relationship_"], fn: (self: any, _rels_prop_: any, relationship_: any): any => {
        let relationships: any;
        _rels_prop_.returnValue = $d({rId22: relationship_});
        relationships = $call(_Relationships, null);
        relationships.pop("rId22");
        $assertEqual(relationships._rels, $d({}));
      } },
    { name: "it_can_serialize_itself_to_XML", args: ["request", "_rels_prop_"], fn: (self: any, request: any, _rels_prop_: any): any => {
        let relationships: any;
        _rels_prop_.returnValue = $d({rId11: $call(instanceMock, request, _Relationship, $kw({rId: "rId11", reltype: RT.SLIDE, target_ref: "../slides/slide1.xml", is_external: false})), rId2: $call(instanceMock, request, _Relationship, $kw({rId: "rId2", reltype: RT.HYPERLINK, target_ref: "http://url", is_external: true})), foo7W: $call(instanceMock, request, _Relationship, $kw({rId: "foo7W", reltype: RT.IMAGE, target_ref: "../media/image1.png", is_external: false}))});
        relationships = $call(_Relationships, null);
        $assertEqual(relationships.xml, snippetBytes("relationships"));
      } },
    { name: "it_can_add_a_relationship_to_a_part_to_help", args: ["request", "_next_rId_prop_", "_Relationship_", "relationship_", "_rels_prop_", "part_"], fn: (self: any, request: any, _next_rId_prop_: any, _Relationship_: any, relationship_: any, _rels_prop_: any, part_: any): any => {
        let rId: any, relationships: any;
        _next_rId_prop_.returnValue = "rId8";
        _Relationship_.returnValue = relationship_;
        _rels_prop_.returnValue = $d({});
        relationships = $call(_Relationships, "/ppt");
        rId = relationships._addRelationship(RT.SLIDE, part_);
        _Relationship_.assertCalledOnceWith("/ppt", "rId8", RT.SLIDE, $kw({target_mode: RTM.INTERNAL, target: part_}));
        $assertEqual(relationships._rels, $d({rId8: relationship_}));
        $assertEqual(rId, "rId8");
      } },
    { name: "and_it_can_add_an_external_relationship_to_help", args: ["request", "_next_rId_prop_", "_rels_prop_", "_Relationship_", "relationship_"], fn: (self: any, request: any, _next_rId_prop_: any, _rels_prop_: any, _Relationship_: any, relationship_: any): any => {
        let rId: any, relationships: any;
        _next_rId_prop_.returnValue = "rId9";
        _Relationship_.returnValue = relationship_;
        _rels_prop_.returnValue = $d({});
        relationships = $call(_Relationships, "/ppt");
        rId = $m(relationships, "_addRelationship", RT.HYPERLINK, "http://url", $kw({is_external: true}));
        _Relationship_.assertCalledOnceWith("/ppt", "rId9", RT.HYPERLINK, $kw({target_mode: RTM.EXTERNAL, target: "http://url"}));
        $assertEqual(relationships._rels, $d({rId9: relationship_}));
        $assertEqual(rId, "rId9");
      } },
    { name: "it_can_get_a_matching_relationship_to_help", args: ["request", "_rels_by_reltype_prop_", "target_ref", "is_external", "expected_value"], parametrize: [{ names: ["target_ref", "is_external", "expected_value"], values: [["http://url", true, "rId1"], ["part_1", false, "rId2"], ["http://foo", true, "rId3"], ["part_2", false, "rId4"], ["http://bar", true, null]] }], fn: (self: any, request: any, _rels_by_reltype_prop_: any, target_ref: any, is_external: any, expected_value: any): any => {
        let matching: any, part_1: any, part_2: any, relationships: any, target: any;
        [part_1, part_2] = $list($list($range(2)).map((_: any) => instanceMock(request, Part)));
        _rels_by_reltype_prop_.returnValue = $d({[RT.SLIDE]: $list([["rId1", null, "http://url", true], ["rId2", part_1, "/ppt/foo.bar", false], ["rId3", null, "http://foo", true], ["rId4", part_2, "/ppt/bar.foo", false]]).map(([rId, target_part, ref, external]: any) => $call(instanceMock, request, _Relationship, $kw({rId: rId, target_part: target_part, target_ref: ref, is_external: external})))});
        target = ($truthy(is_external) ? target_ref : ($truthy($eq(target_ref, "part_1")) ? part_1 : part_2));
        relationships = $call(_Relationships, null);
        matching = relationships._getMatching(RT.SLIDE, target, is_external);
        $assertEqual(matching, expected_value);
      } },
    { name: "but_it_returns_None_when_there_is_no_matching_relationship", args: ["_rels_by_reltype_prop_"], fn: (self: any, _rels_by_reltype_prop_: any): any => {
        let relationships: any;
        _rels_by_reltype_prop_.returnValue = $collections.defaultdict(Array);
        relationships = $call(_Relationships, null);
        $assert($is(relationships._getMatching(RT.HYPERLINK, "http://url", true), null));
      } },
    { name: "it_finds_the_next_rId_to_help", args: ["_rels_prop_", "rIds", "expected_value"], parametrize: [{ names: ["rIds", "expected_value"], values: [[[], "rId1"], [["rId1"], "rId2"], [["rId1", "rId2"], "rId3"], [["rId1", "rId4"], "rId3"], [["rId1", "rId4", "rId6"], "rId3"], [["rId1", "rId2", "rId6"], "rId4"]] }], fn: (self: any, _rels_prop_: any, rIds: any, expected_value: any): any => {
        let relationships: any;
        _rels_prop_.returnValue = $d(Object.fromEntries($list(rIds).map((rId: any) => [rId, null])));
        relationships = $call(_Relationships, null);
        $assertEqual(relationships._nextRId, expected_value);
      } },
    { name: "it_collects_relationships_by_reltype_to_help", args: ["request", "_rels_prop_"], fn: (self: any, request: any, _rels_prop_: any): any => {
        let relationships: any, rels: any, rels_by_reltype: any;
        rels = $d(Object.fromEntries($list($enumerate([RT.SLIDE, RT.IMAGE, RT.SLIDE, RT.HYPERLINK])).map(([i, reltype]: any) => [$mod("rId%d", (i + 1)), $call(instanceMock, request, _Relationship, $kw({reltype: reltype}))])));
        _rels_prop_.returnValue = rels;
        relationships = $call(_Relationships, null);
        rels_by_reltype = relationships._relsByReltype;
        $assert($in($get(rels, "rId1"), $get(rels_by_reltype, RT.SLIDE)));
        $assert($in($get(rels, "rId2"), $get(rels_by_reltype, RT.IMAGE)));
        $assert($in($get(rels, "rId3"), $get(rels_by_reltype, RT.SLIDE)));
        $assert($in($get(rels, "rId4"), $get(rels_by_reltype, RT.HYPERLINK)));
        $assertEqual($dictget(rels_by_reltype, RT.CHART, []), []);
      } },
  ],
};
pytestClass("Describe_Relationships", Describe_Relationships);

const Describe_Relationship: any = {
  fixtures: {
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_xml", args: ["request", "part_"], fn: (self: any, request: any, part_: any): any => {
        let _init_: any, parts: any, rel_elm: any, relationship: any;
        _init_ = initializerMock(request, _Relationship);
        rel_elm = $call(instanceMock, request, CT_Relationship, $kw({rId: "rId42", reltype: RT.SLIDE, targetMode: RTM.INTERNAL, target_ref: "slides/slide7.xml"}));
        parts = $d({"/ppt/slides/slide7.xml": part_});
        relationship = _Relationship.fromXml("/ppt", rel_elm, parts);
        _init_.assertCalledOnceWith(relationship, "/ppt", "rId42", RT.SLIDE, RTM.INTERNAL, part_);
        $assert($isinstance(relationship, _Relationship));
      } },
    { name: "it_knows_whether_it_is_external", args: ["target_mode", "expected_value"], parametrize: [{ names: ["target_mode", "expected_value"], values: [[RTM.INTERNAL, false], [RTM.EXTERNAL, true], [null, false]] }], fn: (self: any, target_mode: any, expected_value: any): any => {
        let relationship: any;
        relationship = $call(_Relationship, null, null, null, target_mode, null);
        $assertEqual(relationship.isExternal, expected_value);
      } },
    { name: "it_knows_its_relationship_type", args: [], fn: (self: any): any => {
        let relationship: any;
        relationship = $call(_Relationship, null, null, RT.SLIDE, null, null);
        $assertEqual(relationship.reltype, RT.SLIDE);
      } },
    { name: "it_knows_its_rId", args: [], fn: (self: any): any => {
        let relationship: any;
        relationship = $call(_Relationship, null, "rId42", null, null, null);
        $assertEqual(relationship.rId, "rId42");
      } },
    { name: "it_provides_access_to_its_target_part", args: ["part_"], fn: (self: any, part_: any): any => {
        let relationship: any;
        relationship = $call(_Relationship, null, null, null, RTM.INTERNAL, part_);
        $assert($is(relationship.targetPart, part_));
      } },
    { name: "but_it_raises_ValueError_on_target_part_for_external_rel", args: [], fn: (self: any): any => {
        let e: any, relationship: any;
        relationship = $call(_Relationship, null, null, null, RTM.EXTERNAL, null);
        e = $raises(ValueError, () => {
          relationship.targetPart;
        });
        $assertEqual($str(e.value), "`.target_part` property on _Relationship is undefined when target-mode is external");
      } },
    { name: "it_knows_its_target_partname", args: ["part_"], fn: (self: any, part_: any): any => {
        let relationship: any;
        part_.partname = $call(PackURI, "/ppt/slideLayouts/slideLayout4.xml");
        relationship = $call(_Relationship, null, null, null, RTM.INTERNAL, part_);
        $assertEqual(relationship.targetPartname, "/ppt/slideLayouts/slideLayout4.xml");
      } },
    { name: "but_it_raises_ValueError_on_target_partname_for_external_rel", args: [], fn: (self: any): any => {
        let e: any, relationship: any;
        relationship = $call(_Relationship, null, null, null, RTM.EXTERNAL, null);
        e = $raises(ValueError, () => {
          relationship.targetPartname;
        });
        $assertEqual($str(e.value), "`.target_partname` property on _Relationship is undefined when target-mode is external");
      } },
    { name: "it_knows_the_target_uri_for_an_external_rel", args: [], fn: (self: any): any => {
        let relationship: any;
        relationship = $call(_Relationship, null, null, null, RTM.EXTERNAL, "http://url");
        $assertEqual(relationship.targetRef, "http://url");
      } },
    { name: "and_it_knows_the_relative_partname_for_an_internal_rel", args: ["request"], fn: (self: any, request: any): any => {
        let relationship: any;
        $call(propertyMock, request, _Relationship, "target_partname", $kw({return_value: $call(PackURI, "/ppt/media/image1.png")}));
        relationship = $call(_Relationship, "/ppt/slides", null, null, null, null);
        $assertEqual(relationship.targetRef, "../media/image1.png");
      } },
  ],
};
pytestClass("Describe_Relationship", Describe_Relationship);

