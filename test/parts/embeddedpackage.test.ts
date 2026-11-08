// Machine-translated from python-pptx tests/parts/test_embeddedpackage.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $b, $call, $is, $isinstance, $kw, $mod, ANY, classMock, initializerMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { PROG_ID as PROG_ID$imported } from "../../src/enum/shapes.js";
import { CONTENT_TYPE as CT$imported } from "../../src/opc/constants.js";
import { OpcPackage as OpcPackage$imported, PackURI as PackURI$imported } from "../../src/opc/package.js";
import { EmbeddedDocxPart as EmbeddedDocxPart$imported, EmbeddedPackagePart, EmbeddedPptxPart as EmbeddedPptxPart$imported, EmbeddedXlsxPart } from "../../src/parts/embeddedpackage.js";
// python binds imported names at import time; patches made later don't change them
const PROG_ID = PROG_ID$imported;
const CT = CT$imported;
const OpcPackage = OpcPackage$imported;
const PackURI = PackURI$imported;
const EmbeddedDocxPart = EmbeddedDocxPart$imported;
const EmbeddedPptxPart = EmbeddedPptxPart$imported;
void $py;


const DescribeEmbeddedPackagePart: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_factory_that_creates_a_package_part_for_MS_Office_files", args: ["request", "prog_id", "EmbeddedPartCls"], parametrize: [{ names: ["prog_id", "EmbeddedPartCls"], values: [[PROG_ID.DOCX, EmbeddedDocxPart], [PROG_ID.PPTX, EmbeddedPptxPart], [PROG_ID.XLSX, EmbeddedXlsxPart]] }], fn: (self: any, request: any, prog_id: any, EmbeddedPartCls: any): any => {
        let EmbeddedPartCls_: any, embedded_object_part_: any, object_blob_: any, ole_object_part: any, package_: any;
        object_blob_ = $b("0123456789");
        package_ = instanceMock(request, OpcPackage);
        embedded_object_part_ = instanceMock(request, EmbeddedPartCls);
        EmbeddedPartCls_ = classMock(request, $mod("pptx.parts.embeddedpackage.%s", EmbeddedPartCls.name));
        EmbeddedPartCls_.new.returnValue = embedded_object_part_;
        ole_object_part = EmbeddedPackagePart.factory(prog_id, object_blob_, package_);
        EmbeddedPartCls_.new.assertCalledOnceWith(object_blob_, package_);
        $assert($is(ole_object_part, embedded_object_part_));
      } },
    { name: "but_it_creates_a_generic_object_part_for_non_MS_Office_files", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, object_blob_: any, ole_object_part: any, package_: any, partname_: any, progId: any;
        progId = "Foo.Bar.42";
        object_blob_ = $b("0123456789");
        package_ = instanceMock(request, OpcPackage);
        _init_ = $call(initializerMock, request, EmbeddedPackagePart, $kw({autospec: true}));
        partname_ = instanceMock(request, PackURI);
        package_.nextPartname.returnValue = partname_;
        ole_object_part = EmbeddedPackagePart.factory(progId, object_blob_, package_);
        package_.nextPartname.assertCalledOnceWith("/ppt/embeddings/oleObject%d.bin");
        _init_.assertCalledOnceWith(ANY, partname_, CT.OFC_OLE_OBJECT, package_, object_blob_);
        $assert($isinstance(ole_object_part, EmbeddedPackagePart));
      } },
    { name: "it_provides_a_contructor_classmethod_for_subclasses", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, blob_: any, package_: any, partname_: any, xlsx_part: any;
        blob_ = $b("0123456789");
        package_ = instanceMock(request, OpcPackage);
        _init_ = $call(initializerMock, request, EmbeddedXlsxPart, $kw({autospec: true}));
        partname_ = instanceMock(request, PackURI);
        package_.nextPartname.returnValue = partname_;
        xlsx_part = EmbeddedXlsxPart.new(blob_, package_);
        package_.nextPartname.assertCalledOnceWith(EmbeddedXlsxPart.partnameTemplate);
        _init_.assertCalledOnceWith(xlsx_part, partname_, EmbeddedXlsxPart.contentType, package_, blob_);
        $assert($isinstance(xlsx_part, EmbeddedXlsxPart));
      } },
  ],
};
pytestClass("DescribeEmbeddedPackagePart", DescribeEmbeddedPackagePart);

