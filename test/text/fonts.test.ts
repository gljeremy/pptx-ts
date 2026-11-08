// Machine-translated from python-pptx tests/text/test_fonts.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { calcsize } from "../../src/text/struct.js";
import { $assert, $assertEqual, $b, $call, $cm1, $d, $encode, $enter, $exit, $find, $get, $is, $isinstance, $iter, $kw, $len, $list, $path, $read, $slice, BytesIO, call, classMock, functionMock, initializerMock, instanceMock, methodMock, openMock, propertyMock, pytestClass, varMock } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { FontFiles as FontFiles$imported, _BaseTable as _BaseTable$imported, _Font as _Font$imported, _HeadTable as _HeadTable$imported, _NameTable as _NameTable$imported, _Stream, _TableFactory as _TableFactory$imported } from "../../src/text/fonts.js";
import { testFileDir, testfile } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const FontFiles = FontFiles$imported;
const _BaseTable = _BaseTable$imported;
const _Font = _Font$imported;
const _HeadTable = _HeadTable$imported;
const _NameTable = _NameTable$imported;
const _TableFactory = _TableFactory$imported;
void $py;


const DescribeFontFiles: any = {
  fixtures: {
    find_fixture: { args: ["request", "_installed_fonts_"], params: [["Foobar", false, false, "foobar.ttf"], ["Foobar", true, false, "foobarb.ttf"], ["Barfoo", false, true, "barfooi.ttf"]], fn: (self: any, request: any, _installed_fonts_: any): any => {
        let expected_path: any, family_name: any, is_bold: any, is_italic: any;
        [family_name, is_bold, is_italic, expected_path] = $list(request.param);
        return [family_name, is_bold, is_italic, expected_path];
      } },
    font_dirs_fixture: { args: ["request", "_os_x_font_directories_", "_windows_font_directories_"], params: [["darwin", ["a", "b"]], ["win32", ["c", "d"]]], fn: (self: any, request: any, _os_x_font_directories_: any, _windows_font_directories_: any): any => {
        let dirs_meth_mock: any, expected_dirs: any, platform: any, sys_: any;
        [platform, expected_dirs] = $list(request.param);
        dirs_meth_mock = $get($d({darwin: _os_x_font_directories_, win32: _windows_font_directories_}), platform);
        sys_ = varMock(request, "pptx.text.fonts.sys");
        sys_.platform = platform;
        dirs_meth_mock.returnValue = expected_dirs;
        return expected_dirs;
      } },
    installed_fixture: { args: ["_iter_font_files_in_", "_font_directories_"], fn: (self: any, _iter_font_files_in_: any, _font_directories_: any): any => {
        let expected_call_args: any, expected_values: any;
        _font_directories_.returnValue = ["d", "d_2"];
        _iter_font_files_in_.sideEffect = [[[["A", true, false], "a.ttf"]], [[["B", false, true], "b.ttf"]]];
        expected_call_args = [call("d"), call("d_2")];
        expected_values = $d({[["A", true, false]]: "a.ttf", [["B", false, true]]: "b.ttf"});
        return [expected_call_args, expected_values];
      } },
    iter_fixture: { args: ["_Font_"], fn: (self: any, _Font_: any): any => {
        let directory: any, expected_calls: any, expected_paths: any, font: any, font_file_path: any;
        directory = testFileDir;
        font_file_path = testfile("calibriz.ttf");
        font = _Font_.open.returnValue.__enter__.returnValue;
        [font.familyName, font.isBold, font.isItalic] = $list(["Arial", true, true]);
        expected_calls = [call(font_file_path)];
        expected_paths = [[["Arial", true, true], font_file_path]];
        return [directory, _Font_, expected_calls, expected_paths];
      } },
    osx_dirs_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let os_: any;
        os_ = varMock(request, "pptx.text.fonts.os");
        os_.path = $path;
        os_.environ = $d({HOME: "/Users/fbar"});
        return ["/Library/Fonts", "/Network/Library/Fonts", "/System/Library/Fonts", "/Users/fbar/Library/Fonts", "/Users/fbar/.fonts"];
      } },
    win_dirs_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        return ["C:\\Windows\\Fonts"];
      } },
    _Font_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.fonts._Font");
      } },
    _font_directories_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FontFiles, "_font_directories", $kw({autospec: false}));
      } },
    _installed_fonts_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FontFiles, "_installed_fonts", $kw({autospec: false, return_value: $d({[["Foobar", false, false]]: "foobar.ttf", [["Foobar", true, false]]: "foobarb.ttf", [["Barfoo", false, true]]: "barfooi.ttf"})}));
      } },
    _iter_font_files_in_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FontFiles, "_iter_font_files_in", $kw({autospec: false}));
      } },
    _os_x_font_directories_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FontFiles, "_os_x_font_directories", $kw({autospec: false}));
      } },
    _windows_font_directories_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FontFiles, "_windows_font_directories", $kw({autospec: false}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_find_a_system_font_file", args: ["find_fixture"], fn: (self: any, find_fixture: any): any => {
        let expected_path: any, family_name: any, is_bold: any, is_italic: any, path: any;
        [family_name, is_bold, is_italic, expected_path] = $list(find_fixture);
        path = $find(FontFiles, family_name, is_bold, is_italic);
        $assertEqual(path, expected_path);
      } },
    { name: "it_catalogs_the_system_fonts_to_help_find", args: ["installed_fixture"], fn: (self: any, installed_fixture: any): any => {
        let expected_call_args: any, expected_values: any, installed_fonts: any;
        [expected_call_args, expected_values] = $list(installed_fixture);
        installed_fonts = FontFiles._installedFonts();
        $assertEqual(FontFiles._iterFontFilesIn.callArgsList, expected_call_args);
        $assertEqual(installed_fonts, expected_values);
      } },
    { name: "it_generates_font_dirs_to_help_find", args: ["font_dirs_fixture"], fn: (self: any, font_dirs_fixture: any): any => {
        let expected_values: any, font_dirs: any;
        expected_values = font_dirs_fixture;
        font_dirs = FontFiles._fontDirectories();
        $assertEqual(font_dirs, expected_values);
      } },
    { name: "it_knows_os_x_font_dirs_to_help_find", args: ["osx_dirs_fixture"], fn: (self: any, osx_dirs_fixture: any): any => {
        let expected_dirs: any, font_dirs: any;
        expected_dirs = osx_dirs_fixture;
        font_dirs = FontFiles._osXFontDirectories();
        $assertEqual(font_dirs, expected_dirs);
      } },
    { name: "it_knows_windows_font_dirs_to_help_find", args: ["win_dirs_fixture"], fn: (self: any, win_dirs_fixture: any): any => {
        let expected_dirs: any, font_dirs: any;
        expected_dirs = win_dirs_fixture;
        font_dirs = FontFiles._windowsFontDirectories();
        $assertEqual(font_dirs, expected_dirs);
      } },
    { name: "it_iterates_over_fonts_in_dir_to_help_find", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let _Font_: any, directory: any, expected_calls: any, expected_paths: any, paths: any;
        [directory, _Font_, expected_calls, expected_paths] = $list(iter_fixture);
        paths = $list(FontFiles._iterFontFilesIn(directory));
        $assertEqual(_Font_.open.callArgsList, expected_calls);
        $assertEqual(paths, expected_paths);
      } },
  ],
};
pytestClass("DescribeFontFiles", DescribeFontFiles);

const Describe_Font: any = {
  fixtures: {
    bold_fixture: { args: ["request", "_tables_", "head_table_"], params: [["head", true, true], ["head", false, false], ["foob", true, false]], fn: (self: any, request: any, _tables_: any, head_table_: any): any => {
        let expected_value: any, font: any, is_bold: any, key: any;
        [key, is_bold, expected_value] = $list(request.param);
        head_table_.isBold = is_bold;
        _tables_.returnValue = $d({[key]: head_table_});
        font = $call(_Font, null);
        return [font, expected_value];
      } },
    family_fixture: { args: ["_tables_", "name_table_"], fn: (self: any, _tables_: any, name_table_: any): any => {
        let expected_name: any, font: any;
        font = $call(_Font, null);
        expected_name = "Foobar";
        _tables_.returnValue = $d({name: name_table_});
        name_table_.familyName = expected_name;
        return [font, expected_name];
      } },
    italic_fixture: { args: ["request", "_tables_", "head_table_"], params: [["head", true, true], ["head", false, false], ["foob", true, false]], fn: (self: any, request: any, _tables_: any, head_table_: any): any => {
        let expected_value: any, font: any, is_italic: any, key: any;
        [key, is_italic, expected_value] = $list(request.param);
        head_table_.isItalic = is_italic;
        _tables_.returnValue = $d({[key]: head_table_});
        font = $call(_Font, null);
        return [font, expected_value];
      } },
    iter_fixture: { args: ["_table_count_", "stream_read_"], fn: (self: any, _table_count_: any, stream_read_: any): any => {
        let expected_values: any, font: any, stream: any;
        stream = $call(_Stream, null);
        font = $call(_Font, stream);
        _table_count_.returnValue = 2;
        stream_read_.returnValue = $b("namexxxx\u0000\u0000\u0000*\u0000\u0000\u0000\u0015headxxxx\u0000\u0000\u0000\u0015\u0000\u0000\u0000*");
        expected_values = [["name", 42, 21], ["head", 21, 42]];
        return [font, expected_values];
      } },
    open_fixture: { args: ["_Stream_"], fn: (self: any, _Stream_: any): any => {
        let path: any, stream_: any;
        path = "foobar.ttf";
        stream_ = _Stream_.open.returnValue;
        return [path, _Stream_, stream_];
      } },
    tables_fixture: { args: ["stream_", "name_table_", "head_table_", "_iter_table_records_", "_TableFactory_"], fn: (self: any, stream_: any, name_table_: any, head_table_: any, _iter_table_records_: any, _TableFactory_: any): any => {
        let expected_calls: any, expected_tables: any, font: any;
        font = $call(_Font, stream_);
        _iter_table_records_.returnValue = $iter([["name", 11, 22], ["head", 33, 44]]);
        _TableFactory_.sideEffect = [name_table_, head_table_];
        expected_calls = [call("name", stream_, 11, 22), call("head", stream_, 33, 44)];
        expected_tables = $d({name: name_table_, head: head_table_});
        return [font, _TableFactory_, expected_calls, expected_tables];
      } },
    table_count_fixture: { args: ["_fields_"], fn: (self: any, _fields_: any): any => {
        let expected_value: any, font: any;
        font = $call(_Font, null);
        _fields_.returnValue = [-666, 42];
        expected_value = 42;
        return [font, expected_value];
      } },
    _fields_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Font, "_fields");
      } },
    head_table_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _HeadTable);
      } },
    _iter_table_records_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _Font, "_iter_table_records");
      } },
    name_table_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _NameTable);
      } },
    _Stream_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.fonts._Stream");
      } },
    stream_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Stream);
      } },
    stream_read_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, _Stream, "read");
      } },
    _TableFactory_: { args: ["request"], fn: (self: any, request: any): any => {
        return functionMock(request, "pptx.text.fonts._TableFactory");
      } },
    _table_count_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Font, "_table_count");
      } },
    _tables_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Font, "_tables");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_a_font_file_path", args: ["open_fixture"], fn: (self: any, open_fixture: any): any => {
        let _Stream_: any, f: any, path: any, stream_: any;
        [path, _Stream_, stream_] = $list(open_fixture);
        const $cm1 = _Font.open(path);
        f = $enter($cm1);
        try {
          _Stream_.open.assertCalledOnceWith(path);
          $assert($isinstance(f, _Font));
        } finally {
          $exit($cm1);
        }
        stream_.close.assertCalledOnceWith();
      } },
    { name: "it_knows_its_family_name", args: ["family_fixture"], fn: (self: any, family_fixture: any): any => {
        let expected_name: any, family_name: any, font: any;
        [font, expected_name] = $list(family_fixture);
        family_name = font.familyName;
        $assertEqual(family_name, expected_name);
      } },
    { name: "it_knows_whether_it_is_bold", args: ["bold_fixture"], fn: (self: any, bold_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(bold_fixture);
        $assert($is(font.isBold, expected_value));
      } },
    { name: "it_knows_whether_it_is_italic", args: ["italic_fixture"], fn: (self: any, italic_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(italic_fixture);
        $assert($is(font.isItalic, expected_value));
      } },
    { name: "it_provides_access_to_its_tables", args: ["tables_fixture"], fn: (self: any, tables_fixture: any): any => {
        let _TableFactory_: any, expected_calls: any, expected_tables: any, font: any, tables: any;
        [font, _TableFactory_, expected_calls, expected_tables] = $list(tables_fixture);
        tables = font._tables;
        $assertEqual(_TableFactory_.callArgsList, expected_calls);
        $assertEqual(tables, expected_tables);
      } },
    { name: "it_generates_table_records_to_help_read_tables", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let expected_values: any, font: any, values: any;
        [font, expected_values] = $list(iter_fixture);
        values = $list(font._iterTableRecords());
        $assertEqual(values, expected_values);
      } },
    { name: "it_knows_the_table_count_to_help_read", args: ["table_count_fixture"], fn: (self: any, table_count_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(table_count_fixture);
        $assertEqual(font._tableCount, expected_value);
      } },
    { name: "it_reads_the_header_to_help_read_font", args: ["request"], fn: (self: any, request: any): any => {
        let fields: any, font: any, stream_: any;
        stream_ = instanceMock(request, _Stream);
        stream_.readFields.returnValue = ["foob", 42, 64, 7, 16];
        font = $call(_Font, stream_);
        fields = font._fields;
        stream_.readFields.assertCalledOnceWith(">4sHHHH", 0);
        $assertEqual(fields, ["foob", 42, 64, 7, 16]);
      } },
  ],
};
pytestClass("Describe_Font", Describe_Font);

const Describe_Stream: any = {
  fixtures: {
    close_fixture: { args: ["file_"], fn: (self: any, file_: any): any => {
        let stream: any;
        stream = $call(_Stream, file_);
        return [stream, file_];
      } },
    read_fixture: { args: ["file_"], fn: (self: any, file_: any): any => {
        let expected_value: any, length: any, offset: any, stream: any;
        stream = $call(_Stream, file_);
        [offset, length] = $list([42, 21]);
        file_.read.returnValue = "foobar";
        expected_value = "foobar";
        return [stream, offset, length, file_, expected_value];
      } },
    read_flds_fixture: { args: ["file_"], fn: (self: any, file_: any): any => {
        let expected_values: any, offset: any, stream: any, tmpl: any;
        stream = $call(_Stream, file_);
        [tmpl, offset] = $list([$b(">4sHH"), 0]);
        file_.read.returnValue = $b("foob\u0000*\u0000\u0015");
        expected_values = [$b("foob"), 42, 21];
        return [stream, tmpl, offset, file_, expected_values];
      } },
    file_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BytesIO);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_a_path", skip: "patches python's builtin open(); _Stream.open() reads the file through pptx-ts's fileio", args: ["request"], fn: (self: any, request: any): any => {
        let _init_: any, file_: any, open_: any, stream: any;
        open_ = openMock(request, "pptx.text.fonts");
        _init_ = initializerMock(request, _Stream);
        file_ = open_.returnValue;
        stream = _Stream.open("foobar.ttf");
        open_.assertCalledOnceWith("foobar.ttf", "rb");
        _init_.assertCalledOnceWith(stream, file_);
        $assert($isinstance(stream, _Stream));
      } },
    { name: "it_can_be_closed", args: ["close_fixture"], fn: (self: any, close_fixture: any): any => {
        let file_: any, stream: any;
        [stream, file_] = $list(close_fixture);
        stream.close();
        file_.close.assertCalledOnceWith();
      } },
    { name: "it_can_read_fields_from_a_template", args: ["read_flds_fixture"], fn: (self: any, read_flds_fixture: any): any => {
        let expected_values: any, fields: any, file_: any, offset: any, stream: any, tmpl: any;
        [stream, tmpl, offset, file_, expected_values] = $list(read_flds_fixture);
        fields = stream.readFields(tmpl, offset);
        file_.seek.assertCalledOnceWith(offset);
        file_.read.assertCalledOnceWith(calcsize(tmpl));
        $assertEqual(fields, expected_values);
      } },
    { name: "it_can_read_bytes", args: ["read_fixture"], fn: (self: any, read_fixture: any): any => {
        let bytes_: any, expected_value: any, file_: any, length: any, offset: any, stream: any;
        [stream, offset, length, file_, expected_value] = $list(read_fixture);
        bytes_ = $read(stream, offset, length);
        file_.seek.assertCalledOnceWith(offset);
        file_.read.assertCalledOnceWith(length);
        $assertEqual(bytes_, expected_value);
      } },
  ],
};
pytestClass("Describe_Stream", Describe_Stream);

const Describe_TableFactory: any = {
  fixtures: {
    fixture: { args: ["request", "stream_"], params: ["name", "head", "foob"], fn: (self: any, request: any, stream_: any): any => {
        let TableClass: any, TableClass_: any, length: any, offset: any, tag: any, target: any;
        tag = request.param;
        [offset, length] = $list([42, 21]);
        [TableClass, target] = $list($get($d({name: [_NameTable, "pptx.text.fonts._NameTable"], head: [_HeadTable, "pptx.text.fonts._HeadTable"], foob: [_BaseTable, "pptx.text.fonts._BaseTable"]}), tag));
        TableClass_ = classMock(request, target);
        return [tag, stream_, offset, length, TableClass_, TableClass];
      } },
    stream_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Stream);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_appropriate_table_object", args: ["fixture"], fn: (self: any, fixture: any): any => {
        let TableClass: any, TableClass_: any, length: any, offset: any, stream_: any, table: any, tag: any;
        [tag, stream_, offset, length, TableClass_, TableClass] = $list(fixture);
        table = $call(_TableFactory, tag, stream_, offset, length);
        TableClass_.assertCalledOnceWith(tag, stream_, offset, length);
        $assert($isinstance(table, TableClass));
      } },
  ],
};
pytestClass("Describe_TableFactory", Describe_TableFactory);

const Describe_HeadTable: any = {
  fixtures: {
    bold_fixture: { args: ["request", "_macStyle_"], params: [[0, false], [1, true]], fn: (self: any, request: any, _macStyle_: any): any => {
        let expected_value: any, head_table: any, macStyle: any;
        [macStyle, expected_value] = $list(request.param);
        _macStyle_.returnValue = macStyle;
        head_table = $call(_HeadTable, null, null, null, null);
        return [head_table, expected_value];
      } },
    italic_fixture: { args: ["request", "_macStyle_"], params: [[5, false], [7, true]], fn: (self: any, request: any, _macStyle_: any): any => {
        let expected_value: any, head_table: any, macStyle: any;
        [macStyle, expected_value] = $list(request.param);
        _macStyle_.returnValue = macStyle;
        head_table = $call(_HeadTable, null, null, null, null);
        return [head_table, expected_value];
      } },
    macStyle_fixture: { args: [], fn: (self: any): any => {
        let bytes_: any, expected_value: any, head_table: any, length: any, offset: any, stream: any;
        bytes_ = $b("xxxxyyyy....................................ðº........");
        stream = $call(_Stream, new BytesIO(bytes_));
        [offset, length] = $list([0, $len(bytes_)]);
        head_table = $call(_HeadTable, null, stream, offset, length);
        expected_value = 61626;
        return [head_table, expected_value];
      } },
    _macStyle_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _HeadTable, "_macStyle");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_the_font_is_bold", args: ["bold_fixture"], fn: (self: any, bold_fixture: any): any => {
        let expected_value: any, head_table: any;
        [head_table, expected_value] = $list(bold_fixture);
        $assert($is(head_table.isBold, expected_value));
      } },
    { name: "it_knows_whether_the_font_is_italic", args: ["italic_fixture"], fn: (self: any, italic_fixture: any): any => {
        let expected_value: any, head_table: any;
        [head_table, expected_value] = $list(italic_fixture);
        $assert($is(head_table.isItalic, expected_value));
      } },
    { name: "it_reads_its_macStyle_field_to_help", args: ["macStyle_fixture"], fn: (self: any, macStyle_fixture: any): any => {
        let expected_value: any, head_table: any;
        [head_table, expected_value] = $list(macStyle_fixture);
        $assertEqual(head_table._macStyle, expected_value);
      } },
  ],
};
pytestClass("Describe_HeadTable", Describe_HeadTable);

const Describe_NameTable: any = {
  fixtures: {
    bytes_fixture: { args: ["stream_"], fn: (self: any, stream_: any): any => {
        let bytes_: any, expected_value: any, name_table: any;
        name_table = $call(_NameTable, null, stream_, 42, 360);
        bytes_ = "\u0000\u0001\u0002\u0003\u0004\u0005";
        stream_.read.returnValue = bytes_;
        expected_value = bytes_;
        return [name_table, expected_value];
      } },
    decode_fixture: { args: ["request"], params: [[1, 0, $b("Foob\u008ar"), "Foobär"], [1, 1, $b("Foobar"), null], [0, 9, $encode("Foobär", "utf-16-be"), "Foobär"], [3, 6, $encode("Foobär", "utf-16-be"), "Foobär"], [2, 0, "Foobar", null]], fn: (self: any, request: any): any => {
        let encoding_id: any, expected_value: any, name_table: any, platform_id: any, raw_name: any;
        [platform_id, encoding_id, raw_name, expected_value] = $list(request.param);
        name_table = $call(_NameTable, null, null, null, null);
        return [name_table, raw_name, platform_id, encoding_id, expected_value];
      } },
    family_fixture: { args: ["request", "_names_prop_"], params: [[$d({[[0, 1]]: "Foobar", [[1, 1]]: "Barfoo"}), "Foobar"], [$d({[[1, 1]]: "Barfoo", [[3, 1]]: "Farbaz"}), "Barfoo"], [$d({[[3, 1]]: "Farbaz", [[6, 2]]: "BazFoo"}), "Farbaz"], [$d({[[9, 1]]: "Foobar", [[6, 1]]: "Barfoo"}), null]], fn: (self: any, request: any, _names_prop_: any): any => {
        let expected_value: any, name_table: any, names: any;
        [names, expected_value] = $list(request.param);
        name_table = $call(_NameTable, null, null, null, null);
        _names_prop_.returnValue = names;
        return [name_table, expected_value];
      } },
    header_fixture: { args: ["_table_bytes_prop_"], fn: (self: any, _table_bytes_prop_: any): any => {
        let expected_value: any, name_table: any;
        name_table = $call(_NameTable, null, null, null, null);
        _table_bytes_prop_.returnValue = $b("\u0000\u0000\u0000\u0002\u0000*");
        expected_value = [0, 2, 42];
        return [name_table, expected_value];
      } },
    name_hdr_fixture: { args: [], fn: (self: any): any => {
        let bufr: any, expected_value: any, idx: any, name_table: any;
        name_table = $call(_NameTable, null, null, null, null);
        bufr = $b("123456123456789012\u0000\u0000\u0000\u0001\u0000\u0002\u0000\u0003\u0000\u0004\u0000\u0005");
        idx = 1;
        expected_value = [0, 1, 2, 3, 4, 5];
        return [name_table, bufr, idx, expected_value];
      } },
    name_text_fixture: { args: ["_raw_name_string_", "_decode_name_"], fn: (self: any, _raw_name_string_: any, _decode_name_: any): any => {
        let bufr: any, encoding_id: any, length: any, name_: any, name_str_offset: any, name_table: any, platform_id: any, raw_name: any, strings_offset: any;
        name_table = $call(_NameTable, null, null, null, null);
        [bufr, platform_id, encoding_id, strings_offset] = $list(["xXx", 6, 7, 8]);
        [name_str_offset, length, raw_name] = $list([9, 10, "Foobar"]);
        _raw_name_string_.returnValue = raw_name;
        name_ = _decode_name_.returnValue;
        return [name_table, bufr, platform_id, encoding_id, strings_offset, name_str_offset, length, raw_name, name_];
      } },
    raw_fixture: { args: [], fn: (self: any): any => {
        let bufr: any, expected_bytes: any, length: any, name_table: any, str_offset: any, strings_offset: any;
        name_table = $call(_NameTable, null, null, null, null);
        bufr = $b("xXxFoobarxXx");
        [strings_offset, str_offset, length] = $list([1, 2, 6]);
        expected_bytes = $b("Foobar");
        return [name_table, bufr, strings_offset, str_offset, length, expected_bytes];
      } },
    _decode_name_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _NameTable, "_decode_name", $kw({autospec: false}));
      } },
    _names_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _NameTable, "_names");
      } },
    _raw_name_string_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _NameTable, "_raw_name_string", $kw({autospec: false}));
      } },
    stream_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Stream);
      } },
    _table_bytes_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _NameTable, "_table_bytes");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_font_family_name", args: ["family_fixture"], fn: (self: any, family_fixture: any): any => {
        let expected_value: any, family_name: any, name_table: any;
        [name_table, expected_value] = $list(family_fixture);
        family_name = name_table.familyName;
        $assertEqual(family_name, expected_value);
      } },
    { name: "it_provides_access_to_its_names_to_help_props", args: ["request"], fn: (self: any, request: any): any => {
        let _iter_names_: any, name_table: any, names: any;
        _iter_names_ = $call(methodMock, request, _NameTable, "_iter_names", $kw({return_value: $iter([[[0, 1], "Foobar"], [[3, 1], "Barfoo"]])}));
        name_table = $call(_NameTable, null, null, null, null);
        names = name_table._names;
        _iter_names_.assertCalledOnceWith(name_table);
        $assertEqual(names, $d({[[0, 1]]: "Foobar", [[3, 1]]: "Barfoo"}));
      } },
    { name: "it_iterates_over_its_names_to_help_read_names", args: ["request", "_table_bytes_prop_"], fn: (self: any, request: any, _table_bytes_prop_: any): any => {
        let _read_name_: any, name_table: any, names: any;
        $call(propertyMock, request, _NameTable, "_table_header", $kw({return_value: [0, 3, 42]}));
        _table_bytes_prop_.returnValue = "xXx";
        _read_name_ = $call(methodMock, request, _NameTable, "_read_name", $kw({side_effect: $iter([[0, 1, "Foobar"], [3, 1, "Barfoo"], [9, 9, null]])}));
        name_table = $call(_NameTable, null, null, null, null);
        names = $list(name_table._iterNames());
        $assertEqual(_read_name_.callArgsList, [call(name_table, "xXx", 0, 42), call(name_table, "xXx", 1, 42), call(name_table, "xXx", 2, 42)]);
        $assertEqual(names, [[[0, 1], "Foobar"], [[3, 1], "Barfoo"]]);
      } },
    { name: "it_reads_the_table_header_to_help_read_names", args: ["header_fixture"], fn: (self: any, header_fixture: any): any => {
        let expected_value: any, header: any, names_table: any;
        [names_table, expected_value] = $list(header_fixture);
        header = names_table._tableHeader;
        $assertEqual(header, expected_value);
      } },
    { name: "it_buffers_the_table_bytes_to_help_read_names", args: ["bytes_fixture"], fn: (self: any, bytes_fixture: any): any => {
        let expected_value: any, name_table: any, table_bytes: any;
        [name_table, expected_value] = $list(bytes_fixture);
        table_bytes = name_table._tableBytes;
        name_table._stream.read.assertCalledOnceWith(name_table._offset, name_table._length);
        $assertEqual(table_bytes, expected_value);
      } },
    { name: "it_reads_a_name_to_help_read_names", args: ["request"], fn: (self: any, request: any): any => {
        let _name_header_: any, _read_name_text_: any, actual: any, bufr: any, encoding_id: any, idx: any, length: any, name: any, name_id: any, name_str_offset: any, name_table: any, platform_id: any, strs_offset: any;
        [bufr, idx, strs_offset, platform_id, name_id] = $list(["buffer", 3, 47, 0, 1]);
        [encoding_id, name_str_offset, length, name] = $list([7, 36, 12, "Arial"]);
        _name_header_ = $call(methodMock, request, _NameTable, "_name_header", $kw({autospec: false, return_value: [platform_id, encoding_id, 666, name_id, length, name_str_offset]}));
        _read_name_text_ = $call(methodMock, request, _NameTable, "_read_name_text", $kw({return_value: name}));
        name_table = $call(_NameTable, null, null, null, null);
        actual = name_table._readName(bufr, idx, strs_offset);
        _name_header_.assertCalledOnceWith(bufr, idx);
        _read_name_text_.assertCalledOnceWith(name_table, bufr, platform_id, encoding_id, strs_offset, name_str_offset, length);
        $assertEqual(actual, [platform_id, name_id, name]);
      } },
    { name: "it_reads_a_name_header_to_help_read_names", args: ["name_hdr_fixture"], fn: (self: any, name_hdr_fixture: any): any => {
        let bufr: any, expected_value: any, header: any, idx: any, name_table: any;
        [name_table, bufr, idx, expected_value] = $list(name_hdr_fixture);
        header = name_table._nameHeader(bufr, idx);
        $assertEqual(header, expected_value);
      } },
    { name: "it_reads_name_text_to_help_read_names", args: ["name_text_fixture"], fn: (self: any, name_text_fixture: any): any => {
        let bufr: any, encoding_id: any, length: any, name: any, name_: any, name_str_offset: any, name_table: any, platform_id: any, raw_name: any, strings_offset: any;
        [name_table, bufr, platform_id, encoding_id] = $list($slice(name_text_fixture, null, 4, null));
        [strings_offset, name_str_offset, length] = $list($slice(name_text_fixture, 4, 7, null));
        [raw_name, name_] = $list($slice(name_text_fixture, 7, null, null));
        name = name_table._readNameText(bufr, platform_id, encoding_id, strings_offset, name_str_offset, length);
        name_table._rawNameString.assertCalledOnceWith(bufr, strings_offset, name_str_offset, length);
        name_table._decodeName.assertCalledOnceWith(raw_name, platform_id, encoding_id);
        $assert($is(name, name_));
      } },
    { name: "it_reads_name_bytes_to_help_read_names", args: ["raw_fixture"], fn: (self: any, raw_fixture: any): any => {
        let bufr: any, bytes_: any, expected_bytes: any, length: any, name_table: any, str_offset: any, strings_offset: any;
        [name_table, bufr, strings_offset, str_offset] = $list($slice(raw_fixture, null, 4, null));
        [length, expected_bytes] = $list($slice(raw_fixture, 4, null, null));
        bytes_ = name_table._rawNameString(bufr, strings_offset, str_offset, length);
        $assertEqual(bytes_, expected_bytes);
      } },
    { name: "it_decodes_a_raw_name_to_help_read_names", args: ["decode_fixture"], fn: (self: any, decode_fixture: any): any => {
        let encoding_id: any, expected_value: any, name: any, name_table: any, platform_id: any, raw_name: any;
        [name_table, raw_name, platform_id, encoding_id, expected_value] = $list(decode_fixture);
        name = name_table._decodeName(raw_name, platform_id, encoding_id);
        $assertEqual(name, expected_value);
      } },
  ],
};
pytestClass("Describe_NameTable", Describe_NameTable);

