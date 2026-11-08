// Machine-translated from python-pptx tests/oxml/test_slide.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assertEqual, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CT_NotesMaster as CT_NotesMaster$imported, CT_NotesSlide as CT_NotesSlide$imported } from "../../src/oxml/slide.js";
import { snippetText } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const CT_NotesMaster = CT_NotesMaster$imported;
const CT_NotesSlide = CT_NotesSlide$imported;
void $py;


const DescribeCT_NotesMaster: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_default_notesMaster_element", args: [], fn: (self: any): any => {
        let notesMaster: any;
        notesMaster = CT_NotesMaster.newDefault();
        $assertEqual(notesMaster.xml, snippetText("default-notesMaster"));
      } },
  ],
};
pytestClass("DescribeCT_NotesMaster", DescribeCT_NotesMaster);

const DescribeCT_NotesSlide: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_new_notes_element", args: [], fn: (self: any): any => {
        let notes: any;
        notes = CT_NotesSlide.new();
        $assertEqual(notes.xml, snippetText("default-notes"));
      } },
  ],
};
pytestClass("DescribeCT_NotesSlide", DescribeCT_NotesSlide);

