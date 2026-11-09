// Machine-translated from python-pptx features/steps/coreprops.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $cmp, $getattr, $kw, $list, $mod, $setattr } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { no_core_props_pptx_path, saved_pptx_path } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
const datetime = $py.$datetime.datetime;
const timedelta = $py.$datetime.timedelta;
void $py;


function step_given_ref_to_core_doc_props(context: any): any {
  context.prs = $call(Presentation);
  context.coreProperties = context.prs.coreProperties;
}
step("given", "I have a reference to the core properties of a presentation", [], step_given_ref_to_core_doc_props);

function step_when_open_presentation_with_no_core_props_part(context: any): any {
  context.prs = $call(Presentation, no_core_props_pptx_path);
}
step("when", "I open a presentation having no core properties part", [], step_when_open_presentation_with_no_core_props_part);

function step_when_set_core_doc_props_to_valid_values(context: any): any {
  context.propvals = [["author", "Creator"], ["category", "Category"], ["comments", "Description"], ["content_status", "Content Status"], ["created", datetime(2013, 6, 15, 12, 34, 56)], ["identifier", "Identifier"], ["keywords", "key; word; keyword"], ["language", "Language"], ["last_modified_by", "Last Modified By"], ["last_printed", datetime(2013, 6, 15, 12, 34, 56)], ["modified", datetime(2013, 6, 15, 12, 34, 56)], ["revision", 9], ["subject", "Subject"], ["title", "åß∂Title°"], ["version", "Version"]];
  for (const [name, value] of $list(context.propvals)) {
    $setattr(context.prs.coreProperties, name, value);
  }
}
step("when", "I set the core properties to valid values", [], step_when_set_core_doc_props_to_valid_values);

function step_then_a_core_props_part_with_def_vals_is_added(context: any): any {
  let core_props: any, max_expected_timedelta: any, modified_timedelta: any;
  core_props = context.prs.coreProperties;
  $assertEqual(core_props.title, "PowerPoint Presentation");
  $assertEqual(core_props.lastModifiedBy, "python-pptx");
  $assertEqual(core_props.revision, 1);
  modified_timedelta = (datetime.utcnow() - core_props.modified);
  max_expected_timedelta = $call(timedelta, $kw({seconds: 2}));
  $assert($cmp(modified_timedelta, "<", max_expected_timedelta));
}
step("then", "a core properties part with default values is added", [], step_then_a_core_props_part_with_def_vals_is_added);

function step_then_core_props_have_values_previously_set(context: any): any {
  let core_props: any;
  core_props = $call(Presentation, saved_pptx_path).coreProperties;
  for (const [name, value] of $list(context.propvals)) {
    $assertEqual($getattr(core_props, name), value, $mod("for core property '%s'", name));
  }
}
step("then", "the core properties of the presentation have the values I set", [], step_then_core_props_have_values_previously_set);

// the names python's eval() sees in this module
const $scope = { Presentation };
