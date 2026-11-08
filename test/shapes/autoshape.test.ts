// Machine-translated from python-pptx tests/shapes/test_autoshape.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $cast, $dict, $eq, $get, $is, $isinstance, $kw, $len, $list, $raises, $range, $set, $str, classMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { FillFormat as FillFormat$imported } from "../../src/dml/fill.js";
import { LineFormat as LineFormat$imported } from "../../src/dml/line.js";
import { MSO_SHAPE as MSO_SHAPE$imported, MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported } from "../../src/enum/shapes.js";
import { parseXml as parseXml$imported } from "../../src/oxml/index.js";
import { CT_PresetGeometry2D as CT_PresetGeometry2D$imported, CT_Shape as CT_Shape$imported } from "../../src/oxml/shapes/autoshape.js";
import { Adjustment as Adjustment$imported, AdjustmentCollection as AdjustmentCollection$imported, AutoShapeType as AutoShapeType$imported, Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { TextFrame as TextFrame$imported } from "../../src/text/text.js";
import { aCNvSpPr, aGd, aPrstGeom, anAvLst, anNvSpPr, anSp, anSpPr } from "../oxml/unitdata/shape.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const FillFormat = FillFormat$imported;
const LineFormat = LineFormat$imported;
const MSO_SHAPE = MSO_SHAPE$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const parseXml = parseXml$imported;
const CT_PresetGeometry2D = CT_PresetGeometry2D$imported;
const CT_Shape = CT_Shape$imported;
const Adjustment = Adjustment$imported;
const AdjustmentCollection = AdjustmentCollection$imported;
const AutoShapeType = AutoShapeType$imported;
const Shape = Shape$imported;
const TextFrame = TextFrame$imported;
void $py;



const DescribeAdjustment: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_effective_value", args: ["def_val", "actual", "expected_value"], parametrize: [{ names: ["def_val", "actual", "expected_value"], values: [[50000, null, 0.5], [50000, 50000, 0.5], [50000, 12500, 0.125], [50000, 0, 0.0], [-20833, null, -0.20833], [-20833, -5678901, -56.78901]] }], fn: (self: any, def_val: any, actual: any, expected_value: any): any => {
        $assertEqual($call(Adjustment, "foobar", def_val, actual).effectiveValue, expected_value);
      } },
  ],
};
pytestClass("DescribeAdjustment", DescribeAdjustment);

const _adj_actuals_cases = (): any => {
  let avLst_bldr: any, gd_bldr: any, gd_bldr_1: any, gd_bldr_2: any, gd_bldr_3: any, mathDivide_bldr: any, rect_bldr: any, wedgeRoundRectCallout_bldr: any;
  gd_bldr = aGd().withName("adj2").withFmla("val 25000");
  avLst_bldr = anAvLst().withChild(gd_bldr);
  mathDivide_bldr = aPrstGeom().withNsdecls().withPrst("mathDivide").withChild(avLst_bldr);
  gd_bldr = aGd().withName("adj").withFmla("val 25000");
  avLst_bldr = anAvLst().withChild(gd_bldr);
  rect_bldr = aPrstGeom().withNsdecls().withPrst("rect").withChild(avLst_bldr);
  gd_bldr_1 = aGd().withName("adj1").withFmla("val 111");
  gd_bldr_2 = aGd().withName("adj2").withFmla("val 222");
  gd_bldr_3 = aGd().withName("adj3").withFmla("val 333");
  avLst_bldr = anAvLst().withChild(gd_bldr_1).withChild(gd_bldr_2).withChild(gd_bldr_3);
  wedgeRoundRectCallout_bldr = aPrstGeom().withNsdecls().withPrst("wedgeRoundRectCallout").withChild(avLst_bldr);
  return [[aPrstGeom().withNsdecls().withPrst("rect").withChild(anAvLst()), []], [aPrstGeom().withNsdecls().withPrst("circularArrow"), [["adj1", null], ["adj2", null], ["adj3", null], ["adj4", null], ["adj5", null]]], [rect_bldr, []], [mathDivide_bldr, [["adj1", null], ["adj2", 25000], ["adj3", null]]], [wedgeRoundRectCallout_bldr, [["adj1", 111], ["adj2", 222], ["adj3", 333]]]];
};
const _effective_val_cases = (): any => {
  return [["rect", []], ["chevron", [0.5]], ["circularArrow", [0.125, 11.42319, 204.57681, 108.0, 0.125]]];
};
const _indexed_assignment_cases = (): any => {
  return [["chevron", 0, 0.5, 50000], ["circularArrow", 4, 99.99, 9999000]];
};
const DescribeAdjustmentCollection: any = {
  fixtures: {
    adjustments: { args: [], fn: (self: any): any => {
        let prstGeom: any;
        prstGeom = aPrstGeom().withNsdecls().withPrst("chevron").element;
        return $call(AdjustmentCollection, prstGeom);
      } },
    adjustments_with_prstGeom_: { args: ["request"], fn: (self: any, request: any): any => {
        let adjustments: any, guides: any, prstGeom: any, prstGeom_: any;
        prstGeom = aPrstGeom().withNsdecls().withPrst("chevron").element;
        adjustments = $call(AdjustmentCollection, prstGeom);
        prstGeom_ = $call(instanceMock, request, CT_PresetGeometry2D, $kw({name: "prstGeom_"}));
        adjustments._prstGeom = prstGeom_;
        guides = [["adj", 99900]];
        return [adjustments, guides];
      } },
    load_adj_actuals_fixture_: { args: ["request"], params: _adj_actuals_cases(), fn: (self: any, request: any): any => {
        let adj_vals: any, expected: any, prstGeom: any, prstGeom_bldr: any, prstGeom_xml: any;
        [prstGeom_bldr, adj_vals] = $list(request.param);
        prstGeom = prstGeom_bldr.element;
        expected = $dict(adj_vals);
        prstGeom_xml = prstGeom_bldr.xml;
        return [prstGeom, expected, prstGeom_xml];
      } },
    indexed_access_fixture_: { args: ["request"], params: _effective_val_cases(), fn: (self: any, request: any): any => {
        let effective_values: any, prst: any, prstGeom: any;
        [prst, effective_values] = $list(request.param);
        prstGeom = aPrstGeom().withNsdecls().withPrst(prst).element;
        return [prstGeom, prst, $list(effective_values)];
      } },
    indexed_assignment_fixture_: { args: ["request"], params: _indexed_assignment_cases(), fn: (self: any, request: any): any => {
        let adjustments: any, expected: any, idx: any, new_val: any, prst: any, prstGeom: any;
        [prst, idx, new_val, expected] = $list(request.param);
        prstGeom = aPrstGeom().withNsdecls().withPrst(prst).element;
        adjustments = $call(AdjustmentCollection, prstGeom);
        return [adjustments, idx, new_val, expected];
      } },
  },
  helpers: {
    _adjActualsCases: (self: any): any => {
        let avLst_bldr: any, gd_bldr: any, gd_bldr_1: any, gd_bldr_2: any, gd_bldr_3: any, mathDivide_bldr: any, rect_bldr: any, wedgeRoundRectCallout_bldr: any;
        gd_bldr = aGd().withName("adj2").withFmla("val 25000");
        avLst_bldr = anAvLst().withChild(gd_bldr);
        mathDivide_bldr = aPrstGeom().withNsdecls().withPrst("mathDivide").withChild(avLst_bldr);
        gd_bldr = aGd().withName("adj").withFmla("val 25000");
        avLst_bldr = anAvLst().withChild(gd_bldr);
        rect_bldr = aPrstGeom().withNsdecls().withPrst("rect").withChild(avLst_bldr);
        gd_bldr_1 = aGd().withName("adj1").withFmla("val 111");
        gd_bldr_2 = aGd().withName("adj2").withFmla("val 222");
        gd_bldr_3 = aGd().withName("adj3").withFmla("val 333");
        avLst_bldr = anAvLst().withChild(gd_bldr_1).withChild(gd_bldr_2).withChild(gd_bldr_3);
        wedgeRoundRectCallout_bldr = aPrstGeom().withNsdecls().withPrst("wedgeRoundRectCallout").withChild(avLst_bldr);
        return [[aPrstGeom().withNsdecls().withPrst("rect").withChild(anAvLst()), []], [aPrstGeom().withNsdecls().withPrst("circularArrow"), [["adj1", null], ["adj2", null], ["adj3", null], ["adj4", null], ["adj5", null]]], [rect_bldr, []], [mathDivide_bldr, [["adj1", null], ["adj2", 25000], ["adj3", null]]], [wedgeRoundRectCallout_bldr, [["adj1", 111], ["adj2", 222], ["adj3", 333]]]];
      },
    _effectiveValCases: (self: any): any => {
        return [["rect", []], ["chevron", [0.5]], ["circularArrow", [0.125, 11.42319, 204.57681, 108.0, 0.125]]];
      },
    _indexedAssignmentCases: (self: any): any => {
        return [["chevron", 0, 0.5, 50000], ["circularArrow", 4, 99.99, 9999000]];
      },
  },
  tests: [
    { name: "it_should_load_default_adjustment_values", args: ["prst", "expected_values"], parametrize: [{ names: ["prst", "expected_values"], values: [["rect", []], ["chevron", [["adj", 50000]]], ["accentBorderCallout1", [["adj1", 18750], ["adj2", -8333], ["adj3", 112500], ["adj4", -38333]]], ["wedgeRoundRectCallout", [["adj1", -20833], ["adj2", 62500], ["adj3", 16667]]], ["circularArrow", [["adj1", 12500], ["adj2", 1142319], ["adj3", 20457681], ["adj4", 10800000], ["adj5", 12500]]]] }], fn: (self: any, prst: any, expected_values: any): any => {
        let actuals: any, adjustments: any, prstGeom: any;
        prstGeom = $cast(CT_PresetGeometry2D, element(`a:prstGeom{prst=${$str(prst)}}/a:avLst`));
        adjustments = $call(AdjustmentCollection, prstGeom)._adjustments;
        actuals = $list($list(adjustments).map((adj: any) => [adj.name, adj.defVal]));
        $assertEqual(actuals, expected_values);
      } },
    { name: "it_should_load_adj_val_actuals_from_xml", args: ["load_adj_actuals_fixture_"], fn: (self: any, load_adj_actuals_fixture_: any): any => {
        let actual_actuals: any, adjustments: any, expected_actuals: any, prstGeom: any, prstGeom_xml: any;
        [prstGeom, expected_actuals, prstGeom_xml] = $list(load_adj_actuals_fixture_);
        adjustments = $call(AdjustmentCollection, prstGeom)._adjustments;
        actual_actuals = $dict($list(adjustments).map((a: any) => [a.name, a.actual]));
        $assertEqual(actual_actuals, expected_actuals);
      } },
    { name: "it_provides_normalized_effective_value_on_indexed_access", args: ["indexed_access_fixture_"], fn: (self: any, indexed_access_fixture_: any): any => {
        let actual_values: any, adjustments: any, expected_values: any, prst: any, prstGeom: any;
        [prstGeom, prst, expected_values] = $list(indexed_access_fixture_);
        adjustments = $call(AdjustmentCollection, prstGeom);
        actual_values = $list($range($len(adjustments))).map((idx: any) => $get(adjustments, idx));
        $assertEqual(actual_values, expected_values);
      } },
    { name: "it_should_update_actual_value_on_indexed_assignment", args: ["indexed_assignment_fixture_"], fn: (self: any, indexed_assignment_fixture_: any): any => {
        let adjs: any, expected: any, idx: any, new_val: any;
        [adjs, idx, new_val, expected] = $list(indexed_assignment_fixture_);
        $set(adjs, idx, new_val);
        $assertEqual($get(adjs._adjustments, idx).actual, expected);
      } },
    { name: "it_should_round_trip_indexed_assignment", args: ["adjustments"], fn: (self: any, adjustments: any): any => {
        let new_value: any;
        new_value = 0.375;
        $assert(!$eq($get(adjustments, 0), new_value));
        $set(adjustments, 0, new_value);
        $assertEqual($get(adjustments, 0), new_value);
      } },
    { name: "it_should_raise_on_bad_index", args: ["adjustments"], fn: (self: any, adjustments: any): any => {
        $raises(IndexError, () => {
          $get(adjustments, -6);
        });
        $raises(IndexError, () => {
          $get(adjustments, 6);
        });
        $raises(TypeError, () => {
          $get(adjustments, "0");
        });
        $raises(IndexError, () => {
          $set(adjustments, -6, 1.0);
        });
        $raises(IndexError, () => {
          $set(adjustments, 6, 1.0);
        });
        $raises(TypeError, () => {
          $set(adjustments, "0", 1.0);
        });
      } },
    { name: "it_should_raise_on_assigned_bad_value", args: ["adjustments"], fn: (self: any, adjustments: any): any => {
        $raises(ValueError, () => {
          $set(adjustments, 0, "1.0");
        });
      } },
    { name: "it_writes_adj_vals_to_xml_on_assignment", args: ["adjustments_with_prstGeom_"], fn: (self: any, adjustments_with_prstGeom_: any): any => {
        let adjs: any, guides: any;
        [adjs, guides] = $list(adjustments_with_prstGeom_);
        $set(adjs, 0, 0.999);
        adjs._prstGeom.rewriteGuides.assertCalledOnceWith(guides);
      } },
  ],
};
pytestClass("DescribeAdjustmentCollection", DescribeAdjustmentCollection);

const DescribeAutoShapeType: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_details_of_the_auto_shape_type_it_represents", args: [], fn: (self: any): any => {
        let autoshape_type: any;
        autoshape_type = $call(AutoShapeType, MSO_SHAPE.ROUNDED_RECTANGLE);
        $assertEqual(autoshape_type.autoshapeTypeId, MSO_SHAPE.ROUNDED_RECTANGLE);
        $assertEqual(autoshape_type.prst, "roundRect");
        $assertEqual(autoshape_type.basename, "Rounded Rectangle");
      } },
    { name: "it_xml_escapes_the_basename_when_the_name_contains_special_characters", args: [], fn: (self: any): any => {
        let autoshape_type: any;
        autoshape_type = $call(AutoShapeType, MSO_SHAPE.NO_SYMBOL);
        $assertEqual(autoshape_type.autoshapeTypeId, MSO_SHAPE.NO_SYMBOL);
        $assertEqual(autoshape_type.prst, "noSmoking");
        $assertEqual(autoshape_type.basename, "&quot;No&quot; Symbol");
      } },
    { name: "it_knows_the_default_adj_vals_for_its_autoshape_type", args: ["prst", "default_adj_vals"], parametrize: [{ names: ["prst", "default_adj_vals"], values: [[MSO_SHAPE.RECTANGLE, []], [MSO_SHAPE.CHEVRON, [["adj", 50000]]], [MSO_SHAPE.LEFT_CIRCULAR_ARROW, [["adj1", 12500], ["adj2", -1142319], ["adj3", 1142319], ["adj4", 10800000], ["adj5", 12500]]]] }], fn: (self: any, prst: any, default_adj_vals: any): any => {
        let _default_adj_vals: any;
        _default_adj_vals = AutoShapeType.defaultAdjustmentValues(prst);
        $assertEqual(_default_adj_vals, default_adj_vals);
      } },
    { name: "it_knows_the_autoshape_type_id_for_each_prst_key", args: [], fn: (self: any): any => {
        $assertEqual(AutoShapeType.idFromPrst("rect"), MSO_SHAPE.RECTANGLE);
      } },
    { name: "it_raises_when_asked_for_autoshape_type_id_with_a_bad_prst", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          AutoShapeType.idFromPrst("badPrst");
        }, "MSO_AUTO_SHAPE_TYPE has no XML mapping for 'badPr");
      } },
    { name: "it_caches_autoshape_type_lookups", args: [], fn: (self: any): any => {
        let autoshape_type_1: any, autoshape_type_2: any, autoshape_type_id: any;
        autoshape_type_id = MSO_SHAPE.ROUNDED_RECTANGLE;
        autoshape_type_1 = $call(AutoShapeType, autoshape_type_id);
        autoshape_type_2 = $call(AutoShapeType, autoshape_type_id);
        $assert($is(autoshape_type_2, autoshape_type_1));
      } },
    { name: "it_raises_on_construction_with_bad_autoshape_type_id", args: [], fn: (self: any): any => {
        $raises(KeyError, () => {
          $call(AutoShapeType, 9999);
        });
      } },
  ],
};
pytestClass("DescribeAutoShapeType", DescribeAutoShapeType);

const DescribeShape: any = {
  fixtures: {
    autoshape_type_fixture_: { args: ["shape", "prst"], fn: (self: any, shape: any, prst: any): any => {
        let sp: any;
        sp = anSp().withNsdecls().withChild(anNvSpPr().withChild(aCNvSpPr())).withChild(anSpPr().withChild(aPrstGeom().withPrst("chevron"))).element;
        shape = $call(Shape, sp, null);
        return [shape, MSO_SHAPE.CHEVRON];
      } },
    init_adjs_fixture_: { args: ["request", "shape", "sp_", "adjustments_", "AdjustmentCollection_"], fn: (self: any, request: any, shape: any, sp_: any, adjustments_: any, AdjustmentCollection_: any): any => {
        return [shape, adjustments_, AdjustmentCollection_, sp_];
      } },
    text_frame_fixture: { args: ["TextFrame_", "text_frame_"], fn: (self: any, TextFrame_: any, text_frame_: any): any => {
        let shape: any, sp: any, txBody: any;
        sp = element("p:sp/p:txBody");
        txBody = $get(sp, 0);
        shape = $call(Shape, sp, null);
        return [shape, text_frame_, TextFrame_, txBody];
      } },
    txBody_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, shape: any, sp_cxml: any;
        sp_cxml = "p:sp";
        expected_cxml = "p:txBody/(a:bodyPr,a:p)";
        shape = $call(Shape, element(sp_cxml), null);
        expected_xml = xml(expected_cxml);
        return [shape, expected_xml];
      } },
    AdjustmentCollection_: { args: ["request", "adjustments_"], fn: (self: any, request: any, adjustments_: any): any => {
        return $call(classMock, request, "pptx.shapes.autoshape.AdjustmentCollection", $kw({return_value: adjustments_}));
      } },
    adjustments_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, AdjustmentCollection);
      } },
    non_autoshape_shape_: { args: ["request", "sp_"], fn: (self: any, request: any, sp_: any): any => {
        sp_.isAutoshape = false;
        return $call(Shape, sp_, null);
      } },
    non_placeholder_shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        let autoshape_shape_: any, autoshape_sp_: any, freeform_: any, freeform_sp_: any, textbox_shape_: any, textbox_sp_: any;
        autoshape_sp_ = $call(instanceMock, request, CT_Shape, $kw({name: "autoshape_sp_", is_autoshape: true, is_textbox: false, has_custom_geometry: false}));
        autoshape_shape_ = $call(Shape, autoshape_sp_, null);
        textbox_sp_ = $call(instanceMock, request, CT_Shape, $kw({name: "textbox_sp_", is_autoshape: false, is_textbox: true, has_custom_geometry: false}));
        textbox_shape_ = $call(Shape, textbox_sp_, null);
        freeform_sp_ = $call(instanceMock, request, CT_Shape, $kw({name: "freeform_sp_", is_autoshape: true, is_textbox: false, has_custom_geometry: true}));
        freeform_ = $call(Shape, freeform_sp_, null);
        $call(propertyMock, request, Shape, "is_placeholder", $kw({return_value: false}));
        return [autoshape_shape_, textbox_shape_, freeform_];
      } },
    placeholder_shape_: { args: ["request", "sp_"], fn: (self: any, request: any, sp_: any): any => {
        let placeholder_shape_: any;
        placeholder_shape_ = $call(Shape, sp_, null);
        $call(propertyMock, request, Shape, "is_placeholder", $kw({return_value: true}));
        return placeholder_shape_;
      } },
    prst: { args: [], fn: (self: any): any => {
        return MSO_SHAPE.CHEVRON;
      } },
    shape: { args: ["request", "sp_"], fn: (self: any, request: any, sp_: any): any => {
        return $call(Shape, sp_, null);
      } },
    sp_: { args: ["request", "prst"], fn: (self: any, request: any, prst: any): any => {
        return $call(instanceMock, request, CT_Shape, $kw({prst: prst, is_autoshape: true}));
      } },
    TextFrame_: { args: ["request", "text_frame_"], fn: (self: any, request: any, text_frame_: any): any => {
        return $call(classMock, request, "pptx.shapes.autoshape.TextFrame", $kw({return_value: text_frame_}));
      } },
    text_frame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, TextFrame);
      } },
    text_frame_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Shape, "text_frame");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_initializes_adjustments_on_first_ref", args: ["init_adjs_fixture_"], fn: (self: any, init_adjs_fixture_: any): any => {
        let AdjustmentCollection_: any, adjs_: any, shape: any, sp_: any;
        [shape, adjs_, AdjustmentCollection_, sp_] = $list(init_adjs_fixture_);
        $assert($is(shape.adjustments, adjs_));
        AdjustmentCollection_.assertCalledOnceWith(sp_.prstGeom);
      } },
    { name: "it_knows_its_autoshape_type", args: ["autoshape_type_fixture_"], fn: (self: any, autoshape_type_fixture_: any): any => {
        let auto_shape_type: any, expected_value: any, shape: any;
        [shape, expected_value] = $list(autoshape_type_fixture_);
        auto_shape_type = shape.autoShapeType;
        $assertEqual(auto_shape_type, expected_value);
      } },
    { name: "but_it_raises_when_auto_shape_type_called_on_non_autoshape", args: ["non_autoshape_shape_"], fn: (self: any, non_autoshape_shape_: any): any => {
        $raises(ValueError, () => {
          non_autoshape_shape_.autoShapeType;
        });
      } },
    { name: "it_has_a_fill", args: ["shape"], fn: (self: any, shape: any): any => {
        $assert($isinstance(shape.fill, FillFormat));
      } },
    { name: "it_knows_it_has_a_text_frame", args: [], fn: (self: any): any => {
        let shape: any;
        shape = $call(Shape, null, null);
        $assert($is(shape.hasTextFrame, true));
      } },
    { name: "it_has_a_line", args: ["shape"], fn: (self: any, shape: any): any => {
        $assert($isinstance(shape.line, LineFormat));
      } },
    { name: "it_knows_its_shape_type_when_its_a_placeholder", args: ["placeholder_shape_"], fn: (self: any, placeholder_shape_: any): any => {
        $assertEqual(placeholder_shape_.shapeType, MSO_SHAPE_TYPE.PLACEHOLDER);
      } },
    { name: "and_it_knows_its_shape_type_when_its_not_a_placeholder", args: ["non_placeholder_shapes_"], fn: (self: any, non_placeholder_shapes_: any): any => {
        let autoshape_shape_: any, freeform_: any, textbox_shape_: any;
        [autoshape_shape_, textbox_shape_, freeform_] = $list(non_placeholder_shapes_);
        $assertEqual(autoshape_shape_.shapeType, MSO_SHAPE_TYPE.AUTO_SHAPE);
        $assertEqual(textbox_shape_.shapeType, MSO_SHAPE_TYPE.TEXT_BOX);
        $assertEqual(freeform_.shapeType, MSO_SHAPE_TYPE.FREEFORM);
      } },
    { name: "but_it_raises_when_shape_type_called_on_unrecognized_shape", args: [], fn: (self: any): any => {
        let shape: any, sp: any, xml_: any;
        xml_ = "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:nvSpPr><p:cNvPr id=\"9\" name=\"Unknown Shape Type 8\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>";
        sp = parseXml(xml_);
        shape = $call(Shape, sp, null);
        $raises(NotImplementedError, () => {
          shape.shapeType;
        });
      } },
    { name: "it_knows_what_text_it_contains", args: ["text_frame_prop_", "text_frame_"], fn: (self: any, text_frame_prop_: any, text_frame_: any): any => {
        let shape: any, text: any;
        text_frame_prop_.returnValue = text_frame_;
        text_frame_.text = "foobar";
        shape = $call(Shape, null, null);
        text = shape.text;
        $assertEqual(text, "foobar");
      } },
    { name: "it_can_change_its_text", args: ["text_frame_prop_", "text_frame_"], fn: (self: any, text_frame_prop_: any, text_frame_: any): any => {
        let shape: any;
        text_frame_prop_.returnValue = text_frame_;
        shape = $call(Shape, null, null);
        shape.text = "føøbår";
        $assertEqual(text_frame_.text, "føøbår");
      } },
    { name: "it_provides_access_to_its_text_frame", args: ["text_frame_fixture"], fn: (self: any, text_frame_fixture: any): any => {
        let TextFrame_: any, shape: any, text_frame: any, text_frame_: any, txBody: any;
        [shape, text_frame_, TextFrame_, txBody] = $list(text_frame_fixture);
        text_frame = shape.textFrame;
        TextFrame_.assertCalledOnceWith(txBody, shape);
        $assert($is(text_frame, text_frame_));
      } },
    { name: "and_it_creates_a_txBody_if_needed", args: ["txBody_fixture"], fn: (self: any, txBody_fixture: any): any => {
        let expected_xml: any, shape: any, text_frame: any;
        [shape, expected_xml] = $list(txBody_fixture);
        text_frame = shape.textFrame;
        $assertEqual(text_frame._txBody.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeShape", DescribeShape);

