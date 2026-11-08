"""Build the charts in chart_build.json with python-pptx and save the deck.

Usage: python scripts/crosscheck/chart_build_py.py charts_py.pptx
Then:  npx tsx scripts/crosscheck/chart_build_ts.ts charts_py.pptx
"""
import datetime
import json
import os
import sys

from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum import chart as chart_enums
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

SPEC = json.load(open(os.path.join(os.path.dirname(__file__), "chart_build.json"), encoding="utf-8"))


def label(v):
    if isinstance(v, dict) and "date" in v:
        return datetime.date.fromisoformat(v["date"])
    if isinstance(v, dict) and "datetime" in v:
        return datetime.datetime.fromisoformat(v["datetime"])
    return v


def make_data(d):
    if d["kind"] == "category":
        cd = CategoryChartData(number_format=d.get("number_format", "General"))
        cats = d["categories"]
        if isinstance(cats, dict):
            for top, subs in cats["multi"]:
                c = cd.add_category(top)
                for s in subs:
                    c.add_sub_category(s)
        else:
            cd.categories = [label(c) for c in cats]
        for s in d["series"]:
            cd.add_series(s[0], s[1], s[2] if len(s) > 2 else None)
        return cd
    cls = XyChartData if d["kind"] == "xy" else BubbleChartData
    cd = cls(number_format=d.get("number_format", "General"))
    for s in d["series"]:
        ser = cd.add_series(s[0], s[2] if len(s) > 2 else None)
        for pt in s[1]:
            ser.add_data_point(*pt)
    return cd


def value(v):
    if isinstance(v, str) and v.startswith("RGB:"):
        return RGBColor.from_string(v[4:])
    if isinstance(v, str) and v.startswith("PT:"):
        return Pt(float(v[3:]))
    if isinstance(v, str) and v.startswith("ENUM:"):
        enum_name, member = v[5:].split(".")
        return getattr(getattr(chart_enums, enum_name), member)
    return v


def apply(chart, path, *args):
    if path == "replace_data":
        chart.replace_data(make_data(SPEC["data"][args[0]]))
        return
    parts = path.split(".")
    obj = chart
    for part in parts[:-1]:
        obj = obj[int(part)] if part.isdigit() else getattr(obj, part)
    last = parts[-1]
    if last.endswith("()"):
        getattr(obj, last[:-2])()
    else:
        setattr(obj, last, value(args[0]))


prs = Presentation()
for i, c in enumerate(SPEC["charts"]):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    gf = slide.shapes.add_chart(getattr(XL_CHART_TYPE, c["type"]), Inches(1), Inches(1), Inches(6), Inches(4.5), make_data(SPEC["data"][c["data"]]))
    for op in c.get("ops", []):
        apply(gf.chart, *op)
prs.save(sys.argv[1])
print(len(SPEC["charts"]), "charts")
