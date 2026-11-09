"""Dump what python-pptx's chart API reports for every chart in the upstream test decks.

Usage: python scripts/crosscheck/charts_py.py <python-pptx checkout> charts_py.json
Then:  npx tsx scripts/crosscheck/charts_ts.ts charts_py.json <python-pptx checkout>

The paths read are in scripts/crosscheck/chart_props.json, shared with the TS side.
"""
import enum
import glob
import json
import os
import sys

from pptx import Presentation

PROPS = json.load(open(os.path.join(os.path.dirname(__file__), "chart_props.json")))


def norm(v):
    if isinstance(v, enum.Enum):
        return v.name
    if isinstance(v, str):
        return str(v)
    if isinstance(v, bool) or v is None:
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, (list, tuple)):
        return [norm(x) for x in v]
    return f"<{type(v).__name__}>"


class Missing(Exception):
    """The object has no such property (as opposed to one that raised AttributeError)."""


def read(obj, path):
    try:
        if path.startswith("len(") and path.endswith(")"):
            return len(read_raw(obj, path[4:-1]))
        return norm(read_raw(obj, path))
    except Missing:
        return "N/A"
    except Exception as e:  # noqa: BLE001
        return "ERR:" + type(e).__name__


def read_raw(obj, path):
    for part in path.split("."):
        name = part[:-2] if part.endswith("()") else part
        if obj is None:
            raise AttributeError(name)
        if not hasattr(type(obj), name) and name not in getattr(obj, "__dict__", {}):
            raise Missing(name)
        obj = getattr(obj, name)
        if part.endswith("()"):
            obj = list(obj())
    return obj


def props(obj, kind):
    return {p: read(obj, p) for p in PROPS[kind]}


def axis_info(chart, name):
    try:
        ax = getattr(chart, name)
    except Exception as e:  # noqa: BLE001
        return "ERR:" + type(e).__name__
    d = props(ax, "axis")
    if name == "value_axis":
        d.update(props(ax, "value_axis"))
    if ax.has_title:
        d["title"] = props(ax.axis_title, "axis_title")
    return d


def chart_info(chart):
    d = props(chart, "chart")
    if chart.has_title:
        d["title"] = props(chart.chart_title, "chart_title")
    if chart.has_legend:
        d["legend"] = props(chart.legend, "legend")
    plots = []
    for plot in chart.plots:
        p = props(plot, "plot")
        if plot.has_data_labels:
            p["data_labels"] = props(plot.data_labels, "data_labels")
        sers = []
        for s in plot.series:
            sd = props(s, "series")
            if hasattr(s, "iter_values") and type(s).__name__ in ("XySeries", "BubbleSeries"):
                sd.update(props(s, "xy_series"))
            try:
                pts = s.points
                sd["points"] = [props(pts[i], "point") for i in range(min(len(pts), 3))]
            except Exception as e:  # noqa: BLE001
                sd["points"] = "ERR:" + type(e).__name__
            sers.append(sd)
        p["series"] = sers
        plots.append(p)
    d["plots"] = plots
    d["category_axis"] = axis_info(chart, "category_axis")
    d["value_axis"] = axis_info(chart, "value_axis")
    d["series_names"] = [s.name for s in chart.series]
    return d


decks = sorted(glob.glob(sys.argv[1] + "/tests/test_files/*.pptx") + glob.glob(sys.argv[1] + "/features/steps/test_files/*.pptx"))
out = {}
for f in decks:
    key = f.replace("\\", "/").split("python-pptx/")[1]
    try:
        prs = Presentation(f)
    except Exception as e:  # noqa: BLE001
        out[key] = "ERR:" + type(e).__name__
        continue
    charts = []
    for si, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if getattr(sh, "has_chart", False) and sh.has_chart:
                try:
                    charts.append({"slide": si, "shape": sh.name, **chart_info(sh.chart)})
                except Exception as e:  # noqa: BLE001
                    charts.append({"slide": si, "shape": sh.name, "error": "ERR:" + type(e).__name__ + ": " + str(e)[:80]})
    if charts:
        out[key] = charts
json.dump(out, open(sys.argv[2], "w"), indent=1)
print(f"{len(out)} decks with charts, {sum(len(v) for v in out.values() if isinstance(v, list))} charts")
