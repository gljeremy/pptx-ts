"""python-pptx Presentation API workloads, for comparison with bench_prs.mjs.

Usage: python bench/bench_prs.py <deck.pptx> [acceptance test_files dir]
"""
import io
import os
import sys
import time

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu

data = open(sys.argv[1], "rb").read()
files_dir = sys.argv[2] if len(sys.argv) > 2 else "test/features/test_files"
decks = [open(os.path.join(files_dir, f), "rb").read() for f in sorted(os.listdir(files_dir)) if f.endswith(".pptx")]
image = open(os.path.join(files_dir, "monty-truth.png"), "rb").read()


def best(fn, n=5):
    m = float("inf")
    for _ in range(n):
        s = time.perf_counter()
        fn()
        m = min(m, time.perf_counter() - s)
    return m * 1000


def open_save():
    Presentation(io.BytesIO(data)).save(io.BytesIO())


def build():
    prs = Presentation()
    for i in range(100):
        prs.slides.add_slide(prs.slide_layouts[i % 2])
    prs.save(io.BytesIO())


def shapes():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.turbo_add_enabled = True
    for i in range(1000):
        tb = slide.shapes.add_textbox(Emu(i * 1000), Emu(i * 1000), Emu(914400), Emu(457200))
        tb.text_frame.text = f"shape {i}"
    prs.save(io.BytesIO())


def charts():
    prs = Presentation()
    for i in range(50):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        cd = CategoryChartData()
        cd.categories = [f"Month {m}" for m in range(1, 13)]
        for s in range(3):
            cd.add_series(f"Series {s}", [(m * 7 + s * 3 + i) % 50 + 0.5 for m in range(12)])
        slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(914400), Emu(914400), Emu(6 * 914400), Emu(4 * 914400), cd)
    prs.save(io.BytesIO())


def open_save_all():
    for d in decks:
        Presentation(io.BytesIO(d)).save(io.BytesIO())


def build200():
    prs = Presentation()
    for i in range(200):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text_frame.text = f"Slide {i + 1}"
        body = slide.placeholders[1].text_frame
        body.text = "First point"
        for j in range(3):
            p = body.add_paragraph()
            p.text = f"Sub-point {j + 1}"
            p.level = 1
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(457200), Emu(5486400), Emu(1828800), Emu(457200))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x1F, 0x77, 0xB4)
        shape.text_frame.text = f"Box {i + 1}"
        slide.shapes.add_picture(io.BytesIO(image), Emu(6400800), Emu(5029200), Emu(1828800))
    prs.save(io.BytesIO())


def shapes10k():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.turbo_add_enabled = True
    for i in range(10000):
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu((i % 100) * 91440), Emu((i // 100) * 68580), Emu(91440), Emu(68580))
    prs.save(io.BytesIO())


print(f"open + save presentation  {best(open_save):8.1f} ms")
print(f"build 100-slide deck      {best(build):8.1f} ms")
print(f"1000 text boxes on a slide {best(shapes, 3):8.1f} ms")
print(f"50 charts with workbooks  {best(charts, 3):8.1f} ms")
print(f"open + save {len(decks)} test decks     {best(open_save_all, 3):8.1f} ms")
print(f"200 slides, text+shape+picture {best(build200, 3):8.1f} ms")
print(f"10000 shapes on a slide   {best(shapes10k, 3):8.1f} ms")
