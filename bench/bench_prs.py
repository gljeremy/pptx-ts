"""python-pptx stage-2 workloads (Presentation API), for comparison with bench_prs.mjs.

Usage: python bench/bench_prs.py <deck.pptx>
"""
import io
import sys
import time

from pptx import Presentation

data = open(sys.argv[1], "rb").read()


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


print(f"open + save presentation  {best(open_save):8.1f} ms")
print(f"build 100-slide deck      {best(build):8.1f} ms")
