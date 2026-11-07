"""python-pptx stage-1 workloads (OPC package + XML), for comparison with bench_opc.mjs.

Usage: py -I bench/bench_opc.py <deck.pptx>
"""
import io
import sys
import time

from lxml import etree
from pptx.opc.package import OpcPackage
from pptx.oxml import parse_xml

path = sys.argv[1]
data = open(path, "rb").read()


def best(fn, n=5):
    t = []
    for _ in range(n):
        s = time.perf_counter()
        fn()
        t.append((time.perf_counter() - s) * 1000)
    return min(t)


def open_save():
    pkg = OpcPackage.open(io.BytesIO(data))
    out = io.BytesIO()
    pkg.save(out)


xml_parts = []
pkg = OpcPackage.open(io.BytesIO(data))
for p in pkg.iter_parts():
    if p.partname.endswith(".xml"):
        xml_parts.append(p.blob)


def parse_serialize():
    for _ in range(20):
        for b in xml_parts:
            etree.tostring(parse_xml(b), encoding="UTF-8", standalone=True)


print(f"open + save package     {best(open_save):8.1f} ms")
print(f"parse+serialize x20     {best(parse_serialize):8.1f} ms  ({sum(map(len, xml_parts)) * 20 / 1e6:.1f} MB of XML)")
