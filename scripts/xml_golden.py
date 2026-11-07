"""Dump lxml's view of every XML member in a set of .pptx files, as JSON, so the TS XML
engine can be checked against it (test/xml_golden.test.ts).

Usage: py -I scripts/xml_golden.py out.json file.pptx [...]
"""
import json
import sys
import zipfile

from lxml import etree

parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
out = []
for path in sys.argv[2:]:
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            raw = z.read(name)
            root = etree.fromstring(raw, parser)
            first = root[0] if len(root) and isinstance(root[0].tag, str) else root
            out.append({
                "file": path.replace("\\", "/").split("/")[-1],
                "name": name,
                "raw": raw.decode("utf-8"),
                "part": etree.tostring(root, encoding="UTF-8", standalone=True).decode("utf-8"),
                "pretty": etree.tostring(root, encoding="unicode", pretty_print=True),
                "sub": etree.tostring(first, encoding="unicode", pretty_print=True),
                "ids": [str(x) for x in root.xpath("//@r:id", namespaces={"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"})],
            })
json.dump(out, open(sys.argv[1], "w", encoding="utf-8"))
print(len(out), "xml members")
